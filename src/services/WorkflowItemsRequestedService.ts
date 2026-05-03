import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/graph/users';

import { IWorkflowItemsRequestedItem, IUserRoleContext } from '../models';
import {
  DEFAULT_SHAREPOINT_FIELD_CONFIG,
  ISharePointFieldConfig,
  SHAREPOINT_FIELDS,
  getBasicSelectFields,
  getConfiguredUserFields,
  getSharePointFieldConfig,
  getUserFieldExpansions
} from '../constants';

const DEFAULT_CREATED_WITHIN_DAYS = 180;
const ROLE_ALIASES: { [role: string]: string[] } = {
  CM: ['CM', 'Credit Manager'],
  BUYER: ['BUYER', 'Buyer'],
  DSR: ['DSR', 'DORDSM', 'District Sales Representative'],
  EXECUTIVE: ['EXECUTIVE', 'Executive'],
  GM: ['GM', 'General Manager'],
  PRESIDENT: ['PRESIDENT', 'President'],
  CUSTOMERSERVICE: ['CUSTOMERSERVICE', 'Customer Service', 'CS']
};

interface ISharePointListInfo {
  Id: string;
  Title: string;
}

interface ISharePointUserValue {
  Title?: string;
  EMail?: string;
  Email?: string;
}

interface ISharePointWorkflowItem {
  [fieldName: string]: unknown;
}

export class WorkflowItemsRequestedService {
  private _sp: SPFI;

  constructor(spInstance: SPFI, _graphInstance?: GraphFI) {
    this._sp = spInstance;
  }

  public async getWorkflowItems(
    listIds: string[],
    configuredFields?: Partial<ISharePointFieldConfig>,
    createdWithinDays: number = DEFAULT_CREATED_WITHIN_DAYS,
    roleFilter?: { userId: number; roles: string[]; fieldConfig: ISharePointFieldConfig }
  ): Promise<IWorkflowItemsRequestedItem[]> {
    console.log(
      'WorkflowItemsRequestedService: Starting to load workflow items...'
    );
    console.log('Lists to search:', listIds);

    if (!listIds || listIds.length === 0) {
      console.log('WorkflowItemsRequestedService: No lists configured');
      return [];
    }

    const results: IWorkflowItemsRequestedItem[] = [];
    for (const listId of listIds) {
      try {
        console.log(
          `WorkflowItemsRequestedService: Loading items from list "${listId}"...`
        );
        const items = await this._getItemsFromList(
          listId,
          configuredFields,
          createdWithinDays,
          roleFilter
        );
        console.log(
          `WorkflowItemsRequestedService: Found ${items.length} items in list "${listId}"`
        );
        results.push(...items);
      } catch (err) {
        console.warn(
          `WorkflowItemsRequestedService: Error loading list "${listId}":`,
          err
        );
        // swallow here; the component can show partial load warnings if desired
        // console.error('List load failed', listId, err);
      }
    }
    console.log(
      `WorkflowItemsRequestedService: Total items loaded: ${results.length}`
    );
    return results;
  }

  public async getUserRoleContext(
    userRoles: string[]
  ): Promise<IUserRoleContext> {
    try {
      // Ensure userRoles is always an array
      const rolesArray = Array.isArray(userRoles)
        ? userRoles
        : userRoles
        ? [userRoles]
        : [];

      // Get current user info
      const currentUser = await this._sp.web.currentUser();
      console.log(
        'WorkflowItemsRequestedService: Current user:',
        currentUser.Title
      );

      return {
        currentUser: currentUser.Title || currentUser.Email || '',
        userRoles: rolesArray,
        displayText:
          rolesArray && rolesArray.length > 0
            ? `${'I am the'} ${rolesArray.join(', ')}`
            : ''
      };
    } catch (error) {
      console.error(
        'WorkflowItemsRequestedService: Error getting user role context:',
        error
      );
      return {
        currentUser: '',
        userRoles: Array.isArray(userRoles)
          ? userRoles
          : userRoles
          ? [userRoles]
          : [],
        displayText: ''
      };
    }
  }

  private async _getItemsFromList(
    listId: string,
    configuredFields?: Partial<ISharePointFieldConfig>,
    createdWithinDays: number = DEFAULT_CREATED_WITHIN_DAYS,
    roleFilter?: { userId: number; roles: string[]; fieldConfig: ISharePointFieldConfig }
  ): Promise<IWorkflowItemsRequestedItem[]> {
    console.log(
      `WorkflowItemsRequestedService: Loading items from list "${listId}"...`
    );
    const fieldConfig = getSharePointFieldConfig(configuredFields);

    try {
      // Check if list exists first (by title)
      const listExists = await this._sp.web.lists
        .getById(listId)
        .select(
          DEFAULT_SHAREPOINT_FIELD_CONFIG.id,
          DEFAULT_SHAREPOINT_FIELD_CONFIG.title
        )()
        .catch(() => null);
      if (!listExists) {
        console.log(
          `WorkflowItemsRequestedService: List "${listId}" does not exist, skipping`
        );
        return [];
      }

      const createdAfter = this._getCreatedAfterDate(createdWithinDays);
      const createdFilter = this._getCreatedFilter(
        fieldConfig.created,
        createdAfter
      );
      const userFields = getConfiguredUserFields(fieldConfig).filter(Boolean);
      const select: string[] = this._withoutUserFields(
        getBasicSelectFields(fieldConfig),
        userFields
      );
      const expand: string[] = [];

      // Add user field expansions
      userFields.forEach(f => {
        const expansions = getUserFieldExpansions(f, fieldConfig);
        select.push(...expansions);
        expand.push(f);
      });

      // Build the final OData filter: date filter AND optional role filter
      const roleODataFilter = roleFilter
        ? this._buildRoleODataFilter(roleFilter.userId, roleFilter.roles, roleFilter.fieldConfig)
        : '';
      const combinedFilter = roleODataFilter
        ? `(${createdFilter}) and (${roleODataFilter})`
        : createdFilter;

      let raw;
      try {
        // Try with expanded user fields first
        raw = await this._executeItemsQuery(
          listId,
          select,
          expand,
          combinedFilter
        );
        console.log(
          `Successfully queried list "${listId}" with user field expansions`
        );
        console.log(`Sample raw item from list "${listId}":`, raw[0]);
      } catch (error) {
        console.warn(
          `WorkflowItemsRequestedService: Error querying list "${listId}" with expansions:`,
          error
        );
        // Try with basic fields and raw user field values
        try {
          const basicSelect = getBasicSelectFields(fieldConfig);
          raw = await this._executeItemsQuery(
            listId,
            basicSelect,
            [],
            combinedFilter
          );
          console.log(
            `Fallback query successful for list "${listId}" without expansions`
          );
          console.log(
            `Sample fallback raw item from list "${listId}":`,
            raw[0]
          );
        } catch (fallbackError) {
          console.warn(
            `WorkflowItemsRequestedService: Fallback query also failed for list "${listId}":`,
            fallbackError
          );
          // Final fallback with minimal fields and no role filter
          const minimalSelect = [
            fieldConfig.id,
            fieldConfig.title,
            fieldConfig.created,
            fieldConfig.modified
          ];
          raw = await this._executeItemsQuery(
            listId,
            minimalSelect,
            [],
            createdFilter  // drop role filter in last-resort fallback
          );
        }
      }

      const listInfo = await this._sp.web.lists
        .getById(listId)
        .select(
          DEFAULT_SHAREPOINT_FIELD_CONFIG.id,
          DEFAULT_SHAREPOINT_FIELD_CONFIG.title
        )();
      console.log(
        `WorkflowItemsRequestedService: Successfully loaded ${raw.length} items from list "${listId}"`
      );

      return raw.map(item => this._transformItem(item, listInfo, fieldConfig));
    } catch (error) {
      console.error(
        `WorkflowItemsRequestedService: Error loading from list "${listId}":`,
        error
      );
      return [];
    }
  }

  private async _executeItemsQuery(
    listId: string,
    select: string[],
    expand: string[],
    createdFilter: string
  ): Promise<ISharePointWorkflowItem[]> {
    let query = this._sp.web.lists
      .getById(listId)
      .items.select(...this._uniqueFields(select))
      .filter(createdFilter)
      .top(5000);

    if (expand.length > 0) {
      query = query.expand(...this._uniqueFields(expand));
    }

    return query() as Promise<ISharePointWorkflowItem[]>;
  }

  /**
   * Builds an OData $filter clause that restricts results to items relevant
   * to the current user based on their selected roles.
   *
   * - Person-field roles (CM / DSR / CUSTOMERSERVICE) use  FieldName/Id eq {userId}
   *   so SharePoint evaluates the lookup on the server.
   * - Text-based roles (BUYER / EXECUTIVE / GM / PRESIDENT) match against the
   *   CurrentAssignedRole text field.
   * - All clauses are joined with  or  so items matching ANY role are returned.
   */
  public _buildRoleODataFilter(
    userId: number,
    normalizedRoles: string[],
    fieldConfig: ISharePointFieldConfig
  ): string {
    const clauses: string[] = [];

    for (const role of normalizedRoles) {
      switch (role) {
        case 'CM':
          clauses.push(`${fieldConfig.creditManager}/Id eq ${userId}`);
          break;
        case 'DSR':
          clauses.push(`${fieldConfig.dsr}/Id eq ${userId}`);
          break;
        case 'CUSTOMERSERVICE':
          clauses.push(`${fieldConfig.customerService}/Id eq ${userId}`);
          break;
        case 'BUYER':
        case 'EXECUTIVE':
        case 'GM':
        case 'PRESIDENT': {
          // Text field equality for roles that have no dedicated person column
          const aliases = ROLE_ALIASES[role] || [role];
          aliases.forEach(alias => {
            clauses.push(`${fieldConfig.currentAssignedRole} eq '${alias}'`);
          });
          break;
        }
        default:
          // Unknown role – try text field match
          clauses.push(`${fieldConfig.currentAssignedRole} eq '${role}'`);
          break;
      }
    }

    if (clauses.length === 0) return '';
    return clauses.join(' or ');
  }

  private _getCreatedAfterDate(createdWithinDays: number): Date {
    const days = Math.max(1, Math.floor(createdWithinDays || 0));
    return new Date(Date.now() - days * 24 * 60 * 60 * 1000);
  }

  private _getCreatedFilter(createdField: string, createdAfter: Date): string {
    return `${createdField} ge datetime'${createdAfter.toISOString()}'`;
  }

  private _withoutUserFields(fields: string[], userFields: string[]): string[] {
    return fields.filter(field => userFields.indexOf(field) === -1);
  }

  private _uniqueFields(fields: string[]): string[] {
    return fields.filter(
      (fieldName, index, allFields) =>
        fieldName && allFields.indexOf(fieldName) === index
    );
  }

  private _transformItem(
    item: ISharePointWorkflowItem,
    listInfo: ISharePointListInfo,
    configuredFields?: Partial<ISharePointFieldConfig>
  ): IWorkflowItemsRequestedItem {
    const fields = getSharePointFieldConfig(configuredFields);
    const getUser = (f: unknown): string => {
      if (!f) return '';
      if (typeof f === 'string') return f;
      if (typeof f !== 'object') return '';

      const user = f as ISharePointUserValue;
      if (user.Title) return user.Title;
      if (user.EMail) return user.EMail;
      if (user.Email) return user.Email;
      return '';
    };

    const workflowStatusField = fields.workflowStatus;
    const creditManagerField = fields.creditManager;
    const dsrField = fields.dsr;
    const customerServiceField = fields.customerService;
    const currentAssignedRoleField = fields.currentAssignedRole;

    // Debug logging for user fields
    console.log(`Transform item ${item[fields.id]}:`, {
      creditManager: item[creditManagerField],
      dsr: item[dsrField],
      customerService: item[customerServiceField],
      currentAssignedRole: item[currentAssignedRoleField]
    });

    const transformed = {
      id: String(item[fields.id]),
      title:
        this._valueToString(item[fields.title]) ||
        this._valueToString(item[SHAREPOINT_FIELDS.SYSTEM.TITLE]),
      workflowStatus: this._normalizeWorkflowStatus(item[workflowStatusField]),
      creditManager: getUser(item[creditManagerField]),
      dsr: getUser(item[dsrField]),
      customerService: getUser(item[customerServiceField]),
      currentAssignedRole: this._valueToString(item[currentAssignedRoleField]),
      createdDate: new Date(this._valueToString(item[fields.created])),
      modifiedDate: new Date(this._valueToString(item[fields.modified])),
      sourceListTitle: listInfo.Title,
      sourceListId: listInfo.Id,
      sourceItemId: this._valueToNumber(item[fields.id])
    };

    console.log(`Transformed item ${item[fields.id]}:`, {
      creditManager: transformed.creditManager,
      dsr: transformed.dsr,
      customerService: transformed.customerService
    });

    return transformed;
  }

  private _normalizeWorkflowStatus(
    status: unknown
  ): IWorkflowItemsRequestedItem['workflowStatus'] {
    if (!status) return 'Pending';
    const s = status.toString().toLowerCase().trim();
    if (s.indexOf('progress') !== -1 || s === 'active') return 'In Progress';
    if (s.indexOf('complete') !== -1 || s === 'done') return 'Completed';
    if (s.indexOf('reject') !== -1 || s === 'declined') return 'Rejected';
    if (s.indexOf('cancel') !== -1) return 'Cancelled';
    return 'Pending';
  }

  private _valueToString(value: unknown): string {
    if (value === undefined || (typeof value === 'object' && !value)) return '';
    return String(value);
  }

  private _valueToNumber(value: unknown): number {
    if (typeof value === 'number') return value;
    const parsed = Number(value);
    return isNaN(parsed) ? 0 : parsed;
  }

  public async getCurrentUser(): Promise<{
    displayName: string;
    email: string;
    loginName: string;
  }> {
    try {
      const u = await this._sp.web.currentUser();
      return {
        displayName: u.Title || '',
        email: u.Email || '',
        loginName: u.LoginName || ''
      };
    } catch {
      return { displayName: '', email: '', loginName: '' };
    }
  }

  /**
   * CLIENT-SIDE role-based filtering.
   *
   * How it works:
   * - All items are loaded from SharePoint first (no server-side role filter).
   * - For each item, we check whether the current user appears in the
   *   person fields that correspond to the selected roles:
   *     CM          → item.creditManager person field
   *     DSR         → item.dsr person field
   *     CUSTOMERSERVICE → item.customerService person field
   *     BUYER / EXECUTIVE / GM / PRESIDENT → item.currentAssignedRole text
   *       (these roles don't have dedicated person fields; the workflow sets the
   *        currentAssignedRole text to indicate who the next step belongs to,
   *        so we show those items when the user has that role)
   * - An item is shown if ANY selected role matches.
   */
  public async filterItemsByUserRoles(
    items: IWorkflowItemsRequestedItem[],
    userRoles: string[],
    currentUser?: string
  ): Promise<IWorkflowItemsRequestedItem[]> {
    const rolesArray = Array.isArray(userRoles)
      ? userRoles
      : userRoles
      ? [userRoles]
      : [];

    if (!rolesArray || rolesArray.length === 0) return items;

    // Resolve the current user's display name and email
    let currentUserName = currentUser || '';
    let currentUserEmail = '';
    if (!currentUserName) {
      const u = await this.getCurrentUser();
      currentUserName = u.displayName;
      currentUserEmail = u.email;
    }

    const userNameLc = currentUserName.toLowerCase();
    const userEmailLc = currentUserEmail.toLowerCase();

    // Normalise selected roles to canonical keys (CM, DSR, CUSTOMERSERVICE, …)
    const normalizedRoles = this._normalizeSelectedRoles(rolesArray);
    if (normalizedRoles.length === 0) return items;

    console.log('Role filter: current user =', currentUserName, '| roles =', normalizedRoles);

    return items.filter(item => {
      return normalizedRoles.some(role => this._itemMatchesRoleForUser(item, role, userNameLc, userEmailLc));
    });
  }

  /**
   * Returns true if the item should be visible for the current user given the specified role.
   */
  private _itemMatchesRoleForUser(
    item: IWorkflowItemsRequestedItem,
    role: string,
    userNameLc: string,
    userEmailLc: string
  ): boolean {
    const matchesUser = (fieldValue: string): boolean => {
      if (!fieldValue) return false;
      const fLc = fieldValue.toLowerCase();
      if (userNameLc && fLc.indexOf(userNameLc) !== -1) return true;
      if (userEmailLc && fLc.indexOf(userEmailLc) !== -1) return true;
      return false;
    };

    switch (role) {
      case 'CM':
        // Show if current user is in the Credit Manager person field
        return matchesUser(item.creditManager || '');

      case 'DSR':
        // Show if current user is in the DSR person field
        return matchesUser(item.dsr || '');

      case 'CUSTOMERSERVICE':
        // Show if current user is in the Customer Service person field
        return matchesUser(item.customerService || '');

      case 'BUYER':
      case 'EXECUTIVE':
      case 'GM':
      case 'PRESIDENT':
        // These roles have no dedicated person field.
        // Show items where currentAssignedRole text equals this role,
        // meaning the workflow is currently at a step this user is responsible for.
        return this._textMatchesRole(item.currentAssignedRole || '', role);

      default:
        // Fallback: check all person fields and the role text field
        return (
          matchesUser(item.creditManager || '') ||
          matchesUser(item.dsr || '') ||
          matchesUser(item.customerService || '') ||
          this._textMatchesRole(item.currentAssignedRole || '', role)
        );
    }
  }

  private _normalizeSelectedRoles(userRoles: string[]): string[] {
    const normalizedRoles: string[] = [];
    userRoles.forEach(role => {
      const normalized = this._normalizeRoleName(role);
      if (normalized && normalizedRoles.indexOf(normalized) === -1) {
        normalizedRoles.push(normalized);
      }
    });
    return normalizedRoles;
  }

  private _normalizeRoleName(role: string): string {
    const normalized = this._normalizeRoleToken(role);
    const knownRoles = Object.keys(ROLE_ALIASES);
    for (let i = 0; i < knownRoles.length; i++) {
      const knownRole = knownRoles[i];
      const aliases = ROLE_ALIASES[knownRole];
      for (let j = 0; j < aliases.length; j++) {
        if (this._normalizeRoleToken(aliases[j]) === normalized) {
          return knownRole;
        }
      }
    }
    return normalized;
  }

  private _textMatchesRole(text: string, role: string): boolean {
    const normalizedText = this._normalizeRoleToken(text);
    if (!normalizedText) return false;

    const aliases = ROLE_ALIASES[role] || [role];
    for (let i = 0; i < aliases.length; i++) {
      const alias = this._normalizeRoleToken(aliases[i]);
      if (
        normalizedText === alias ||
        normalizedText.indexOf(alias) !== -1 ||
        alias.indexOf(normalizedText) !== -1
      ) {
        return true;
      }
    }
    return false;
  }

  private _normalizeRoleToken(value: string): string {
    return (value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

}
