import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/site-users/web';
import '@pnp/graph/users';

import { IWorkflowItemsRequestedItem, IUserRoleContext } from '../models';
import {
  SHAREPOINT_FIELDS,
  USER_FIELD_MAPPINGS,
  USER_FIELDS,
  BASIC_SELECT_FIELDS,
  getUserFieldExpansions
} from '../constants';

export class WorkflowItemsRequestedService {
  private _sp: SPFI;

  constructor(spInstance: SPFI, _graphInstance?: GraphFI) {
    this._sp = spInstance;
  }

  public async getWorkflowItems(
    listIds: string[]
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
        const items = await this._getItemsFromList(listId);
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
    listId: string
  ): Promise<IWorkflowItemsRequestedItem[]> {
    console.log(
      `WorkflowItemsRequestedService: Loading items from list "${listId}"...`
    );

    try {
      // Check if list exists first (by title)
      const listExists = await this._sp.web.lists
        .getById(listId)
        .select(SHAREPOINT_FIELDS.SYSTEM.ID, SHAREPOINT_FIELDS.SYSTEM.TITLE)()
        .catch(() => null);
      if (!listExists) {
        console.log(
          `WorkflowItemsRequestedService: List "${listId}" does not exist, skipping`
        );
        return [];
      }

      // Get list fields to check what's available
      const fields = await this._sp.web.lists
        .getById(listId)
        .fields.select('InternalName', 'Title', 'TypeAsString')();
      console.log(
        `Available fields in list "${listId}":`,
        fields.map(f => ({
          InternalName: f.InternalName,
          Title: f.Title,
          Type: f.TypeAsString
        }))
      );

      // Check if user fields exist in the list using predefined mappings
      const userFieldMappings = USER_FIELD_MAPPINGS;

      const availableUserFields = userFieldMappings.filter(mapping =>
        fields.some(f => f.InternalName === mapping.internal)
      );

      console.log(
        `Available user fields in list "${listId}":`,
        availableUserFields
      );

      const select: string[] = [...BASIC_SELECT_FIELDS];

      const expand: string[] = [];
      const userFields = USER_FIELDS.filter(Boolean) as string[];

      // Add user field expansions
      userFields.forEach(f => {
        const expansions = getUserFieldExpansions(f);
        select.push(...expansions);
        expand.push(f);
      });

      let raw;
      try {
        // Try with expanded user fields first
        let query = this._sp.web.lists.getById(listId).items.select(...select);
        if (expand.length > 0) {
          query = query.expand(...expand);
        }
        raw = await query();
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
          const basicSelect = [...BASIC_SELECT_FIELDS];
          raw = await this._sp.web.lists
            .getById(listId)
            .items.select(...basicSelect)();
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
          // Final fallback with minimal fields
          raw = await this._sp.web.lists
            .getById(listId)
            .items.select(
              SHAREPOINT_FIELDS.SYSTEM.ID,
              SHAREPOINT_FIELDS.SYSTEM.TITLE,
              SHAREPOINT_FIELDS.SYSTEM.CREATED,
              SHAREPOINT_FIELDS.SYSTEM.MODIFIED
            )();
        }
      }

      const listInfo = await this._sp.web.lists
        .getById(listId)
        .select(SHAREPOINT_FIELDS.SYSTEM.ID, SHAREPOINT_FIELDS.SYSTEM.TITLE)();
      console.log(
        `WorkflowItemsRequestedService: Successfully loaded ${raw.length} items from list "${listId}"`
      );

      return raw.map(item => this._transformItem(item, listInfo));
    } catch (error) {
      console.error(
        `WorkflowItemsRequestedService: Error loading from list "${listId}":`,
        error
      );
      return [];
    }
  }

  private _transformItem(
    item: any,
    listInfo: { Id: string; Title: string }
  ): IWorkflowItemsRequestedItem {
    const getUser = (f: any): string => {
      if (!f) return '';
      if (typeof f === 'string') return f;
      if (f.Title) return f.Title;
      if (f.EMail) return f.EMail;
      if (f.Email) return f.Email;
      return '';
    };

    const titleField = SHAREPOINT_FIELDS.BASIC.SIMPLE_TITLE;
    const workflowStatusField = SHAREPOINT_FIELDS.BASIC.WORKFLOW_STATUS;
    const creditManagerField = SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER;
    const dsrField = SHAREPOINT_FIELDS.USERS.DSR;
    const customerServiceField = SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE;
    const currentAssignedRoleField =
      SHAREPOINT_FIELDS.BASIC.CURRENT_ASSIGNED_ROLE;

    // Debug logging for user fields
    console.log(`Transform item ${item[SHAREPOINT_FIELDS.SYSTEM.ID]}:`, {
      creditManager: item[creditManagerField],
      dsr: item[dsrField],
      customerService: item[customerServiceField],
      currentAssignedRole: item[currentAssignedRoleField]
    });

    const transformed = {
      id: String(item[SHAREPOINT_FIELDS.SYSTEM.ID]),
      title:
        item[titleField] ||
        item[SHAREPOINT_FIELDS.BASIC.SIMPLE_TITLE] ||
        item[SHAREPOINT_FIELDS.SYSTEM.TITLE] ||
        '',
      workflowStatus: this._normalizeWorkflowStatus(item[workflowStatusField]),
      creditManager: getUser(item[creditManagerField]),
      dsr: getUser(item[dsrField]),
      customerService: getUser(item[customerServiceField]),
      currentAssignedRole: item[currentAssignedRoleField] || '',
      createdDate: new Date(item[SHAREPOINT_FIELDS.SYSTEM.CREATED]),
      modifiedDate: new Date(item[SHAREPOINT_FIELDS.SYSTEM.MODIFIED]),
      sourceListTitle: listInfo.Title,
      sourceListId: listInfo.Id,
      sourceItemId: item[SHAREPOINT_FIELDS.SYSTEM.ID]
    };

    console.log(`Transformed item ${item[SHAREPOINT_FIELDS.SYSTEM.ID]}:`, {
      creditManager: transformed.creditManager,
      dsr: transformed.dsr,
      customerService: transformed.customerService
    });

    return transformed;
  }

  private _normalizeWorkflowStatus(
    status: any
  ): IWorkflowItemsRequestedItem['workflowStatus'] {
    if (!status) return 'Pending';
    const s = status.toString().toLowerCase().trim();
    if (s.indexOf('progress') !== -1 || s === 'active') return 'In Progress';
    if (s.indexOf('complete') !== -1 || s === 'done') return 'Completed';
    if (s.indexOf('reject') !== -1 || s === 'declined') return 'Rejected';
    if (s.indexOf('cancel') !== -1) return 'Cancelled';
    return 'Pending';
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

  public async filterItemsByUserRoles(
    items: IWorkflowItemsRequestedItem[],
    userRoles: string[],
    currentUser?: string
  ): Promise<IWorkflowItemsRequestedItem[]> {
    // Ensure userRoles is always an array
    const rolesArray = Array.isArray(userRoles)
      ? userRoles
      : userRoles
      ? [userRoles]
      : [];

    if (!rolesArray || rolesArray.length === 0) return items;
    if (!currentUser) {
      const u = await this.getCurrentUser();
      currentUser = u.displayName || u.email;
    }
    const userLc = (currentUser || '').toLowerCase();
    return items.filter(it => {
      const fields = [
        it.creditManager,
        it.dsr,
        it.customerService,
        it.currentAssignedRole
      ]
        .filter(f => f)
        .map(f => f!.toLowerCase());

      if (userLc && fields.some(f => f.indexOf(userLc) !== -1)) return true;
      return rolesArray.some(r =>
        fields.some(f => f.indexOf((r || '').toLowerCase()) !== -1)
      );
    });
  }
}
