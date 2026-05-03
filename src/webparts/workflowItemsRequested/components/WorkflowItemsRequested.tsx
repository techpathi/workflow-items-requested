import * as React from 'react';
import styles from './WorkflowItemsRequested.module.scss';
import type { IWorkflowItemsRequestedProps } from './IWorkflowItemsRequestedProps';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  IconButton,
  IGroup,
  SearchBox,
  Callout,
  Checkbox,
  DirectionalHint,
  PrimaryButton,
  DefaultButton
} from '@fluentui/react';
import type { IDetailsColumnProps } from '@fluentui/react';
import { FilterOverlay, IFilterConfig } from '../../../components';
import { useServices } from '../../../contexts/ServiceContext';
import { WorkflowItemsRequestedService } from '../../../services';
import { FilterUtils, IColumnFilter } from '../../../utils';
import { IWorkflowItemsRequestedItem, IUserRoleContext } from '../../../models';

// Default fallback roles (used if no role definitions are provided)
const BUILTIN_ROLES: { key: string; label: string }[] = [
  { key: 'CM', label: 'CM (Credit Manager)' },
  { key: 'DSR', label: 'DSR' },
  { key: 'CUSTOMERSERVICE', label: 'Customer Service' },
  { key: 'BUYER', label: 'Buyer' },
  { key: 'EXECUTIVE', label: 'Executive' },
  { key: 'GM', label: 'GM' },
  { key: 'PRESIDENT', label: 'President' }
];

/** Returns the list of roles available for the user to choose from. */
function getAvailableRoles(roleDefinitions?: { key: string; text: string }[]): { key: string; label: string }[] {
  if (!roleDefinitions || roleDefinitions.length === 0) {
    return BUILTIN_ROLES;
  }
  return roleDefinitions.map(r => ({ key: r.key, label: r.text }));
}

// Interface for component state
interface IWorkflowItemsRequestedState {
  workflowItems: IWorkflowItemsRequestedItem[];
  filteredItems: IWorkflowItemsRequestedItem[];
  groupedItems: IGroup[];
  columns: IColumn[];
  isLoading: boolean;
  error?: string;
  warnings?: string[];
  userRoleContext?: IUserRoleContext;

  // Session-overridden roles (undefined = use props.userRoles)
  activeUserRoles: string[] | undefined;
  // Role edit callout state
  isRoleEditOpen: boolean;
  roleEditDraft: string[]; // roles being edited in the callout before Apply
  roleEditTargetId: string;

  // Filter state
  filters: IColumnFilter;
  activeFilterColumn?: string;
  filterTargetRef?: React.RefObject<HTMLElement>;
  searchText: string;
}

interface IWorkflowItemsRequestedPropsWithService
  extends IWorkflowItemsRequestedProps {
  workflowService?: WorkflowItemsRequestedService;
}

class WorkflowItemsRequestedClass extends React.Component<
  IWorkflowItemsRequestedPropsWithService,
  IWorkflowItemsRequestedState
> {
  private _selection: Selection;
  private _workflowService?: WorkflowItemsRequestedService;

  constructor(props: IWorkflowItemsRequestedPropsWithService) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => {
        // Handle selection changes if needed
      }
    });

    this._workflowService = props.workflowService;

    this.state = {
      workflowItems: [],
      filteredItems: [],
      groupedItems: [],
      columns: this._buildColumns(),
      isLoading: true,
      error: undefined,
      warnings: [],
      userRoleContext: undefined,
      activeUserRoles: undefined,
      isRoleEditOpen: false,
      roleEditDraft: props.userRoles || [],
      roleEditTargetId: 'role-edit-anchor',
      filters: {},
      activeFilterColumn: undefined,
      filterTargetRef: undefined,
      searchText: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadData();
  }

  public async componentDidUpdate(
    prevProps: IWorkflowItemsRequestedPropsWithService
  ): Promise<void> {
    // Reload data if workflow lists changed
    if (
      JSON.stringify(prevProps.workflowLists) !==
        JSON.stringify(this.props.workflowLists) ||
      JSON.stringify(prevProps.sharePointFields) !==
        JSON.stringify(this.props.sharePointFields) ||
      prevProps.createdWithinDays !== this.props.createdWithinDays
    ) {
      await this._loadData();
    }

    // Reload if userRoles (array contents) changed or role-based filtering toggle changed
    if (
      JSON.stringify(prevProps.userRoles || []) !==
        JSON.stringify(this.props.userRoles || []) ||
      prevProps.enableRoleBasedFiltering !== this.props.enableRoleBasedFiltering
    ) {
      await this._loadData();
    }

    // Rebuild columns if any visibility / sorting related props changed
    if (
      prevProps.showWorkflowStatus !== this.props.showWorkflowStatus ||
      prevProps.showCreditManager !== this.props.showCreditManager ||
      prevProps.showDSR !== this.props.showDSR ||
      prevProps.showCustomerService !== this.props.showCustomerService ||
      prevProps.showCurrentAssignedRole !==
        this.props.showCurrentAssignedRole ||
      prevProps.showCreatedDate !== this.props.showCreatedDate ||
      prevProps.showSourceList !== this.props.showSourceList ||
      prevProps.defaultSortColumn !== this.props.defaultSortColumn ||
      prevProps.defaultSortDirection !== this.props.defaultSortDirection
    ) {
      // Only rebuild columns (don't reload data)
      // Preserve existing filters & items
      // eslint-disable-next-line react/no-did-update-set-state
      this.setState({ columns: this._buildColumns() });
    }
  }

  private async _loadData(): Promise<void> {
    this.setState({ isLoading: true, error: undefined });

    try {
      if (!this._workflowService) {
        throw new Error('Workflow service not initialized');
      }

      if (!this.props.workflowLists || this.props.workflowLists.length === 0) {
        this.setState({
          workflowItems: [],
          filteredItems: [],
          groupedItems: [],
          isLoading: false,
          error:
            'No workflow lists configured. Please configure lists in the web part properties.'
        });
        return;
      }

      console.log(
        'Loading workflow items from lists:',
        this.props.workflowLists
      );

      // --- SERVER-SIDE ROLE FILTERING ---
      // activeUserRoles (session override) takes priority over props.userRoles.
      // This allows end-users to change their active roles without touching the property pane.
      const effectiveRoles =
        this.state.activeUserRoles !== undefined
          ? this.state.activeUserRoles
          : this.props.userRoles || [];

      let roleFilter:
        | {
            userId: number;
            roles: string[];
            fieldConfig: typeof this.props.sharePointFields;
          }
        | undefined = undefined;

      if (
        this.props.enableRoleBasedFiltering &&
        effectiveRoles.length > 0
      ) {
        const spUser = await this._workflowService['_sp'].web.currentUser();
        const userId: number = spUser.Id;
        const normalizedRoles = effectiveRoles
          .map(r => r.toUpperCase().replace(/[^A-Z0-9]/g, ''))
          .filter(r => r.length > 0);

        console.log(
          `Server-side role filter: userId=${userId} | roles=[${normalizedRoles.join(', ')}]`
        );
        roleFilter = {
          userId,
          roles: normalizedRoles,
          fieldConfig: this.props.sharePointFields
        };
      }

      // Load workflow items – SharePoint applies the role filter server-side
      const items = await this._workflowService.getWorkflowItems(
        this.props.workflowLists,
        this.props.sharePointFields,
        this.props.createdWithinDays,
        roleFilter
      );

      console.log(
        `Loaded ${items.length} workflow items${roleFilter ? ' (server-side role filter applied)' : ''}`
      );

      // Get user role context – use effectiveRoles so "I am the..." reflects session overrides
      const userRoleContext = await this._workflowService.getUserRoleContext(
        effectiveRoles
      );

      // filteredItems = items returned from SharePoint (already filtered server-side)
      const filteredItems = items;

      // Initialize filters for all columns
      const filterableColumns = [
        'title',
        'workflowStatus',
        'creditManager',
        'dsr',
        'customerService',
        'currentAssignedRole',
        'createdDate',
        'sourceListTitle'
      ];
      const filters = FilterUtils.initializeFilters(
        filteredItems,
        filterableColumns
      );

      // Apply column filters
      const columnFilteredItems = FilterUtils.applyFilters(
        filteredItems,
        filters
      );
      const sortedItems = this._applyCurrentSorting(columnFilteredItems);

      const groupedView = this._createGroupedView(sortedItems);

      this.setState({
        workflowItems: filteredItems,
        filteredItems: groupedView.items,
        groupedItems: groupedView.groups,
        userRoleContext,
        filters,
        isLoading: false
      });
    } catch (error) {
      console.error('Error loading workflow items:', error);
      this.setState({
        isLoading: false,
        error: error instanceof Error ? error.message : 'Unknown error occurred'
      });
    }
  }

  private _buildColumns(): IColumn[] {
    // Fluent UI justified mode distributes available horizontal space proportionally
    // based on each column's maxWidth. Title gets a large maxWidth so it absorbs
    // most of the spare space; secondary columns have smaller, content-appropriate
    // maxWidths so they stay compact and readable.
    //
    // minWidth = the column will never shrink below this (scrollbar appears instead)
    // maxWidth = how much of extra space this column can claim (higher = more space)

    const columns: IColumn[] = [];

    // Title – dominant column; absorbs most available width
    columns.push({
      key: 'title',
      name: 'Title',
      fieldName: 'title',
      minWidth: 180,
      maxWidth: 800,
      isResizable: true,
      isSorted: this.props.defaultSortColumn === 'title',
      isSortedDescending: this.props.defaultSortDirection === 'desc',
      onRender: (item: IWorkflowItemsRequestedItem) => (
        <span className={styles.titleCell} title={item.title}>
          {item.title}
        </span>
      ),
      onColumnClick: this._onColumnClick
    });

    // Workflow Status – short values ("In Progress", "Completed")
    if (this.props.showWorkflowStatus) {
      columns.push({
        key: 'workflowStatus',
        name: 'Workflow Status',
        fieldName: 'workflowStatus',
        minWidth: 110,
        maxWidth: 140,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <span
            className={styles.wrapCell}
            style={{ fontWeight: 600, color: this._getStatusColor(item.workflowStatus) }}
            title={item.workflowStatus}
          >
            {item.workflowStatus}
          </span>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Credit Manager – person name; header is 14 chars
    if (this.props.showCreditManager) {
      columns.push({
        key: 'creditManager',
        name: 'Credit Manager',
        fieldName: 'creditManager',
        minWidth: 120,
        maxWidth: 160,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <span className={styles.wrapCell} title={item.creditManager || '-'}>
            {item.creditManager || '-'}
          </span>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // DSR – person name; header is 3 chars
    if (this.props.showDSR) {
      columns.push({
        key: 'dsr',
        name: 'DSR',
        fieldName: 'dsr',
        minWidth: 110,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <span className={styles.wrapCell} title={item.dsr || '-'}>
            {item.dsr || '-'}
          </span>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Customer Service – person name; header is 16 chars
    if (this.props.showCustomerService) {
      columns.push({
        key: 'customerService',
        name: 'Customer Service',
        fieldName: 'customerService',
        minWidth: 130,
        maxWidth: 170,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <span className={styles.wrapCell} title={item.customerService || '-'}>
            {item.customerService || '-'}
          </span>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Current Assigned Role – longest header (21 chars); needs most room
    if (this.props.showCurrentAssignedRole) {
      columns.push({
        key: 'currentAssignedRole',
        name: 'Current Assigned Role',
        fieldName: 'currentAssignedRole',
        minWidth: 150,
        maxWidth: 180,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <span className={styles.wrapCell} title={item.currentAssignedRole || '-'}>
            {item.currentAssignedRole || '-'}
          </span>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Created Date – fixed-format date string
    if (this.props.showCreatedDate) {
      columns.push({
        key: 'createdDate',
        name: 'Created Date',
        fieldName: 'createdDate',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <span className={styles.wrapCell} title={item.createdDate.toLocaleDateString()}>
            {item.createdDate.toLocaleDateString()}
          </span>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Source List – list name
    if (this.props.showSourceList) {
      columns.push({
        key: 'sourceListTitle',
        name: 'Source List',
        fieldName: 'sourceListTitle',
        minWidth: 110,
        maxWidth: 160,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <span className={styles.wrapCell} title={item.sourceListTitle || 'Unknown'}>
            {item.sourceListTitle || 'Unknown'}
          </span>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Add filter buttons to columns using onRenderHeader (correct prop for DetailsList)
    return columns.map(column => ({
      ...column,
      onRenderHeader: (
        props?: IDetailsColumnProps,
        defaultRender?: (p?: IDetailsColumnProps) => React.ReactNode
      ) => {
        const currentFilters = this.state.filters || {};
        const cf = currentFilters[column.key!];
        const hasActiveFilter = !!(
          cf &&
          ((cf.selectedValues && cf.selectedValues.length > 0) ||
            (cf.textValue && cf.textValue.trim().length > 0) ||
            (cf.dateRange && (cf.dateRange.from || cf.dateRange.to)))
        );

        // We intentionally do NOT call defaultRender because it already outputs the column name
        // inside its own button, which was causing duplicated accessible text like
        // "Title Filter Title". Instead we replicate only the label portion and keep the
        // filter button outside of the label so the column header's accessible name remains
        // just the column name.
        return (
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              width: '100%',
              height: '100%',
              padding: '0 4px',
              gap: 4
            }}
            onClick={(ev: React.MouseEvent<HTMLElement>) => {
              if (props && props.onColumnClick) {
                props.onColumnClick(ev, props.column);
              }
            }}
          >
            <span
              className={styles.headerTextTruncate}
              title={column.name}
              role='button'
              aria-label={column.name}
              // Keep tab focus behavior consistent with default header: allow focus via parent button semantics
              onClick={(ev: React.MouseEvent<HTMLElement>) => {
                // Sort trigger
                if (props && props.onColumnClick) {
                  props.onColumnClick(ev, props.column);
                }
              }}
              style={{ flex: '1 1 auto', minWidth: 0 }}
            >
              {column.name}
            </span>
            <IconButton
              iconProps={{
                iconName: hasActiveFilter ? 'FilterSolid' : 'Filter'
              }}
              title={
                hasActiveFilter
                  ? `Active filter on ${column.name}`
                  : `Filter ${column.name}`
              }
              aria-label={`Filter ${column.name}`}
              onClick={(ev: React.MouseEvent<HTMLElement>) => {
                ev.stopPropagation();
                ev.preventDefault();
                this._onFilterColumnClick(column.key!, ev);
              }}
              styles={{
                root: {
                  width: 24,
                  height: 24,
                  background: hasActiveFilter ? '#deecf9' : 'transparent',
                  borderRadius: 4,
                  padding: 0,
                  flex: '0 0 auto'
                },
                rootHovered: {
                  background: hasActiveFilter ? '#c7e0f4' : '#f3f2f1'
                },
                icon: {
                  fontSize: 12,
                  color: hasActiveFilter ? '#0078d4' : '#605e5c'
                }
              }}
            />
          </div>
        );
      }
    }));
  }

  private _getStatusColor(status: string): string {
    switch (status) {
      case 'Completed':
        return '#107c10'; // Green
      case 'In Progress':
        return '#0078d4'; // Blue
      case 'Pending':
        return '#ff8c00'; // Orange
      case 'Rejected':
        return '#d13438'; // Red
      case 'Cancelled':
        return '#605e5c'; // Gray
      default:
        return '#323130'; // Default
    }
  }

  private _createGroupedView(items: IWorkflowItemsRequestedItem[]): {
    items: IWorkflowItemsRequestedItem[];
    groups: IGroup[];
  } {
    if (!this.props.enableGrouping) {
      return { items, groups: [] };
    }

    // Group by source list title
    const groupMap = new Map<string, IWorkflowItemsRequestedItem[]>();

    items.forEach(item => {
      const groupKey = item.sourceListTitle || 'Unknown List';
      if (!groupMap.has(groupKey)) {
        groupMap.set(groupKey, []);
      }
      groupMap.get(groupKey)!.push(item);
    });

    const groupedItems: IWorkflowItemsRequestedItem[] = [];
    const groups: IGroup[] = [];

    groupMap.forEach((groupItems, groupKey) => {
      // IMPORTANT: Do not include the count in the name; DetailsList GroupHeader already appends count.
      groups.push({
        key: groupKey,
        name: groupKey,
        startIndex: groupedItems.length,
        count: groupItems.length,
        level: 0,
        isCollapsed: false
      });
      groupedItems.push(...groupItems);
    });

    return { items: groupedItems, groups };
  }

  private _onColumnClick = (
    ev?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ): void => {
    if (!column) return;

    const newColumns = this.state.columns.slice();
    const currColumn = newColumns.filter(
      currCol => column.key === currCol.key
    )[0];

    newColumns.forEach(newCol => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    // Sort the items
    const sortedItems = this._sortItems(
      this.state.filteredItems,
      column.key!,
      currColumn.isSortedDescending!
    );
    const groupedView = this._createGroupedView(sortedItems);

    this.setState({
      columns: newColumns,
      filteredItems: groupedView.items,
      groupedItems: groupedView.groups
    });
  };

  private _sortItems(
    items: IWorkflowItemsRequestedItem[],
    columnKey: string,
    isSortedDescending: boolean
  ): IWorkflowItemsRequestedItem[] {
    return items.slice().sort((a, b) => {
      const aValue = this._getSortValue(a, columnKey);
      const bValue = this._getSortValue(b, columnKey);

      if (aValue < bValue) {
        return isSortedDescending ? 1 : -1;
      }
      if (aValue > bValue) {
        return isSortedDescending ? -1 : 1;
      }
      return 0;
    });
  }

  private _getSortValue(
    item: IWorkflowItemsRequestedItem,
    columnKey: string
  ): string | number {
    switch (columnKey) {
      case 'title':
        return item.title || '';
      case 'workflowStatus':
        return item.workflowStatus || '';
      case 'creditManager':
        return item.creditManager || '';
      case 'dsr':
        return item.dsr || '';
      case 'customerService':
        return item.customerService || '';
      case 'currentAssignedRole':
        return item.currentAssignedRole || '';
      case 'createdDate':
        return item.createdDate ? item.createdDate.getTime() : 0;
      default:
        return '';
    }
  }

  private _getColumnName(columns: IColumn[], columnKey: string): string {
    for (let i = 0; i < columns.length; i++) {
      if (columns[i].key === columnKey) {
        return columns[i].name;
      }
    }
    return columnKey;
  }

  private _onFilterColumnClick = (
    columnKey: string,
    event: React.MouseEvent<HTMLElement>
  ): void => {
    event.stopPropagation();
    const targetRef: React.RefObject<HTMLElement> = {
      current: event.currentTarget
    };

    // Migration: if title filter previously stored as text, upgrade to checkbox config
    if (columnKey === 'title') {
      const existing = this.state.filters.title;
      if (existing && existing.type === 'text') {
        const upgraded = FilterUtils.getColumnFilterConfig(
          'title',
          this.state.workflowItems
        );
        this.setState(prev => ({
          filters: { ...prev.filters, title: upgraded }
        }));
      }
    }
    // Cascading filters: rebuild the value list for this column from the CURRENT filtered + searched items
    // excluding this column's own filter (so user can adjust it with full remaining domain)
    this.setState(prev => {
      const existingFilter = prev.filters[columnKey];
      let baseItems = prev.workflowItems;
      // Apply all other active filters first
      const otherFilters: IColumnFilter = {};
      Object.keys(prev.filters).forEach(k => {
        if (k !== columnKey) {
          otherFilters[k] = prev.filters[k];
        }
      });
      baseItems = FilterUtils.applyFilters(baseItems, otherFilters);
      // Apply current search text as well to keep consistency with visible list
      baseItems = this._applySearch([...baseItems], prev.searchText);

      // Rebuild checkbox value list if this column supports discrete values
      let updatedFilter =
        existingFilter ||
        FilterUtils.getColumnFilterConfig(columnKey, baseItems);
      if (existingFilter && existingFilter.type === 'checkbox') {
        const newValues = FilterUtils.extractColumnValues(baseItems, columnKey);
        // Preserve any selectedValues that might not appear (edge case)
        const selectedValues = existingFilter.selectedValues || [];
        const valueKeys = new Set(newValues.map(v => v.key));
        selectedValues.forEach(sel => {
          if (!valueKeys.has(sel)) {
            newValues.push({ key: sel, text: sel, count: 0 });
          }
        });
        updatedFilter = { ...existingFilter, values: newValues };
      }

      return {
        filters: { ...prev.filters, [columnKey]: updatedFilter },
        activeFilterColumn: columnKey,
        filterTargetRef: targetRef
      };
    });
  };

  private _onFilterDismiss = (): void => {
    this.setState({
      activeFilterColumn: undefined,
      filterTargetRef: undefined
    });
  };

  private _onFilterApply = (
    columnKey: string,
    config: IFilterConfig
  ): void => {
    const newFilters = FilterUtils.updateColumnFilter(
      this.state.filters,
      columnKey,
      config
    );

    let filteredItems = FilterUtils.applyFilters(
      this.state.workflowItems,
      newFilters
    );
    filteredItems = this._applySearch(filteredItems, this.state.searchText);
    filteredItems = this._applyCurrentSorting(filteredItems);
    const groupedView = this._createGroupedView(filteredItems);

    this.setState({
      filters: newFilters,
      filteredItems: groupedView.items,
      groupedItems: groupedView.groups
    });
  };

  private _onFilterClear = (columnKey: string): void => {
    const reset = FilterUtils.getColumnFilterConfig(
      columnKey,
      this.state.workflowItems
    );
    this._onFilterApply(columnKey, reset);
  };

  private _applySearch(
    items: IWorkflowItemsRequestedItem[],
    search: string
  ): IWorkflowItemsRequestedItem[] {
    const s = (search || '').trim().toLowerCase();
    if (!s) return items;
    return items.filter(it => {
      const hay = [
        it.title,
        it.workflowStatus,
        it.creditManager || '',
        it.dsr || '',
        it.customerService || '',
        it.currentAssignedRole || '',
        it.sourceListTitle || ''
      ]
        .join(' ')
        .toLowerCase();
      return hay.indexOf(s) !== -1;
    });
  }

  private _applyCurrentSorting(
    items: IWorkflowItemsRequestedItem[]
  ): IWorkflowItemsRequestedItem[] {
    let sortedColumn: IColumn | undefined = undefined;
    for (let i = 0; i < this.state.columns.length; i++) {
      const c: IColumn = this.state.columns[i];
      if (c.isSorted) {
        sortedColumn = c;
        break;
      }
    }
    if (!sortedColumn) return items;
    return this._sortItems(
      items,
      sortedColumn.key!,
      !!sortedColumn.isSortedDescending
    );
  }

  // Legacy placeholder retained for potential future command items (now unused)
  // private _getCommandBarItems(): ICommandBarItemProps[] { return []; }

  /** Refresh data (used by refresh button in unified control bar) */
  private _onRefresh = (): void => {
    this._loadData().catch(error => {
      console.error('Error refreshing workflow items:', error);
    });
  };

  /** Clear all active filters and search */
  private _onClearAll = (): void => {
    const clearedFilters = FilterUtils.clearAllFilters(this.state.filters);
    const filteredItems = FilterUtils.applyFilters(
      this.state.workflowItems,
      clearedFilters
    );
    const searchText = '';
    const afterSearch = this._applySearch(filteredItems, searchText);
    const sorted = this._applyCurrentSorting(afterSearch);
    const groupedView = this._createGroupedView(sorted);
    this.setState({
      filters: clearedFilters,
      filteredItems: groupedView.items,
      groupedItems: groupedView.groups,
      searchText
    });
  };

  public render(): React.ReactElement<IWorkflowItemsRequestedProps> {
    const {
      filteredItems,
      groupedItems,
      columns,
      isLoading,
      error,
      userRoleContext,
      activeFilterColumn,
      filterTargetRef,
      filters
    } = this.state;

    if (error) {
      return (
        <div className={styles.workflowItemsRequested}>
          <MessageBar messageBarType={MessageBarType.error}>
            Error: {error}
          </MessageBar>
        </div>
      );
    }

    return (
      <div className={styles.workflowItemsRequested}>
        <div className={styles.webpartHeader}>
          <Text
            variant='xxLarge'
            styles={{
              root: {
                fontWeight: 600,
                marginBottom: 8,
                color: '#323130'
              }
            }}
          >
            {this.props.webpartTitle || 'Workflow Items Requested'}
          </Text>
        </div>

        {/* User Context Header */}
        {this.props.showUserContext && userRoleContext && userRoleContext.displayText && (
          <div className={styles.userContext}>
            <div className={styles.userContextCard}>
              {/* Row: "I am the …" text + optional edit icon */}
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <Text
                  id={this.state.roleEditTargetId}
                  variant='large'
                  styles={{ root: { fontWeight: 600, color: '#0078d4' } }}
                >
                  {userRoleContext.displayText}
                </Text>

                {/* Edit icon — only shown when the web-part allows user editing */}
                {this.props.allowUserRoleEdit && (
                  <IconButton
                    id={`${this.state.roleEditTargetId}-btn`}
                    iconProps={{ iconName: 'Edit' }}
                    title='Edit my active roles'
                    ariaLabel='Edit active roles'
                    styles={{
                      root: {
                        width: 28,
                        height: 28,
                        borderRadius: 4,
                        color: '#0078d4'
                      },
                      rootHovered: { background: '#deecf9' }
                    }}
                    onClick={() => {
                      // Open callout; pre-populate draft with currently active roles
                      const current =
                        this.state.activeUserRoles !== undefined
                          ? this.state.activeUserRoles
                          : this.props.userRoles || [];
                      this.setState({
                        isRoleEditOpen: true,
                        roleEditDraft: [...current]
                      });
                    }}
                  />
                )}
              </div>

              {/* Role edit Callout */}
              {this.state.isRoleEditOpen && (
                <Callout
                  target={`#${this.state.roleEditTargetId}-btn`}
                  directionalHint={DirectionalHint.bottomLeftEdge}
                  onDismiss={() => this.setState({ isRoleEditOpen: false })}
                  styles={{
                    root: { padding: 16, minWidth: 240, maxWidth: 320 }
                  }}
                  setInitialFocus
                >
                  <Text
                    variant='mediumPlus'
                    styles={{ root: { fontWeight: 600, marginBottom: 12, display: 'block' } }}
                  >
                    My Active Roles
                  </Text>

                  <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 16 } }}>
                    {getAvailableRoles(this.props.roleDefinitions).map(role => (
                      <Checkbox
                        key={role.key}
                        label={role.label}
                        checked={this.state.roleEditDraft.indexOf(role.key) !== -1}
                        onChange={(_ev, checked) => {
                          const draft = [...this.state.roleEditDraft];
                          const idx = draft.indexOf(role.key);
                          if (checked && idx === -1) draft.push(role.key);
                          else if (!checked && idx !== -1) draft.splice(idx, 1);
                          this.setState({ roleEditDraft: draft });
                        }}
                      />
                    ))}
                  </Stack>

                  <Stack horizontal tokens={{ childrenGap: 8 }}>
                    <PrimaryButton
                      text='Apply'
                      onClick={() => {
                        this.setState(
                          { activeUserRoles: [...this.state.roleEditDraft], isRoleEditOpen: false },
                          () => this._loadData()
                        );
                      }}
                    />
                    <DefaultButton
                      text='Cancel'
                      onClick={() => this.setState({ isRoleEditOpen: false })}
                    />
                    <DefaultButton
                      text='Reset'
                      title='Reset to property pane defaults'
                      onClick={() => {
                        this.setState(
                          { activeUserRoles: undefined, isRoleEditOpen: false },
                          () => this._loadData()
                        );
                      }}
                    />
                  </Stack>
                </Callout>
              )}
            </div>
          </div>
        )}

        {/* Unified Control Bar (Clear All | Search | Refresh) */}
        {(() => {
          const activeFilterCount = Object.keys(this.state.filters).filter(
            (columnKey: string) =>
              FilterUtils.isFilterActive(this.state.filters[columnKey])
          ).length;
          const hasSearch = !!this.state.searchText.trim();
          const clearDisabled = activeFilterCount === 0 && !hasSearch;
          return (
            <div className={styles.controlBar}>
              <div className={styles.controlBarCenter}>
                <SearchBox
                  placeholder='Search items...'
                  value={this.state.searchText}
                  onChange={(ev, val) => {
                    const searchText = val || '';
                    let filtered = FilterUtils.applyFilters(
                      this.state.workflowItems,
                      this.state.filters
                    );
                    filtered = this._applySearch(filtered, searchText);
                    filtered = this._applyCurrentSorting(filtered);
                    const groupedView = this._createGroupedView(filtered);
                    this.setState({
                      searchText,
                      filteredItems: groupedView.items,
                      groupedItems: groupedView.groups
                    });
                  }}
                />
              </div>
              <div className={styles.controlBarRight}>
                <IconButton
                  iconProps={{ iconName: 'ClearFilter' }}
                  text={
                    activeFilterCount > 0 || hasSearch
                      ? `Clear All${
                          activeFilterCount > 0 ? ` (${activeFilterCount})` : ''
                        }`
                      : 'Clear All'
                  }
                  aria-label='Clear all filters and search'
                  disabled={clearDisabled}
                  onClick={this._onClearAll}
                  styles={{ rootDisabled: { opacity: 0.5 } }}
                />
                <IconButton
                  iconProps={{ iconName: 'Refresh' }}
                  text='Refresh'
                  aria-label='Refresh data'
                  onClick={this._onRefresh}
                  styles={{ root: { marginLeft: 4 } }}
                />
              </div>
            </div>
          );
        })()}

        {/* Loading Spinner */}
        {isLoading && (
          <Stack
            horizontalAlign='center'
            tokens={{ childrenGap: 16, padding: 32 }}
          >
            <Spinner
              size={SpinnerSize.large}
              label='Loading workflow items...'
            />
          </Stack>
        )}

        {/* Items List */}
        {!isLoading && (
          <>
            {filteredItems.length === 0 ? (
              <Stack horizontalAlign='center' tokens={{ padding: 32 }}>
                <Text variant='medium' styles={{ root: { color: '#605e5c' } }}>
                  No workflow items found.{' '}
                  {this.props.workflowLists.length === 0
                    ? 'Please configure workflow lists in the web part properties.'
                    : 'Try adjusting your filters or check that the configured lists contain data.'}
                </Text>
              </Stack>
            ) : (
              <DetailsList
                items={filteredItems}
                columns={columns}
                groups={this.props.enableGrouping ? groupedItems : undefined}
                setKey='set'
                layoutMode={DetailsListLayoutMode.justified}
                selection={this._selection}
                selectionMode={SelectionMode.none}
                selectionPreservedOnEmptyClick={true}
                ariaLabelForSelectionColumn='Toggle selection'
                ariaLabelForSelectAllCheckbox='Toggle selection for all items'
                checkButtonAriaLabel='select row'
                styles={{
                  root: {
                    '& .ms-DetailsHeader': {
                      paddingTop: 16
                    }
                  }
                }}
              />
            )}
          </>
        )}

        {/* Filter Overlay */}
        {activeFilterColumn && filterTargetRef && (
          <FilterOverlay
            target={filterTargetRef}
            isVisible={true}
            onDismiss={this._onFilterDismiss}
            onApply={(cfg: IFilterConfig) =>
              this._onFilterApply(activeFilterColumn, cfg)
            }
            onClear={() => this._onFilterClear(activeFilterColumn)}
            columnKey={activeFilterColumn}
            columnName={this._getColumnName(columns, activeFilterColumn)}
            filterConfig={
              filters[activeFilterColumn] ||
              FilterUtils.getColumnFilterConfig(
                activeFilterColumn,
                this.state.workflowItems
              )
            }
          />
        )}
      </div>
    );
  }
}

// Wrapper component that uses hooks to get services
const WorkflowItemsRequested: React.FC<
  IWorkflowItemsRequestedProps
> = props => {
  const { sp, graph } = useServices();

  // Create service instance
  const workflowService = React.useMemo(() => {
    if (!sp || !graph) return undefined;
    return new WorkflowItemsRequestedService(sp, graph);
  }, [sp, graph]);

  return (
    <WorkflowItemsRequestedClass {...props} workflowService={workflowService} />
  );
};

export default WorkflowItemsRequested;
