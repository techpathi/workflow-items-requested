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
  SearchBox
} from '@fluentui/react';
import type { IDetailsColumnProps } from '@fluentui/react';
import { FilterOverlay } from '../../../components';
import { useServices } from '../../../contexts/ServiceContext';
import { WorkflowItemsRequestedService } from '../../../services';
import { FilterUtils, IColumnFilter } from '../../../utils';
import { IWorkflowItemsRequestedItem, IUserRoleContext } from '../../../models';

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
  // no-op placeholder for legacy targeting; using event.currentTarget instead

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
      JSON.stringify(this.props.workflowLists)
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
      (prevProps as any).showSourceList !==
        (this.props as any).showSourceList ||
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

      // Load workflow items
      let items = await this._workflowService.getWorkflowItems(
        this.props.workflowLists
      );

      console.log(`Loaded ${items.length} workflow items`);

      // Exclude items with no role (require currentAssignedRole non-empty)
      items = items.filter(
        it => it.currentAssignedRole && it.currentAssignedRole.trim().length > 0
      );
      console.log(
        `After excluding items without roles: ${items.length} items remain`
      );

      // Get user role context
      const userRoleContext = await this._workflowService.getUserRoleContext(
        this.props.userRoles || []
      );

      // Apply role-based filtering if enabled
      let filteredItems = items;
      if (
        this.props.enableRoleBasedFiltering &&
        this.props.userRoles &&
        this.props.userRoles.length > 0
      ) {
        filteredItems = await this._workflowService.filterItemsByUserRoles(
          items,
          this.props.userRoles,
          userRoleContext.currentUser
        );
        console.log(
          `Applied role-based filtering: ${filteredItems.length} items remain`
        );
      }

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

      // Create groups if grouping is enabled
      const groups = this.props.enableGrouping
        ? this._createGroups(columnFilteredItems)
        : [];

      this.setState({
        workflowItems: items,
        filteredItems: columnFilteredItems,
        groupedItems: groups,
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
    const columns: IColumn[] = [];

    // Title column (always visible)
    columns.push({
      key: 'title',
      name: 'Title',
      fieldName: 'title',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      isSorted: this.props.defaultSortColumn === 'title',
      isSortedDescending: this.props.defaultSortDirection === 'desc',
      onRender: (item: IWorkflowItemsRequestedItem) => (
        <Text variant='medium'>{item.title}</Text>
      ),
      onColumnClick: this._onColumnClick
    });

    // Workflow Status column
    if (this.props.showWorkflowStatus) {
      columns.push({
        key: 'workflowStatus',
        name: 'Workflow Status',
        fieldName: 'workflowStatus',
        minWidth: 120,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <Text
            variant='medium'
            styles={{
              root: {
                color: this._getStatusColor(item.workflowStatus),
                fontWeight: 600
              }
            }}
          >
            {item.workflowStatus}
          </Text>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Credit Manager column
    if (this.props.showCreditManager) {
      columns.push({
        key: 'creditManager',
        name: 'Credit Manager',
        fieldName: 'creditManager',
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <Text variant='medium'>{item.creditManager || '-'}</Text>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // DSR column
    if (this.props.showDSR) {
      columns.push({
        key: 'dsr',
        name: 'DSR',
        fieldName: 'dsr',
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <Text variant='medium'>{item.dsr || '-'}</Text>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Customer Service column
    if (this.props.showCustomerService) {
      columns.push({
        key: 'customerService',
        name: 'Customer Service',
        fieldName: 'customerService',
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <Text variant='medium'>{item.customerService || '-'}</Text>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Current Assigned Role column
    if (this.props.showCurrentAssignedRole) {
      columns.push({
        key: 'currentAssignedRole',
        name: 'Current Assigned Role',
        fieldName: 'currentAssignedRole',
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <Text variant='medium'>{item.currentAssignedRole || '-'}</Text>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Created Date column
    if (this.props.showCreatedDate) {
      columns.push({
        key: 'createdDate',
        name: 'Created Date',
        fieldName: 'createdDate',
        minWidth: 120,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <Text variant='medium'>{item.createdDate.toLocaleDateString()}</Text>
        ),
        onColumnClick: this._onColumnClick
      });
    }

    // Source List column
    if (this.props.showSourceList) {
      columns.push({
        key: 'sourceListTitle',
        name: 'Source List',
        fieldName: 'sourceListTitle',
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IWorkflowItemsRequestedItem) => (
          <Text variant='medium'>{item.sourceListTitle || 'Unknown'}</Text>
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

        return (
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              width: '100%',
              height: '100%',
              padding: '0 4px'
            }}
            onClick={(ev: React.MouseEvent<HTMLElement>) => {
              if (props && props.onColumnClick) {
                props.onColumnClick(ev, props.column);
              }
            }}
          >
            <span style={{ flex: 1, cursor: 'pointer' }}>
              {defaultRender ? defaultRender(props) : column.name}
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
                  padding: 0
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

  private _createGroups(items: IWorkflowItemsRequestedItem[]): IGroup[] {
    // Group by source list title
    const groupMap = new Map<string, IWorkflowItemsRequestedItem[]>();

    items.forEach(item => {
      const groupKey = item.sourceListTitle || 'Unknown List';
      if (!groupMap.has(groupKey)) {
        groupMap.set(groupKey, []);
      }
      groupMap.get(groupKey)!.push(item);
    });

    const groups: IGroup[] = [];
    let startIndex = 0;

    groupMap.forEach((groupItems, groupKey) => {
      groups.push({
        key: groupKey,
        name: `${groupKey} (${groupItems.length})`,
        startIndex,
        count: groupItems.length,
        level: 0,
        isCollapsed: false
      });
      startIndex += groupItems.length;
    });

    return groups;
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
    const groups = this.props.enableGrouping
      ? this._createGroups(sortedItems)
      : [];

    this.setState({
      columns: newColumns,
      filteredItems: sortedItems,
      groupedItems: groups
    });
  };

  private _sortItems(
    items: IWorkflowItemsRequestedItem[],
    columnKey: string,
    isSortedDescending: boolean
  ): IWorkflowItemsRequestedItem[] {
    return items.slice().sort((a, b) => {
      let aValue: any;
      let bValue: any;

      switch (columnKey) {
        case 'title':
          aValue = a.title;
          bValue = b.title;
          break;
        case 'workflowStatus':
          aValue = a.workflowStatus;
          bValue = b.workflowStatus;
          break;
        case 'creditManager':
          aValue = a.creditManager;
          bValue = b.creditManager;
          break;
        case 'dsr':
          aValue = a.dsr;
          bValue = b.dsr;
          break;
        case 'customerService':
          aValue = a.customerService;
          bValue = b.customerService;
          break;
        case 'currentAssignedRole':
          aValue = a.currentAssignedRole;
          bValue = b.currentAssignedRole;
          break;
        case 'createdDate':
          aValue = a.createdDate;
          bValue = b.createdDate;
          break;
        default:
          return 0;
      }

      if (aValue === null || aValue === undefined) aValue = '';
      if (bValue === null || bValue === undefined) bValue = '';

      if (aValue < bValue) {
        return isSortedDescending ? 1 : -1;
      }
      if (aValue > bValue) {
        return isSortedDescending ? -1 : 1;
      }
      return 0;
    });
  }

  private _getColumnName(columns: IColumn[], columnKey: string): string {
    for (let i = 0; i < columns.length; i++) {
      if (columns[i].key === columnKey) {
        return columns[i].name;
      }
    }
    return columnKey;
  }

  private _onFilterColumnClick = (columnKey: string, event: any): void => {
    event.stopPropagation();
    const targetRef: React.RefObject<HTMLElement> = {
      current: event.currentTarget as HTMLElement
    } as any;

    // Migration: if title filter previously stored as text, upgrade to checkbox config
    if (columnKey === 'title') {
      const existing = this.state.filters['title'];
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
    this.setState({
      activeFilterColumn: columnKey,
      filterTargetRef: targetRef
    });
  };

  private _onFilterDismiss = (): void => {
    this.setState({
      activeFilterColumn: undefined,
      filterTargetRef: undefined
    });
  };

  private _onFilterApply = (columnKey: string, config: any): void => {
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
    const groups = this.props.enableGrouping
      ? this._createGroups(filteredItems)
      : [];

    this.setState({
      filters: newFilters,
      filteredItems,
      groupedItems: groups
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
    void this._loadData();
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
    const groups = this.props.enableGrouping ? this._createGroups(sorted) : [];
    this.setState({
      filters: clearedFilters,
      filteredItems: sorted,
      groupedItems: groups,
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
        {this.props.showUserContext &&
          userRoleContext &&
          userRoleContext.displayText && (
            <div className={styles.userContext}>
              <div className={styles.userContextCard}>
                <Text
                  variant='large'
                  styles={{
                    root: {
                      fontWeight: 600,
                      marginBottom: 4,
                      color: '#0078d4'
                    }
                  }}
                >
                  {userRoleContext.displayText}
                </Text>
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
                    const groups = this.props.enableGrouping
                      ? this._createGroups(filtered)
                      : [];
                    this.setState({
                      searchText,
                      filteredItems: filtered,
                      groupedItems: groups
                    });
                  }}
                  styles={{ root: { width: '100%' } }}
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
            onApply={(cfg: any) => this._onFilterApply(activeFilterColumn, cfg)}
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
