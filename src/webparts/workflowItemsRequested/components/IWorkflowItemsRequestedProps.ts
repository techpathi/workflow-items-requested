import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IWorkflowItemsRequestedProps {
  // Web Part Context
  spContext: WebPartContext;

  // Data Source Settings
  workflowLists: string[]; // Multiple SharePoint lists

  // User Role Configuration
  userRoles: string[]; // User's current roles ['CM', 'BUYER', etc.]

  // Column Display Options
  showWorkflowStatus: boolean;
  showCreditManager: boolean;
  showDSR: boolean;
  showCustomerService: boolean;
  showCurrentAssignedRole: boolean;
  showCreatedDate: boolean;
  showSourceList: boolean; // Show source list column

  // Grouping and Sorting (Group by List)
  enableGrouping: boolean; // Group by source list
  defaultSortColumn: string;
  defaultSortDirection: 'asc' | 'desc';

  // Filtering
  enableRoleBasedFiltering: boolean; // Show only items where user matches role fields
  statusFilter: string[]; // Which statuses to show

  // Advanced Settings
  refreshInterval: number;

  // Header Customization
  webpartTitle: string; // Main header title
  showUserContext: boolean; // Show "I am the..." header
}
