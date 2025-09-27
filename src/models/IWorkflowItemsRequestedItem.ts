export interface IWorkflowItemsRequestedItem {
  id: string;
  title: string;
  workflowStatus:
    | 'In Progress'
    | 'Completed'
    | 'Pending'
    | 'Rejected'
    | 'Cancelled';
  creditManager?: string;
  dsr?: string;
  customerService?: string;
  currentAssignedRole?: string;
  createdDate: Date;
  modifiedDate?: Date;
  // Grouping context
  sourceListTitle?: string;
  sourceListId?: string;
  sourceItemId?: number;

  // Optional extras for future parity with other webparts
  requestor?: string;
  priority?: 'High' | 'Medium' | 'Low';
  dueDate?: Date;
}

export interface IRoleFieldMappings {
  title?: string;
  workflowStatus?: string;
  creditManager?: string;
  dsr?: string;
  customerService?: string;
  currentAssignedRole?: string;
}

export interface IUserRoleContext {
  currentUser: string;
  userRoles: string[];
  displayText?: string;
}
