# Workflow Items Requested WebPart Implementation Plan

## Overview

This document outlines the comprehensive plan to implement a SharePoint Framework (SPFx) webpart that mimics the "Items Requested" interface shown in the provided screenshot. The implementation leverages existing patterns and architecture from the `workflow-requests-items-created-by-me` webpart in the workspace.

## ✅ IMPLEMENTATION COMPLETED (September 21, 2025)

### Implementation Summary

Successfully updated the Items Requested webpart with grouping by list title and default columns matching the screenshot. The solution reuses shared TypeScript sources (models, services, components) and implements role-based filtering with a modern React functional component pattern.

### Key Changes Made

#### 1. TypeScript Sources Added

- **`src/models/IWorkflowItemsRequestedItem.ts`**: Core data model with grouping context
- **`src/services/WorkflowItemsRequestedService.ts`**: Multi-list data loading with ES5 compatibility
- **`src/components/FilterOverlay.tsx`**: Reusable filter overlay component
- **`src/index.ts`**: Barrel exports for clean imports

#### 2. UI Component Updates

- **Fixed import typo**: MessageBar,KW → MessageBar, MessageBarType
- **Parameter typing**: Added explicit types for filter callbacks
- **Grouping implementation**: Groups items by `sourceListTitle` in DetailsList
- **Default columns**: Title, Workflow Status, Credit Manager, DSR, Customer Service, Current Assigned Role (Created Date hidden by default)
- **Status colors**: Inline styles matching created-by-me colors

#### 3. Property Defaults Verified

- `showWorkflowStatus: true`
- `showCreditManager: true`
- `showDSR: true`
- `showCustomerService: true`
- `showCurrentAssignedRole: true`
- `showCreatedDate: false`
- `enableGrouping: true`
- `defaultSortColumn: 'title'`
- `defaultSortDirection: 'asc'`

#### 4. Build & Quality Assurance

- **Build Status**: ✅ Successful (warnings only, no errors)
- **Dev Server**: ✅ Running at https://localhost:4321
- **Features Verified**:
  - Grouping by source list title
  - Column filtering with FilterOverlay
  - Default column visibility
  - Status color coding
  - Error/warning handling for partial list loads

### Next Steps & Future Enhancements

1. **Live Testing**: Deploy to SharePoint environment and test with real lists
2. **Performance**: Consider virtualization for large datasets
3. **Accessibility**: Add ARIA labels for screen readers
4. **i18n**: Implement localization for international users
5. **Advanced Filters**: Date range, multi-select status filters
6. **Export**: Add Excel/CSV export functionality

---

## Original Implementation Plan (Reference)

### Phase 1: Architecture Setup

#### 1.1 Update Package Dependencies

```json
{
  "@fluentui/react": "^8.106.4",
  "@pnp/sp": "^4.16.0",
  "@pnp/graph": "^4.16.0",
  "@pnp/spfx-property-controls": "^3.21.0"
}
```

- Copy `contexts/ServiceContext.tsx` for PnP JS integration
- Copy `contexts/AppConfigContext.tsx` for application configuration
- Copy `services/` folder for data access patterns
- Copy `utils/` folder for helper utilities

#### 1.3 Create Data Model

```typescript
export interface IWorkflowItemsRequestedItem {
  id: string;
  title: string; // Simple text field
  workflowStatus:
    | 'In Progress'
    | 'Completed'
    | 'Pending'
    | 'Rejected'
    | 'Cancelled';
  creditManager: string; // User field for role-based filtering
  dsr: string; // User field for role-based filtering
  customerService: string; // User field for role-based filtering
  currentAssignedRole: string; // User field for role-based filtering
  createdDate: Date;
  modifiedDate: Date;
  sourceListTitle: string; // For grouping by list
  sourceItemId: number;
  sourceListId: string;
}

export interface IUserRoleContext {
  currentUser: string;
  userRoles: string[]; // ['CM', 'BUYER', 'DORDSM', 'EXECUTIVE', 'GM', 'PRESIDENT']
  displayText: string; // "I am the CM,BUYER,DORDSM,EXECUTIVE,GM,PRESIDENT"
}
```

### Phase 2: WebPart Properties and Configuration

#### 2.1 Update WebPart Properties Interface

```typescript
export interface IWorkflowItemsRequestedWebPartProps {
  // Data Source Settings
  workflowLists: string[]; // Multiple SharePoint lists

  // User Role Configuration
  userRoles: string[]; // User's current roles ['CM', 'BUYER', etc.]
  roleFieldMappings: {
    creditManager: string; // Field name in lists
    dsr: string; // Field name in lists
    customerService: string; // Field name in lists
    currentAssignedRole: string; // Field name in lists
  };

  // Column Display Options
  showWorkflowStatus: boolean;
  showCreditManager: boolean;
  showDSR: boolean;
  showCustomerService: boolean;
  showCurrentAssignedRole: boolean;
  showCreatedDate: boolean;

  // Grouping and Sorting (Group by List)
  enableGrouping: boolean; // Group by source list
  defaultSortColumn: string;
  defaultSortDirection: 'asc' | 'desc';

  // Filtering
  enableRoleBasedFiltering: boolean; // Show only items where user matches role fields
  statusFilter: string[]; // Which statuses to show

  // Advanced Settings
  refreshInterval: number;
  showUserContext: boolean; // Show "I am the..." header

  // Setup Panel
  showSetupPanel: boolean;
}
```

#### 2.2 Implement Property Pane Configuration

- **Data Source Tab**: List picker for selecting workflow lists
- **Display Options Tab**: Column visibility toggles
- **Advanced Settings Tab**: Grouping, sorting, and refresh settings

### Phase 3: Component Implementation

#### 3.1 Main Component Structure

```typescript
interface IWorkflowItemsRequestedState {
  workflowItems: IWorkflowItemsRequestedItem[];
  filteredItems: IWorkflowItemsRequestedItem[];
  groupedItems: IGroup[];
  columns: IColumn[];
  isLoading: boolean;
  error?: string;
  currentUserRoles: string[];
}
```

#### 3.2 Key Features Implementation

**User Context Display**

- Show current user's roles at the top
- Format: "I am the {ROLE1},{ROLE2},{ROLE3}..."

**Data Grid with Grouping**

- Use Fluent UI DetailsList component with groups
- Group by source SharePoint list title
- Support expand/collapse for each list group
- Show count of items per group

**Role-Based Filtering**

- Filter items where current user appears in any role field (Credit Manager, DSR, Customer Service, Current Assigned Role)
- Support multiple user roles simultaneously
- Display user context header: "I am the {ROLE1},{ROLE2},{ROLE3}..."

**Status Management**

- Support all workflow status values (In Progress, Completed, Pending, Rejected, Cancelled)
- Color coding for different statuses
- Optional status filtering in property pane

**Data Loading**

- Fetch from multiple SharePoint lists
- Error handling and partial load scenarios
- Auto-refresh functionality

### Phase 4: Service Layer Implementation

#### 4.1 Workflow Items Service

```typescript
export class WorkflowItemsRequestedService {
  constructor(private sp: SPFI, private graph: GraphFI) {}

  async getWorkflowItems(
    listTitles: string[],
    userRoles: string[]
  ): Promise<IWorkflowItemsRequestedItem[]>;
  async getCurrentUserRoles(): Promise<string[]>;
  async filterItemsByUserRoles(
    items: IWorkflowItemsRequestedItem[],
    currentUser: string,
    userRoles: string[]
  ): Promise<IWorkflowItemsRequestedItem[]>;
  async getUserRoleContext(userRoles: string[]): Promise<IUserRoleContext>;
  async refreshData(): Promise<void>;
}
```

#### 4.2 Data Mapping and Transformation

- Map SharePoint list fields to component model
- Handle different list schemas
- Provide default values for missing fields

### Phase 5: Styling and User Experience

#### 5.1 SCSS Styling

- Match the visual design from the screenshot
- Responsive design for different screen sizes
- Proper Fluent UI theming integration

#### 5.2 Interactive Features

- Click handlers for workflow items
- Hover effects and tooltips
- Loading states and error messaging

### Phase 6: Advanced Features

#### 6.1 Filtering and Search

- Command bar with filter options
- Quick search functionality
- Column-specific filters

#### 6.2 Actions and Commands

- Refresh button
- Export functionality
- Bulk operations (if needed)

## File Structure

```
src/
├── webparts/
│   └── workflowItemsRequested/
│       ├── WorkflowItemsRequestedWebPart.ts
│       ├── components/
│       │   ├── WorkflowItemsRequested.tsx
│       │   ├── IWorkflowItemsRequestedProps.ts
│       │   └── WorkflowItemsRequested.module.scss
│       └── loc/
├── contexts/
│   ├── ServiceContext.tsx
│   └── AppConfigContext.tsx
├── services/
│   └── WorkflowItemsRequestedService.ts
├── models/
│   └── IWorkflowItemsRequestedItem.ts
├── utils/
│   ├── SetupUtils.ts
│   └── FilterUtils.ts
└── components/
    ├── SetupPanel.tsx
    └── FilterOverlay.tsx
```

## Implementation Steps

### Step 1: Environment Setup

1. Update package.json with required dependencies
2. Copy context providers and service layer
3. Update tsconfig.json if needed

### Step 2: Data Model and Interfaces

1. Create interfaces for workflow items
2. Update webpart properties interface
3. Create state interfaces for components

### Step 3: WebPart Class Updates

1. Update constructor and property initialization
2. Implement property pane configuration
3. Add service provider integration

### Step 4: Component Development

1. Create main component with basic structure
2. Implement data loading and state management
3. Add DetailsList with grouping
4. Implement column management

### Step 5: Service Implementation

1. Create workflow items service
2. Implement data fetching from SharePoint
3. Add error handling and data transformation

### Step 6: Styling and Polish

1. Update SCSS to match design
2. Add loading states and error handling
3. Implement responsive design

### Step 7: Testing and Refinement

1. Test with different list configurations
2. Verify responsive behavior
3. Optimize performance

## Dependencies and Prerequisites

### Required SharePoint Lists

- Lists must contain fields mappable to the component model
- Minimum required fields: Title, Status, assigned users
- Optional fields: Category, dates, priorities

### User Permissions

- Read access to configured workflow lists
- User profile access for role determination

### Browser Support

- Modern browsers supporting SPFx 1.21.1
- Internet Explorer 11 (if required by organization)

## Success Criteria

1. **Visual Accuracy**: Interface matches the provided screenshot
2. **Functionality**: All features from workflow-requests webpart work
3. **Performance**: Fast loading even with large datasets
4. **Configurability**: Easy setup through property pane
5. **Error Handling**: Graceful handling of missing lists or permissions
6. **Responsive Design**: Works on desktop, tablet, and mobile

## Risk Mitigation

### Data Access Issues

- Implement robust error handling
- Provide clear error messages
- Support partial data loading

### Performance Concerns

- Implement pagination for large datasets
- Use React.memo for optimization
- Implement data caching where appropriate

### Configuration Complexity

- Provide sensible defaults
- Include setup guidance
- Implement validation for required settings

## Timeline Estimate

- **Phase 1-2**: 2-3 days (Architecture and Properties)
- **Phase 3-4**: 3-4 days (Component and Service Implementation)
- **Phase 5-6**: 2-3 days (Styling and Advanced Features)
- **Testing and Polish**: 1-2 days

**Total Estimated Time**: 8-12 days

## Conclusion

This implementation plan provides a comprehensive roadmap for creating a feature-rich workflow items webpart that leverages existing patterns and provides the functionality shown in the target screenshot. By following the existing architecture patterns from the workflow-requests webpart, we ensure consistency, maintainability, and rapid development.
