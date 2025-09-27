import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { PropertyFieldListPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import { graphfi, GraphFI, SPFx as GraphSPFx } from '@pnp/graph';

import WorkflowItemsRequested from './components/WorkflowItemsRequested';
import { IWorkflowItemsRequestedProps } from './components/IWorkflowItemsRequestedProps';
import { AppConfigProvider } from '../../contexts/AppConfigContext';
import { ServiceProvider } from '../../contexts/ServiceContext';

export interface IWorkflowItemsRequestedWebPartProps {
  // Data Source Settings
  workflowLists: string[]; // Multiple SharePoint lists

  // User Role Configuration
  userRoles: string; // Comma-separated user roles: 'CM,BUYER,DORDSM,EXECUTIVE,GM,PRESIDENT'

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

export default class WorkflowItemsRequestedWebPart extends BaseClientSideWebPart<IWorkflowItemsRequestedWebPartProps> {
  private _sp?: SPFI;
  private _graph?: GraphFI;

  public render(): void {
    const workflowListsKey =
      (this.properties.workflowLists || []).sort().join('|') || 'empty';
    console.log('🔑 Generated workflowLists key:', workflowListsKey);

    // Convert comma-separated userRoles string to array
    const userRolesArray = this.properties.userRoles
      ? this.properties.userRoles
          .split(',')
          .map(role => role.trim())
          .filter(role => role.length > 0)
      : [];

    const element: React.ReactElement<IWorkflowItemsRequestedProps> =
      React.createElement(WorkflowItemsRequested, {
        key: `workflow-items-requested-${workflowListsKey}`, // Force re-mount when lists change

        spContext: this.context, // Pass the web part context

        // Data Source Settings
        workflowLists: (this.properties.workflowLists || []).slice(),

        // User Role Configuration
        userRoles: userRolesArray,

        // Column Display Options
        showWorkflowStatus: this.properties.showWorkflowStatus,
        showCreditManager: this.properties.showCreditManager,
        showDSR: this.properties.showDSR,
        showCustomerService: this.properties.showCustomerService,
        showCurrentAssignedRole: this.properties.showCurrentAssignedRole,
        showCreatedDate: this.properties.showCreatedDate,
        showSourceList: this.properties.showSourceList,

        // Grouping and Sorting
        enableGrouping: this.properties.enableGrouping,
        defaultSortColumn: this.properties.defaultSortColumn,
        defaultSortDirection: this.properties.defaultSortDirection,

        // Filtering
        enableRoleBasedFiltering: this.properties.enableRoleBasedFiltering,
        statusFilter: this.properties.statusFilter || [],

        // Advanced Settings
        refreshInterval: this.properties.refreshInterval,
        showUserContext: this.properties.showUserContext,

        webpartTitle: this.properties.webpartTitle
      });

    // Wrap with AppConfigProvider for multi-list context
    const appConfigElement: React.ReactElement = React.createElement(
      AppConfigProvider,
      {
        key: `app-config-${workflowListsKey}`, // Force re-mount when lists change
        selectedListTitles: (this.properties.workflowLists || []).slice(),
        currentSiteUrl: this.context.pageContext.web.absoluteUrl,
        onRefreshData: async () => {
          console.log('Refresh data requested from app config');
        }
      },
      element
    );

    const providerElement: React.ReactElement = React.createElement(
      ServiceProvider,
      {
        sp: this._sp ?? spfi().using(SPFx(this.context)),
        graph: this._graph ?? graphfi().using(GraphSPFx(this.context))
      },
      appConfigElement
    );

    ReactDom.render(providerElement, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Initialize PnP JS instances
    this._sp = spfi().using(SPFx(this.context));
    this._graph = graphfi().using(GraphSPFx(this.context));
    return Promise.resolve();
  }

  protected onPropertyPaneConfigurationStart(): void {
    // Set default values for properties if they haven't been set
    if (this.properties.workflowLists === undefined) {
      this.properties.workflowLists = ['Documents', 'Site Pages']; // Default lists for testing
    }
    if (this.properties.userRoles === undefined) {
      this.properties.userRoles = '';
    }
    if (this.properties.showWorkflowStatus === undefined) {
      this.properties.showWorkflowStatus = true;
    }
    if (this.properties.showCreditManager === undefined) {
      this.properties.showCreditManager = true;
    }
    if (this.properties.showDSR === undefined) {
      this.properties.showDSR = true;
    }
    if (this.properties.showCustomerService === undefined) {
      this.properties.showCustomerService = true;
    }
    if (this.properties.showCurrentAssignedRole === undefined) {
      this.properties.showCurrentAssignedRole = true;
    }
    if (this.properties.showCreatedDate === undefined) {
      this.properties.showCreatedDate = false;
    }
    if (this.properties.enableGrouping === undefined) {
      this.properties.enableGrouping = true;
    }
    if (this.properties.showSourceList === undefined) {
      this.properties.showSourceList = false;
    }
    if (this.properties.defaultSortColumn === undefined) {
      this.properties.defaultSortColumn = 'title';
    }
    if (this.properties.defaultSortDirection === undefined) {
      this.properties.defaultSortDirection = 'asc';
    }
    if (this.properties.enableRoleBasedFiltering === undefined) {
      this.properties.enableRoleBasedFiltering = true;
    }
    if (this.properties.statusFilter === undefined) {
      this.properties.statusFilter = ['In Progress', 'Pending', 'Completed'];
    }
    if (this.properties.refreshInterval === undefined) {
      this.properties.refreshInterval = 0;
    }
    if (this.properties.showUserContext === undefined) {
      this.properties.showUserContext = true;
    }
    if (this.properties.webpartTitle === undefined) {
      this.properties.webpartTitle = 'Workflow Items Requested';
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure your Workflow Items Requested web part'
          },
          groups: [
            {
              groupName: 'Data Source Settings',
              groupFields: [
                PropertyFieldListPicker('workflowLists', {
                  label: 'Select workflow lists',
                  selectedList: this.properties.workflowLists || [],
                  includeHidden: false,
                  orderBy: 2, // Title
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  key: 'listPickerFieldId',
                  multiSelect: true
                }),
                PropertyPaneTextField('description', {
                  label: 'Description'
                })
              ]
            }
          ]
        },
        {
          header: {
            description: 'Column Display Options'
          },
          groups: [
            {
              groupName: 'Visible Columns',
              groupFields: [
                PropertyPaneToggle('showWorkflowStatus', {
                  label: 'Show Workflow Status column',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('showCreditManager', {
                  label: 'Show Credit Manager column',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('showDSR', {
                  label: 'Show DSR column',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('showCustomerService', {
                  label: 'Show Customer Service column',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('showCurrentAssignedRole', {
                  label: 'Show Current Assigned Role column',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('showCreatedDate', {
                  label: 'Show Created Date column',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('showSourceList', {
                  label: 'Show Source List column',
                  onText: 'Show',
                  offText: 'Hide'
                })
              ]
            }
          ]
        },
        {
          header: {
            description: 'Advanced Settings'
          },
          groups: [
            {
              groupName: 'Grouping and Sorting',
              groupFields: [
                PropertyPaneToggle('enableGrouping', {
                  label: 'Enable grouping by list',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneDropdown('defaultSortColumn', {
                  label: 'Default sort column',
                  options: [
                    { key: 'title', text: 'Title' },
                    { key: 'workflowStatus', text: 'Workflow Status' },
                    { key: 'creditManager', text: 'Credit Manager' },
                    { key: 'dsr', text: 'DSR' },
                    { key: 'customerService', text: 'Customer Service' },
                    {
                      key: 'currentAssignedRole',
                      text: 'Current Assigned Role'
                    },
                    { key: 'createdDate', text: 'Created Date' }
                  ]
                }),
                PropertyPaneDropdown('defaultSortDirection', {
                  label: 'Default sort direction',
                  options: [
                    { key: 'asc', text: 'Ascending' },
                    { key: 'desc', text: 'Descending' }
                  ]
                })
              ]
            },
            {
              groupName: 'Filtering and User Context',
              groupFields: [
                PropertyPaneToggle('enableRoleBasedFiltering', {
                  label: 'Enable role-based filtering',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneTextField('userRoles', {
                  label: 'User roles (comma-separated)',
                  description:
                    'Enter user roles like: CM,BUYER,DORDSM,EXECUTIVE,GM,PRESIDENT',
                  multiline: true
                }),
                PropertyPaneToggle('showUserContext', {
                  label: 'Show user context header',
                  onText: 'Show',
                  offText: 'Hide'
                })
              ]
            }
          ]
        },
        {
          header: {
            description: 'Header and Display Customization'
          },
          groups: [
            {
              groupName: 'Webpart Header',
              groupFields: [
                PropertyPaneToggle('showWebpartHeader', {
                  label: 'Show webpart header',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneTextField('webpartTitle', {
                  label: 'Webpart title',
                  description: 'Main title displayed at the top'
                }),
                PropertyPaneTextField('webpartSubtitle', {
                  label: 'Webpart subtitle',
                  description: 'Subtitle or description text',
                  multiline: true
                })
              ]
            },
            {
              groupName: 'User Context Display',
              groupFields: [
                PropertyPaneTextField('userContextPrefix', {
                  label: 'User context prefix',
                  description:
                    'Text before user role (e.g., "I am the", "Acting as", "Current role:")'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
