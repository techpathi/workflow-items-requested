import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { PropertyFieldListPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import { graphfi, GraphFI, SPFx as GraphSPFx } from '@pnp/graph';

import WorkflowItemsRequested from './components/WorkflowItemsRequested';
import { IWorkflowItemsRequestedProps } from './components/IWorkflowItemsRequestedProps';
import { AppConfigProvider } from '../../contexts/AppConfigContext';
import { ServiceProvider } from '../../contexts/ServiceContext';
import {
  DEFAULT_SHAREPOINT_FIELD_CONFIG,
  ISharePointFieldConfig,
  getSharePointFieldConfig
} from '../../constants';

const DEFAULT_CREATED_WITHIN_DAYS = 180;

const USER_ROLE_OPTIONS: {
  key: string;
  text: string;
}[] = [
  { key: 'CM', text: 'CM' },
  { key: 'BUYER', text: 'Buyer' },
  { key: 'DSR', text: 'DSR' },
  { key: 'EXECUTIVE', text: 'Executive' },
  { key: 'GM', text: 'GM' },
  { key: 'PRESIDENT', text: 'President' },
  { key: 'CUSTOMERSERVICE', text: 'Customer Service' }
];

export interface IWorkflowItemsRequestedWebPartProps {
  // Data Source Settings
  workflowLists: string[]; // Multiple SharePoint lists
  description?: string;
  createdWithinDays: number;

  // SharePoint Internal Field Names
  spFieldId: string;
  spFieldTitle: string;
  spFieldCreated: string;
  spFieldModified: string;
  spFieldWorkflowStatus: string;
  spFieldCurrentAssignedRole: string;
  spFieldCreditManager: string;
  spFieldDSR: string;
  spFieldCustomerService: string;

  // User Role Configuration
  userRoles: string; // Comma-separated user roles selected for active filtering
  roleDefinitions: { key: string; text: string }[];

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
  showWebpartHeader?: boolean;
  webpartSubtitle?: string;
  userContextPrefix?: string;
  allowUserRoleEdit: boolean; // Allow end-users to temporarily override their active roles
}

export default class WorkflowItemsRequestedWebPart extends BaseClientSideWebPart<IWorkflowItemsRequestedWebPartProps> {
  private _sp?: SPFI;
  private _graph?: GraphFI;

  public render(): void {
    const workflowListsKey =
      (this.properties.workflowLists || []).slice().sort().join('|') ||
      'empty';
    console.log('🔑 Generated workflowLists key:', workflowListsKey);

    const userRolesArray = this._parseUserRoles(this.properties.userRoles);

    const element: React.ReactElement<IWorkflowItemsRequestedProps> =
      React.createElement(WorkflowItemsRequested, {
        key: `workflow-items-requested-${workflowListsKey}`, // Force re-mount when lists change

        spContext: this.context, // Pass the web part context

        // Data Source Settings
        workflowLists: (this.properties.workflowLists || []).slice(),
        sharePointFields: this._getSharePointFieldConfig(),
        createdWithinDays: this._getCreatedWithinDays(),

        // User Role Configuration
        userRoles: userRolesArray,
        roleDefinitions: this.properties.roleDefinitions || [],

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
        allowUserRoleEdit: !!this.properties.allowUserRoleEdit,

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
      this.properties.workflowLists = ['Documents', 'Site Pages']; // Default lists
    }
    this._setDefaultRoleDefinitions();
    this._setDefaultSharePointFields();
    if (
      this.properties.createdWithinDays === undefined ||
      this.properties.createdWithinDays <= 0
    ) {
      this.properties.createdWithinDays = DEFAULT_CREATED_WITHIN_DAYS;
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

  private _setDefaultSharePointFields(): void {
    if (this.properties.spFieldId === undefined) {
      this.properties.spFieldId = DEFAULT_SHAREPOINT_FIELD_CONFIG.id;
    }
    if (this.properties.spFieldTitle === undefined) {
      this.properties.spFieldTitle = DEFAULT_SHAREPOINT_FIELD_CONFIG.title;
    }
    if (this.properties.spFieldCreated === undefined) {
      this.properties.spFieldCreated = DEFAULT_SHAREPOINT_FIELD_CONFIG.created;
    }
    if (this.properties.spFieldModified === undefined) {
      this.properties.spFieldModified = DEFAULT_SHAREPOINT_FIELD_CONFIG.modified;
    }
    if (this.properties.spFieldWorkflowStatus === undefined) {
      this.properties.spFieldWorkflowStatus =
        DEFAULT_SHAREPOINT_FIELD_CONFIG.workflowStatus;
    }
    if (this.properties.spFieldCurrentAssignedRole === undefined) {
      this.properties.spFieldCurrentAssignedRole =
        DEFAULT_SHAREPOINT_FIELD_CONFIG.currentAssignedRole;
    }
    if (this.properties.spFieldCreditManager === undefined) {
      this.properties.spFieldCreditManager =
        DEFAULT_SHAREPOINT_FIELD_CONFIG.creditManager;
    }
    if (this.properties.spFieldDSR === undefined) {
      this.properties.spFieldDSR = DEFAULT_SHAREPOINT_FIELD_CONFIG.dsr;
    }
    if (this.properties.spFieldCustomerService === undefined) {
      this.properties.spFieldCustomerService =
        DEFAULT_SHAREPOINT_FIELD_CONFIG.customerService;
    }
  }

  private _getSharePointFieldConfig(): ISharePointFieldConfig {
    return getSharePointFieldConfig({
      id: this.properties.spFieldId,
      title: this.properties.spFieldTitle,
      created: this.properties.spFieldCreated,
      modified: this.properties.spFieldModified,
      workflowStatus: this.properties.spFieldWorkflowStatus,
      currentAssignedRole: this.properties.spFieldCurrentAssignedRole,
      creditManager: this.properties.spFieldCreditManager,
      dsr: this.properties.spFieldDSR,
      customerService: this.properties.spFieldCustomerService
    });
  }

  private _setDefaultRoleDefinitions(): void {
    if (!this.properties.roleDefinitions || this.properties.roleDefinitions.length === 0) {
      this.properties.roleDefinitions = USER_ROLE_OPTIONS.map(opt => ({
        key: opt.key,
        text: opt.text
      }));
      this.properties.userRoles = this._getSelectedUserRoles().join(',');
    }
  }

  private _getSelectedUserRoles(): string[] {
    return (this.properties.roleDefinitions || []).map(r => r.key);
  }

  private _parseUserRoles(userRoles: string | undefined): string[] {
    if (!userRoles) return [];

    const roles: string[] = [];
    userRoles
      .split(',')
      .map(role => this._normalizeRoleKey(role.trim()))
      .filter(role => role.length > 0)
      .forEach(role => {
        if (roles.indexOf(role) === -1) {
          roles.push(role);
        }
      });
    return roles;
  }

  private _normalizeRoleKey(role: string): string {
    const normalizedRole = (role || '').toUpperCase();
    return normalizedRole === 'DORDSM' ? 'DSR' : normalizedRole;
  }

  private _getCreatedWithinDays(): number {
    const value = Number(this.properties.createdWithinDays);
    if (!value || value <= 0) return DEFAULT_CREATED_WITHIN_DAYS;
    return Math.floor(value);
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'roleDefinitions') {
      this.properties.userRoles = this._getSelectedUserRoles().join(',');
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
                  context: this.context,
                  key: 'listPickerFieldId',
                  multiSelect: true
                }),
                PropertyPaneTextField('description', {
                  label: 'Description'
                }),
                PropertyPaneSlider('createdWithinDays', {
                  label: 'Created in last days',
                  min: 1,
                  max: 365,
                  step: 1,
                  showValue: true,
                  value: this._getCreatedWithinDays()
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
            description: 'SharePoint Field Mapping'
          },
          groups: [
            {
              groupName: 'Internal Column Names',
              groupFields: [
                PropertyPaneTextField('spFieldId', {
                  label: 'ID field',
                  description: 'Internal name sent to SharePoint for item ID'
                }),
                PropertyPaneTextField('spFieldTitle', {
                  label: 'Title field',
                  description: 'Internal name sent to SharePoint for the title'
                }),
                PropertyPaneTextField('spFieldWorkflowStatus', {
                  label: 'Workflow Status field',
                  description:
                    'Internal name sent to SharePoint for workflow status'
                }),
                PropertyPaneTextField('spFieldCreditManager', {
                  label: 'Credit Manager field',
                  description:
                    'Internal name sent to SharePoint for the Credit Manager person field'
                }),
                PropertyPaneTextField('spFieldDSR', {
                  label: 'DSR field',
                  description:
                    'Internal name sent to SharePoint for the DSR person field'
                }),
                PropertyPaneTextField('spFieldCustomerService', {
                  label: 'Customer Service field',
                  description:
                    'Internal name sent to SharePoint for the Customer Service person field'
                }),
                PropertyPaneTextField('spFieldCurrentAssignedRole', {
                  label: 'Current Assigned Role field',
                  description:
                    'Internal name sent to SharePoint for current assigned role'
                }),
                PropertyPaneTextField('spFieldCreated', {
                  label: 'Created field',
                  description:
                    'Internal name sent to SharePoint for created date'
                }),
                PropertyPaneTextField('spFieldModified', {
                  label: 'Modified field',
                  description:
                    'Internal name sent to SharePoint for modified date'
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
                PropertyFieldCollectionData('roleDefinitions', {
                  key: 'roleDefinitions',
                  label: 'User role definitions',
                  panelHeader: 'Edit Role Definitions',
                  manageBtnLabel: 'Manage Roles',
                  value: this.properties.roleDefinitions,
                  fields: [
                    {
                      id: 'key',
                      title: 'Internal Key (e.g. FINANCE)',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'text',
                      title: 'Display Name (e.g. Finance Team)',
                      type: CustomCollectionFieldType.string,
                      required: true
                    }
                  ],
                  disabled: false
                }),
                PropertyPaneToggle('showUserContext', {
                  label: 'Show user context header',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('allowUserRoleEdit', {
                  label: 'Allow users to edit their active roles',
                  onText: 'Enabled',
                  offText: 'Disabled (property pane only)'
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
