/**
 * SharePoint Internal Field Names for Workflow Items
 * These constants centralize all field references used across the solution
 */

/**
 * Primary SharePoint field constants organized by category
 */
export const SHAREPOINT_FIELDS = {
  // System Fields - Standard SharePoint fields
  SYSTEM: {
    /** Standard SharePoint ID field */
    ID: 'Id',
    /** Standard SharePoint Title field */
    TITLE: 'Title',
    /** Standard SharePoint Created timestamp */
    CREATED: 'Created',
    /** Standard SharePoint Modified timestamp */
    MODIFIED: 'Modified'
  },

  // Basic Workflow Fields - Core workflow-related fields
  BASIC: {
    /** Workflow Status choice field */
    WORKFLOW_STATUS: 'Workflow_x0020_Status',
    /** Current Assigned Role text field */
    CURRENT_ASSIGNED_ROLE: 'Current_x0020_Assigned_x0020_Rol'
  },

  // User/Person Fields - All person or group fields
  USERS: {
    /** Credit Manager person field */
    CREDIT_MANAGER: 'Credit_x0020_Manager',
    /** District Sales Representative person field */
    DSR: 'District_x0020_Sales_x0020_Repre',
    /** Customer Service person field */
    CUSTOMER_SERVICE: 'Customer_x0020_Service'
  }
} as const;

export interface ISharePointFieldConfig {
  id: string;
  title: string;
  created: string;
  modified: string;
  workflowStatus: string;
  currentAssignedRole: string;
  creditManager: string;
  dsr: string;
  customerService: string;
}

export const DEFAULT_SHAREPOINT_FIELD_CONFIG: ISharePointFieldConfig = {
  id: SHAREPOINT_FIELDS.SYSTEM.ID,
  title: SHAREPOINT_FIELDS.SYSTEM.TITLE,
  created: SHAREPOINT_FIELDS.SYSTEM.CREATED,
  modified: SHAREPOINT_FIELDS.SYSTEM.MODIFIED,
  workflowStatus: SHAREPOINT_FIELDS.BASIC.WORKFLOW_STATUS,
  currentAssignedRole: SHAREPOINT_FIELDS.BASIC.CURRENT_ASSIGNED_ROLE,
  creditManager: SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER,
  dsr: SHAREPOINT_FIELDS.USERS.DSR,
  customerService: SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE
};

const getConfiguredFieldName = (
  configuredValue: string | undefined,
  defaultValue: string
): string => {
  const trimmedValue = (configuredValue || '').trim();
  return trimmedValue || defaultValue;
};

export const getSharePointFieldConfig = (
  configuredFields?: Partial<ISharePointFieldConfig>
): ISharePointFieldConfig => {
  return {
    id: getConfiguredFieldName(
      configuredFields && configuredFields.id,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.id
    ),
    title: getConfiguredFieldName(
      configuredFields && configuredFields.title,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.title
    ),
    created: getConfiguredFieldName(
      configuredFields && configuredFields.created,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.created
    ),
    modified: getConfiguredFieldName(
      configuredFields && configuredFields.modified,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.modified
    ),
    workflowStatus: getConfiguredFieldName(
      configuredFields && configuredFields.workflowStatus,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.workflowStatus
    ),
    currentAssignedRole: getConfiguredFieldName(
      configuredFields && configuredFields.currentAssignedRole,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.currentAssignedRole
    ),
    creditManager: getConfiguredFieldName(
      configuredFields && configuredFields.creditManager,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.creditManager
    ),
    dsr: getConfiguredFieldName(
      configuredFields && configuredFields.dsr,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.dsr
    ),
    customerService: getConfiguredFieldName(
      configuredFields && configuredFields.customerService,
      DEFAULT_SHAREPOINT_FIELD_CONFIG.customerService
    )
  };
};

/**
 * Display names for fields - used for documentation and UI labels
 */
export const FIELD_DISPLAY_NAMES = {
  [SHAREPOINT_FIELDS.SYSTEM.ID]: 'ID',
  [SHAREPOINT_FIELDS.SYSTEM.TITLE]: 'Title',
  [SHAREPOINT_FIELDS.SYSTEM.CREATED]: 'Created',
  [SHAREPOINT_FIELDS.SYSTEM.MODIFIED]: 'Modified',
  [SHAREPOINT_FIELDS.BASIC.WORKFLOW_STATUS]: 'Workflow Status',
  [SHAREPOINT_FIELDS.BASIC.CURRENT_ASSIGNED_ROLE]: 'Current Assigned Role',
  [SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER]: 'Credit Manager',
  [SHAREPOINT_FIELDS.USERS.DSR]: 'District Sales Representative',
  [SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE]: 'Customer Service'
} as const;

/**
 * User field mappings for easier iteration and validation
 */
export const USER_FIELD_MAPPINGS = [
  {
    internal: SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER,
    display: FIELD_DISPLAY_NAMES[SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER]
  },
  {
    internal: SHAREPOINT_FIELDS.USERS.DSR,
    display: FIELD_DISPLAY_NAMES[SHAREPOINT_FIELDS.USERS.DSR]
  },
  {
    internal: SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE,
    display: FIELD_DISPLAY_NAMES[SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE]
  }
] as const;

/**
 * All user fields as an array for easier processing
 */
export const USER_FIELDS = [
  SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER,
  SHAREPOINT_FIELDS.USERS.DSR,
  SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE
];

export const getConfiguredUserFieldMappings = (
  configuredFields?: Partial<ISharePointFieldConfig>
): { internal: string; display: string }[] => {
  const fields = getSharePointFieldConfig(configuredFields);

  return [
    {
      internal: fields.creditManager,
      display: 'Credit Manager'
    },
    {
      internal: fields.dsr,
      display: 'District Sales Representative'
    },
    {
      internal: fields.customerService,
      display: 'Customer Service'
    }
  ];
};

export const getConfiguredUserFields = (
  configuredFields?: Partial<ISharePointFieldConfig>
): string[] => {
  const fields = getSharePointFieldConfig(configuredFields);
  return [fields.creditManager, fields.dsr, fields.customerService];
};

const uniqueFields = (fields: string[]): string[] => {
  return fields.filter(
    (fieldName, index, allFields) =>
      fieldName && allFields.indexOf(fieldName) === index
  );
};

/**
 * Standard select fields for basic queries
 */
export const BASIC_SELECT_FIELDS = [
  SHAREPOINT_FIELDS.SYSTEM.ID,
  SHAREPOINT_FIELDS.SYSTEM.TITLE,
  SHAREPOINT_FIELDS.BASIC.WORKFLOW_STATUS,
  SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER,
  SHAREPOINT_FIELDS.USERS.DSR,
  SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE,
  SHAREPOINT_FIELDS.BASIC.CURRENT_ASSIGNED_ROLE,
  SHAREPOINT_FIELDS.SYSTEM.CREATED,
  SHAREPOINT_FIELDS.SYSTEM.MODIFIED
];

export const getBasicSelectFields = (
  configuredFields?: Partial<ISharePointFieldConfig>
): string[] => {
  const fields = getSharePointFieldConfig(configuredFields);

  return uniqueFields([
    fields.id,
    fields.title,
    fields.workflowStatus,
    fields.creditManager,
    fields.dsr,
    fields.customerService,
    fields.currentAssignedRole,
    fields.created,
    fields.modified
  ]);
};

/**
 * Type definitions for better type safety
 */
export type SharePointFieldKey =
  | keyof typeof SHAREPOINT_FIELDS.SYSTEM
  | keyof typeof SHAREPOINT_FIELDS.BASIC
  | keyof typeof SHAREPOINT_FIELDS.USERS;

export type SharePointFieldValue =
  | (typeof SHAREPOINT_FIELDS.SYSTEM)[keyof typeof SHAREPOINT_FIELDS.SYSTEM]
  | (typeof SHAREPOINT_FIELDS.BASIC)[keyof typeof SHAREPOINT_FIELDS.BASIC]
  | (typeof SHAREPOINT_FIELDS.USERS)[keyof typeof SHAREPOINT_FIELDS.USERS];

/**
 * Helper function to get display name for a field
 */
export const getFieldDisplayName = (internalName: string): string => {
  return (
    FIELD_DISPLAY_NAMES[internalName as keyof typeof FIELD_DISPLAY_NAMES] ||
    internalName
  );
};

/**
 * Helper function to check if a field is a user field
 */
export const isUserField = (
  fieldName: string,
  configuredFields?: Partial<ISharePointFieldConfig>
): boolean => {
  const fields = getSharePointFieldConfig(configuredFields);

  return (
    fieldName === fields.creditManager ||
    fieldName === fields.dsr ||
    fieldName === fields.customerService
  );
};

/**
 * Helper function to get user field expansion properties for SharePoint queries
 */
export const getUserFieldExpansions = (
  fieldName: string,
  configuredFields?: Partial<ISharePointFieldConfig>
): string[] => {
  if (!isUserField(fieldName, configuredFields)) return [];

  return [`${fieldName}/Id`, `${fieldName}/Title`, `${fieldName}/EMail`];
};
