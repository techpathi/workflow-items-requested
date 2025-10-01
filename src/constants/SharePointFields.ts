/**
 * SharePoint Internal Field Names for Workflow Items
 * These constants centralize all field references used across the solution
 *
 * Note: SharePoint internal names use x0020 to represent spaces in field names
 * Example: "Simple Title" becomes "Simple_x0020_Title"
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
    /** Simple Title field for workflow items */
    SIMPLE_TITLE: 'Simple_x0020_Title',
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

/**
 * Display names for fields - used for documentation and UI labels
 */
export const FIELD_DISPLAY_NAMES = {
  [SHAREPOINT_FIELDS.SYSTEM.ID]: 'ID',
  [SHAREPOINT_FIELDS.SYSTEM.TITLE]: 'Title',
  [SHAREPOINT_FIELDS.SYSTEM.CREATED]: 'Created',
  [SHAREPOINT_FIELDS.SYSTEM.MODIFIED]: 'Modified',
  [SHAREPOINT_FIELDS.BASIC.SIMPLE_TITLE]: 'Simple Title',
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

/**
 * Standard select fields for basic queries
 */
export const BASIC_SELECT_FIELDS = [
  SHAREPOINT_FIELDS.SYSTEM.ID,
  SHAREPOINT_FIELDS.BASIC.SIMPLE_TITLE,
  SHAREPOINT_FIELDS.SYSTEM.TITLE, // Fallback for Simple Title
  SHAREPOINT_FIELDS.BASIC.WORKFLOW_STATUS,
  SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER,
  SHAREPOINT_FIELDS.USERS.DSR,
  SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE,
  SHAREPOINT_FIELDS.BASIC.CURRENT_ASSIGNED_ROLE,
  SHAREPOINT_FIELDS.SYSTEM.CREATED,
  SHAREPOINT_FIELDS.SYSTEM.MODIFIED
];

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
export const isUserField = (fieldName: string): boolean => {
  return (
    fieldName === SHAREPOINT_FIELDS.USERS.CREDIT_MANAGER ||
    fieldName === SHAREPOINT_FIELDS.USERS.DSR ||
    fieldName === SHAREPOINT_FIELDS.USERS.CUSTOMER_SERVICE
  );
};

/**
 * Helper function to get user field expansion properties for SharePoint queries
 */
export const getUserFieldExpansions = (fieldName: string): string[] => {
  if (!isUserField(fieldName)) return [];

  return [
    `${fieldName}/Id`,
    `${fieldName}/Title`,
    `${fieldName}/EMail`,
    `${fieldName}/Email`
  ];
};
