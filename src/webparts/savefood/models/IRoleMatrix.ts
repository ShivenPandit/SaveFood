/**
 * Person field representation in RoleMatrix
 */
export interface IAdminPerson {
  Id: number;
  Title: string;
  EMail?: string;
  UserName?: string;
}

/**
 * Model representing the RoleMatrix SharePoint list item
 */
export interface IRoleMatrixItem {
  Id: number;
  Title: string; // Fixed value: "Application Roles"
  Admin: IAdminPerson[];
}

export type UserRole = "Admin" | "Employee";
