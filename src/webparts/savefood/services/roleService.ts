import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { UserRole, IAdminPerson } from "../models/IRoleMatrix";

export { IAdminPerson, UserRole };

/**
 * Normalizes email or UPN string for strict, case-insensitive comparison
 */
export function normalizeIdentity(val: string): string {
  return (val || "").trim().toLowerCase();
}

export interface IRoleResolutionResult {
  role: UserRole;
  isAdmin: boolean;
  adminUsers: IAdminPerson[];
  error?: string;
}

/**
 * Resolves current user role from RoleMatrix SharePoint list.
 * Default rule: All users are Employees by default.
 * A user is an Admin ONLY if their normalized Email/UPN matches a person in RoleMatrix.Admin.
 * If query fails, silently default to Employee and log diagnostic error.
 */
export async function resolveUserRole(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  userEmail: string,
  userLoginName?: string
): Promise<IRoleResolutionResult> {
  const normEmail = normalizeIdentity(userEmail);
  const normLogin = normalizeIdentity(userLoginName || "");

  try {
    const listUrl = `${siteUrl}/_api/web/lists/getbytitle('RoleMatrix')/items?$filter=Title eq 'Application Roles'&$select=Id,Title,Admin/Id,Admin/Title,Admin/EMail,Admin/UserName&$expand=Admin`;

    const response: SPHttpClientResponse = await spHttpClient.get(
      listUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
          "odata-version": "",
        },
      }
    );

    if (!response.ok) {
      console.warn(`[RoleService] Failed to fetch RoleMatrix (${response.status}: ${response.statusText}). Defaulting to Employee role.`);
      return {
        role: "Employee",
        isAdmin: false,
        adminUsers: [],
        error: `HTTP ${response.status}`,
      };
    }

    const data = await response.json();
    const items = data?.d?.results || [];

    if (items.length === 0) {
      console.warn("[RoleService] No 'Application Roles' item found in RoleMatrix. Defaulting to Employee role.");
      return {
        role: "Employee",
        isAdmin: false,
        adminUsers: [],
      };
    }

    const adminResults = items[0]?.Admin?.results || [];
    const adminUsers: IAdminPerson[] = adminResults.map((p: any) => ({
      Id: p.Id,
      Title: p.Title,
      EMail: p.EMail,
      UserName: p.UserName,
    }));

    // Check normalized comparison against both EMail and UserName/UPN
    const isMatch = adminUsers.some((admin) => {
      const adminEmail = normalizeIdentity(admin.EMail || "");
      const adminUPN = normalizeIdentity(admin.UserName || "");

      return (
        (normEmail && (adminEmail === normEmail || adminUPN === normEmail)) ||
        (normLogin && (adminEmail === normLogin || adminUPN === normLogin))
      );
    });

    return {
      role: isMatch ? "Admin" : "Employee",
      isAdmin: isMatch,
      adminUsers,
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[RoleService] Error resolving user role:", message);
    return {
      role: "Employee",
      isAdmin: false,
      adminUsers: [],
      error: message,
    };
  }
}
