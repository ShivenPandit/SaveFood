/**
 * Vedanta VOGL Save Food Management System
 * Portal Environments & URLs Configuration
 */

export interface IPortalEnvironment {
  key: "uat" | "production";
  name: string;
  url: string;
  description: string;
}

export const PORTAL_ENVIRONMENTS: Record<"UAT" | "PRODUCTION", IPortalEnvironment> = {
  UAT: {
    key: "uat",
    name: "Save Food UAT Portal",
    url: "https://vedanta.sharepoint.com/sites/SaveFoodUAT",
    description: "User Acceptance Testing & Staging Environment",
  },
  PRODUCTION: {
    key: "production",
    name: "Save Food Production Portal",
    url: "https://vedanta.sharepoint.com/sites/SaveFood",
    description: "Live Enterprise Production Environment",
  },
};

/**
 * Resolve current portal environment based on SharePoint web context URL.
 */
export const getPortalEnvironment = (currentWebUrl?: string): IPortalEnvironment => {
  const urlLower = (currentWebUrl || "").toLowerCase();
  if (urlLower.includes("savefooduat")) {
    return PORTAL_ENVIRONMENTS.UAT;
  }
  return PORTAL_ENVIRONMENTS.PRODUCTION;
};

/**
 * Default fallback site URL when context is not yet loaded or in local mock harness.
 */
export const DEFAULT_PORTAL_URL: string = PORTAL_ENVIRONMENTS.UAT.url;

