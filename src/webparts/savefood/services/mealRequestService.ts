import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ILocationMasterItem } from "../models/ILocation";
import {
  IMealOptOutRequestItem,
  ICreateMealOptOutPayload,
  MealType,
} from "../models/IMealOptOut";
import { validateMealCutOff, validateRetractionCutOff } from "./cutOffService";

export interface IPagedResult<T> {
  items: T[];
  nextUrl: string | null;
  totalCount?: number;
}

export interface IBatchCreateResult {
  success: boolean;
  successfulMeals: MealType[];
  failedMeals: { mealType: MealType; error: string; isDuplicate?: boolean }[];
  message: string;
}

/**
 * Format ISO date string (e.g. "2026-11-26" or "2026-11-26T00:00:00Z") to "26 Nov 2026"
 */
export const formatIsoToDisplayDate = (isoStr: string): string => {
  if (!isoStr) return "";
  try {
    const raw = isoStr.split("T")[0];
    const parts = raw.split("-");
    if (parts.length === 3) {
      const year = parts[0];
      const monthIdx = parseInt(parts[1], 10) - 1;
      const day = parts[2];
      const monthNames = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
      ];
      if (monthIdx >= 0 && monthIdx < 12) {
        return `${day} ${monthNames[monthIdx]} ${year}`;
      }
      return `${day}-${parts[1]}-${year}`;
    }
    return raw;
  } catch {
    return isoStr.split("T")[0];
  }
};

/**
 * Join meal types with natural English conjunction (e.g. "Breakfast, Lunch & Dinner")
 */
export const joinMealTypes = (meals: MealType[]): string => {
  if (meals.length === 0) return "";
  if (meals.length === 1) return meals[0];
  if (meals.length === 2) return `${meals[0]} & ${meals[1]}`;
  return `${meals.slice(0, -1).join(", ")} & ${meals[meals.length - 1]}`;
};

/**
 * Service handling all SharePoint List operations for the SaveFood system.
 */
export class MealRequestService {
  private static _itemEntityTypeCache: { [listName: string]: string } = {};
  private static _listFieldsCache: { [listName: string]: any[] } = {};
  private static _cachedAllRequests: IPagedResult<IMealOptOutRequestItem> | null = null;
  private static _cachedAllRequestsTimestamp: number = 0;
  private static _cachedMyRequests: { [userId: number]: { data: IMealOptOutRequestItem[]; timestamp: number } } = {};
  private static _cachedLocations: ILocationMasterItem[] | null = null;
  private static _cachedLocationsTimestamp: number = 0;

  /**
   * Synchronously checks if My Requests data is already cached in memory
   */
  public static getCachedMyMealOptOuts(userId: number): IMealOptOutRequestItem[] | null {
    if (!userId || userId <= 0) return null;
    const entry = this._cachedMyRequests[userId];
    if (entry && Date.now() - entry.timestamp < 60000) {
      return entry.data;
    }
    return null;
  }

  /**
   * Synchronously checks if All Requests data is already cached in memory
   */
  public static getCachedAllMealOptOuts(): IMealOptOutRequestItem[] | null {
    if (this._cachedAllRequests && Date.now() - this._cachedAllRequestsTimestamp < 60000) {
      return this._cachedAllRequests.items;
    }
    return null;
  }

  /**
   * Clears in-memory caches when new records are written or deleted
   */
  public static clearCache(): void {
    this._cachedAllRequests = null;
    this._cachedAllRequestsTimestamp = 0;
    this._cachedMyRequests = {};
    this._cachedLocations = null;
    this._cachedLocationsTimestamp = 0;
  }

  /**
   * Retrieves the ListItemEntityTypeFullName dynamically from the list (with caching)
   */
  public static async getItemEntityType(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    listName: string
  ): Promise<string> {
    if (this._itemEntityTypeCache[listName]) {
      return this._itemEntityTypeCache[listName];
    }
    try {
      const response = await spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );
      if (response.ok) {
        const json = await response.json();
        const typeName = json?.d?.ListItemEntityTypeFullName;
        if (typeName) {
          this._itemEntityTypeCache[listName] = typeName;
          return typeName;
        }
      }
    } catch (e) {
      console.warn(`[MealRequestService] getItemEntityType fallback for ${listName}:`, e);
    }
    return `SP.Data.${listName}ListItem`;
  }

  /**
   * Retrieves list fields metadata to adapt to column types (Lookup vs Text vs User)
   */
  public static async getListFields(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    listName: string
  ): Promise<any[]> {
    if (this._listFieldsCache[listName]) {
      return this._listFieldsCache[listName];
    }
    try {
      const response = await spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,TypeAsString,Title`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );
      if (response.ok) {
        const json = await response.json();
        const fields = json?.d?.results || [];
        this._listFieldsCache[listName] = fields;
        return fields;
      }
    } catch (e) {
      console.warn(`[MealRequestService] getListFields fallback for ${listName}:`, e);
    }
    return [];
  }

  /**
   * Resolves the current user's SharePoint numeric ID from _api/web/currentuser
   */
  public static async resolveCurrentUserId(
    spHttpClient: SPHttpClient,
    siteUrl: string
  ): Promise<number> {
    try {
      const response = await spHttpClient.get(
        `${siteUrl}/_api/web/currentuser?$select=Id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );
      if (response.ok) {
        const json = await response.json();
        return json?.d?.Id ?? -1;
      }
    } catch {
      // fallback
    }
    return -1;
  }

  /**
   * 1. Fetch active dining locations from LocationMaster for request form dropdowns
   */
  public static async getActiveLocations(
    spHttpClient: SPHttpClient,
    siteUrl: string
  ): Promise<ILocationMasterItem[]> {
    try {
      const url = `${siteUrl}/_api/web/lists/getbytitle('LocationMaster')/items?$filter=IsActive eq 1&$select=Id,Title,IsActive&$orderby=Title asc&$top=4800`;
      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );

      if (!response.ok) {
        console.warn(`[MealRequestService] Failed to fetch LocationMaster: ${response.statusText}`);
        return [];
      }

      const data = await response.json();
      return (data?.d?.results || []).map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        IsActive: item.IsActive,
      }));
    } catch (error) {
      console.error("[MealRequestService] getActiveLocations error:", error);
      return [];
    }
  }

  /**
   * 2. Fetch all locations from LocationMaster (Admin view)
   */
  public static async getAllLocations(
    spHttpClient: SPHttpClient,
    siteUrl: string
  ): Promise<ILocationMasterItem[]> {
    if (this._cachedLocations && Date.now() - this._cachedLocationsTimestamp < 300000) {
      return this._cachedLocations;
    }

    try {
      const url = `${siteUrl}/_api/web/lists/getbytitle('LocationMaster')/items?$select=Id,Title,IsActive&$orderby=Title asc&$top=4800`;
      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );

      if (!response.ok) {
        return [];
      }

      const data = await response.json();
      const locations = (data?.d?.results || []).map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        IsActive: item.IsActive,
      }));

      this._cachedLocations = locations;
      this._cachedLocationsTimestamp = Date.now();
      return locations;
    } catch (error) {
      console.error("[MealRequestService] getAllLocations error:", error);
      return [];
    }
  }

  /**
   * 3. Toggle location active status (Admin only)
   */
  public static async updateLocationStatus(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    locationId: number,
    isActive: boolean
  ): Promise<boolean> {
    const url = `${siteUrl}/_api/web/lists/getbytitle('LocationMaster')/items(${locationId})`;
    const response = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-type": "application/json;odata=verbose",
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
        "odata-version": "",
      },
      body: JSON.stringify({ IsActive: isActive }),
    });
    return response.ok;
  }

  /**
   * 4. Add a new location to LocationMaster (Admin only)
   */
  public static async addLocation(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    title: string
  ): Promise<ILocationMasterItem | null> {
    const entityType = await this.getItemEntityType(spHttpClient, siteUrl, "LocationMaster");
    const url = `${siteUrl}/_api/web/lists/getbytitle('LocationMaster')/items`;
    const response = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-type": "application/json;odata=verbose",
        "odata-version": "",
      },
      body: JSON.stringify({
        __metadata: { type: entityType },
        Title: title.trim(),
        IsActive: true,
      }),
    });

    if (!response.ok) return null;
    const json = await response.json();
    return {
      Id: json.d.Id,
      Title: json.d.Title,
      IsActive: json.d.IsActive,
    };
  }

  /**
   * 5. Check if user already submitted an opt-out for this exact meal date, type & location
   */
  public static async checkDuplicateOptOut(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    userId: number,
    mealDateISO: string,
    mealType: MealType
  ): Promise<boolean> {
    try {
      if (!userId || userId <= 0) return false;
      const formattedDate = mealDateISO.split("T")[0];
      const url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items?$filter=Requester/Id eq ${userId} and MealType eq '${mealType}' and MealDate ge datetime'${formattedDate}T00:00:00Z' and MealDate le datetime'${formattedDate}T23:59:59Z'&$select=Id`;

      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );

      if (!response.ok) return false;
      const data = await response.json();
      return (data?.d?.results?.length || 0) > 0;
    } catch {
      return false;
    }
  }

  /**
   * 6. Create Meal Opt-Out Request with Double-Barrier Pre-Write Cut-Off Guard
   */
  public static async createMealOptOut(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    payload: ICreateMealOptOutPayload
  ): Promise<{ success: boolean; item?: IMealOptOutRequestItem; error?: string; isDuplicate?: boolean }> {
    // PRE-WRITE GUARD (Barrier 2): Server/data-write revalidation immediately before saving
    const cutOffCheck = validateMealCutOff(payload.MealDate, payload.MealType);

    if (!cutOffCheck.isValid) {
      return {
        success: false,
        error: cutOffCheck.errorMessage || "Submission closed for this meal cut-off.",
      };
    }

    // Check duplicate
    if (payload.RequesterId && payload.RequesterId > 0) {
      const isDuplicate = await this.checkDuplicateOptOut(
        spHttpClient,
        siteUrl,
        payload.RequesterId,
        payload.MealDate,
        payload.MealType
      );

      if (isDuplicate) {
        return {
          success: false,
          error: `You have already opted out of ${payload.MealType} on ${formatIsoToDisplayDate(payload.MealDate)}.`,
          isDuplicate: true,
        };
      }
    }

    try {
      const entityType = await this.getItemEntityType(spHttpClient, siteUrl, "MealOptOutRequest");
      const fields = await this.getListFields(spHttpClient, siteUrl, "MealOptOutRequest");

      // Verify user ID
      let requesterId = payload.RequesterId;
      if (!requesterId || requesterId <= 0) {
        requesterId = await this.resolveCurrentUserId(spHttpClient, siteUrl);
      }

      const body: any = {
        __metadata: { type: entityType },
        Title: payload.Title,
        MealType: payload.MealType,
        MealDate: payload.MealDate,
      };

      // Check how Location column was created in SharePoint list:
      const locationField = fields.find((f: any) => f.InternalName === "Location");
      if (locationField) {
        if (locationField.TypeAsString === "Lookup") {
          body.LocationId = payload.LocationId;
        } else {
          // Text or Choice column
          body.Location = payload.LocationTitle || String(payload.LocationId);
        }
      } else {
        // Fallback default: set LocationId for Lookup
        body.LocationId = payload.LocationId;
      }

      // Check how Requester column was created in SharePoint list:
      const requesterField = fields.find((f: any) => f.InternalName === "Requester");
      if (requesterField) {
        if (requesterField.TypeAsString === "User") {
          if (requesterId > 0) {
            body.RequesterId = requesterId;
          }
        } else {
          // Text column
          body.Requester = payload.RequesterEmail || String(requesterId);
        }
      } else {
        if (requesterId > 0) {
          body.RequesterId = requesterId;
        }
      }

      const url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items`;
      const response: SPHttpClientResponse = await spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-type": "application/json;odata=verbose",
            "odata-version": "",
          },
          body: JSON.stringify(body),
        }
      );

      if (!response.ok) {
        let errMsg = `HTTP error ${response.status}: ${response.statusText}`;
        try {
          const rawText = await response.text();
          try {
            const errorJson = JSON.parse(rawText);
            errMsg =
              errorJson?.error?.message?.value ||
              errorJson?.["odata.error"]?.message?.value ||
              errorJson?.message ||
              rawText;
          } catch {
            if (rawText) errMsg = rawText;
          }
        } catch {
          // ignore
        }
        console.error("[MealRequestService] createMealOptOut error:", errMsg, "Payload:", body);
        return { success: false, error: errMsg };
      }

      const json = await response.json();
      this.clearCache();
      return {
        success: true,
        item: json.d,
      };
    } catch (err: unknown) {
      return {
        success: false,
        error: err instanceof Error ? err.message : String(err),
      };
    }
  }

  /**
   * 6b. Create Multiple Meal Opt-Out Requests (one database entry per selected meal type)
   */
  public static async createMultipleMealOptOuts(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    payloads: ICreateMealOptOutPayload[]
  ): Promise<IBatchCreateResult> {
    const successfulMeals: MealType[] = [];
    const failedMeals: { mealType: MealType; error: string; isDuplicate?: boolean }[] = [];

    for (const payload of payloads) {
      const res = await this.createMealOptOut(spHttpClient, siteUrl, payload);
      if (res.success) {
        successfulMeals.push(payload.MealType);
      } else {
        failedMeals.push({
          mealType: payload.MealType,
          error: res.error || `Failed to save ${payload.MealType} opt-out`,
          isDuplicate: res.isDuplicate,
        });
      }
    }

    this.clearCache();

    const allSucceeded = failedMeals.length === 0;
    const anySucceeded = successfulMeals.length > 0;
    const formattedDate = formatIsoToDisplayDate(payloads[0]?.MealDate || "");

    let message = "";
    if (allSucceeded) {
      message = `Opt-out submitted successfully for ${joinMealTypes(successfulMeals)} on ${formattedDate}.`;
    } else {
      const allDuplicates = failedMeals.every((f) => f.isDuplicate);
      const duplicateMeals = failedMeals.filter((f) => f.isDuplicate).map((f) => f.mealType);
      const otherErrors = failedMeals.filter((f) => !f.isDuplicate).map((f) => `${f.mealType}: ${f.error}`);

      if (allDuplicates) {
        const dupText = `You have already opted out of ${joinMealTypes(duplicateMeals)} on ${formattedDate}.`;
        if (anySucceeded) {
          message = `Opt-out submitted successfully for ${joinMealTypes(successfulMeals)} on ${formattedDate}. (Already opted out of ${joinMealTypes(duplicateMeals)}.)`;
        } else {
          message = dupText;
        }
      } else {
        const parts: string[] = [];
        if (anySucceeded) {
          parts.push(`Opt-out submitted for ${joinMealTypes(successfulMeals)} on ${formattedDate}.`);
        }
        if (duplicateMeals.length > 0) {
          parts.push(`Already opted out of ${joinMealTypes(duplicateMeals)} on ${formattedDate}.`);
        }
        if (otherErrors.length > 0) {
          parts.push(otherErrors.join("; "));
        }
        message = parts.join(" ");
      }
    }

    return {
      success: anySucceeded,
      successfulMeals,
      failedMeals,
      message,
    };
  }

  /**
   * 7. Fetch current user's submitted requests (My Meal Opt-Outs)
   */
  public static async getMyMealOptOuts(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    currentUserId: number,
    currentUserEmail: string
  ): Promise<IMealOptOutRequestItem[]> {
    // Return cached requests immediately if fresh (< 60s)
    if (currentUserId > 0) {
      const cached = this.getCachedMyMealOptOuts(currentUserId);
      if (cached) {
        return cached;
      }
    }

    try {
      const filter = currentUserId > 0 ? `Requester/Id eq ${currentUserId}` : "";
      const filterQuery = filter ? `?$filter=${filter}&` : "?";
      const url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items${filterQuery}$select=Id,Title,MealType,MealDate,Created,Location/Id,Location/Title,Requester/Id,Requester/Title,Requester/EMail&$expand=Location,Requester&$orderby=MealDate desc&$top=4800`;

      let response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );

      // Graceful fallback if $expand is invalid (e.g. if Location or Requester were created as Text)
      if (!response.ok) {
        const fallbackUrl = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items?$orderby=MealDate desc&$top=4800`;
        response = await spHttpClient.get(
          fallbackUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=verbose",
              "odata-version": "",
            },
          }
        );
      }

      if (!response.ok) {
        console.warn(`[MealRequestService] getMyMealOptOuts error: ${response.statusText}`);
        return [];
      }

      const data = await response.json();
      const results: IMealOptOutRequestItem[] = (data?.d?.results || []).map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        MealType: item.MealType,
        MealDate: item.MealDate,
        LocationId: item.Location?.Id ?? (typeof item.Location === "number" ? item.Location : 0),
        Location: item.Location
          ? (typeof item.Location === "object" ? { Id: item.Location.Id, Title: item.Location.Title } : { Id: 0, Title: String(item.Location) })
          : undefined,
        RequesterId: item.Requester?.Id ?? (typeof item.Requester === "number" ? item.Requester : 0),
        Requester: item.Requester
          ? (typeof item.Requester === "object" ? { Id: item.Requester.Id, Title: item.Requester.Title, EMail: item.Requester.EMail } : { Id: 0, Title: String(item.Requester) })
          : undefined,
        Created: item.Created,
      }));

      if (currentUserId > 0) {
        this._cachedMyRequests[currentUserId] = {
          data: results,
          timestamp: Date.now(),
        };
      }

      return results;
    } catch (error) {
      console.error("[MealRequestService] getMyMealOptOuts exception:", error);
      return [];
    }
  }

  /**
   * 8. Fetch all opt-out requests with filtering and pagination (Admin only)
   */
  public static async getAllMealOptOuts(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    options: {
      dateISO?: string | null; // e.g. "2026-11-18" or null for "All Dates"
      locationId?: number | null; // null for all locations
      nextUrl?: string | null;
      pageSize?: number;
    }
  ): Promise<IPagedResult<IMealOptOutRequestItem>> {
    const isFullQuery = !options.nextUrl && !options.dateISO && !options.locationId;
    if (isFullQuery && this._cachedAllRequests && Date.now() - this._cachedAllRequestsTimestamp < 30000) {
      return this._cachedAllRequests;
    }

    try {
      let url = options.nextUrl;

      if (!url) {
        const filters: string[] = [];

        if (options.dateISO) {
          const dateOnly = options.dateISO.split("T")[0];
          filters.push(`MealDate ge datetime'${dateOnly}T00:00:00Z' and MealDate le datetime'${dateOnly}T23:59:59Z'`);
        }

        if (options.locationId && options.locationId > 0) {
          filters.push(`Location/Id eq ${options.locationId}`);
        }

        const filterQuery = filters.length > 0 ? `&$filter=${filters.join(" and ")}` : "";
        const top = options.pageSize || 4800;

        url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items?$select=Id,Title,MealType,MealDate,Created,Location/Id,Location/Title,Requester/Id,Requester/Title,Requester/EMail&$expand=Location,Requester&$orderby=MealDate desc&$top=${top}${filterQuery}`;
      }

      let response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );

      // Fallback without $expand if needed
      if (!response.ok && !options.nextUrl) {
        const top = options.pageSize || 4800;
        const fallbackUrl = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items?$orderby=MealDate desc&$top=${top}`;
        response = await spHttpClient.get(
          fallbackUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=verbose",
              "odata-version": "",
            },
          }
        );
      }

      if (!response.ok) {
        return { items: [], nextUrl: null };
      }

      const data = await response.json();
      const rawResults = data?.d?.results || [];
      const nextLink = data?.d?.__next || null;

      const items: IMealOptOutRequestItem[] = rawResults.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        MealType: item.MealType,
        MealDate: item.MealDate,
        LocationId: item.Location?.Id ?? (typeof item.Location === "number" ? item.Location : 0),
        Location: item.Location
          ? (typeof item.Location === "object" ? { Id: item.Location.Id, Title: item.Location.Title } : { Id: 0, Title: String(item.Location) })
          : undefined,
        RequesterId: item.Requester?.Id ?? (typeof item.Requester === "number" ? item.Requester : 0),
        Requester: item.Requester
          ? (typeof item.Requester === "object" ? { Id: item.Requester.Id, Title: item.Requester.Title, EMail: item.Requester.EMail } : { Id: 0, Title: String(item.Requester) })
          : undefined,
        Created: item.Created,
      }));

      const result: IPagedResult<IMealOptOutRequestItem> = {
        items,
        nextUrl: nextLink,
      };

      if (isFullQuery) {
        this._cachedAllRequests = result;
        this._cachedAllRequestsTimestamp = Date.now();
      }

      return result;
    } catch (error) {
      console.error("[MealRequestService] getAllMealOptOuts error:", error);
      return { items: [], nextUrl: null };
    }
  }

  /**
   * 9. Cancel / Delete an Opt-Out Request before cut-off
   */
  public static async cancelOptOut(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    item: IMealOptOutRequestItem,
    context: WebPartContext
  ): Promise<{ success: boolean; error?: string }> {
    // Check retraction cut-off (10 min buffer) before permitting cancellation
    const cutOffCheck = validateRetractionCutOff(item.MealDate, item.MealType);

    if (!cutOffCheck.isValid) {
      return {
        success: false,
        error: cutOffCheck.errorMessage || `Cannot cancel opt-out. Cut-off time for this meal has already passed.`,
      };
    }

    try {
      const url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items(${item.Id})`;
      const response = await spHttpClient.fetch(url, SPHttpClient.configurations.v1, {
        method: "POST",
        headers: {
          Accept: "*/*",
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*",
          "X-RequestDigest": context.pageContext.legacyPageContext.formDigestValue,
        },
      });

      if (response.status === 204 || response.status === 200 || response.status === 404) {
        this.clearCache();
        return { success: true };
      }

      return {
        success: false,
        error: `Failed to cancel request (HTTP ${response.status}).`,
      };
    } catch (err: unknown) {
      return {
        success: false,
        error: err instanceof Error ? err.message : String(err),
      };
    }
  }
}
