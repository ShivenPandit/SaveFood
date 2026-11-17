import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ILocationMasterItem } from "../models/ILocation";
import {
  IMealOptOutRequestItem,
  ICreateMealOptOutPayload,
  MealType,
} from "../models/IMealOptOut";
import { validateMealCutOff } from "./cutOffService";

export interface IPagedResult<T> {
  items: T[];
  nextUrl: string | null;
  totalCount?: number;
}

/**
 * Service handling all SharePoint List operations for the Save Food system.
 */
export class MealRequestService {
  /**
   * 1. Fetch active dining locations from LocationMaster for request form dropdowns
   */
  public static async getActiveLocations(
    spHttpClient: SPHttpClient,
    siteUrl: string
  ): Promise<ILocationMasterItem[]> {
    try {
      const url = `${siteUrl}/_api/web/lists/getbytitle('LocationMaster')/items?$filter=IsActive eq 1&$select=Id,Title,IsActive&$orderby=Title asc`;
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
    try {
      const url = `${siteUrl}/_api/web/lists/getbytitle('LocationMaster')/items?$select=Id,Title,IsActive&$orderby=Title asc`;
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
      return (data?.d?.results || []).map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        IsActive: item.IsActive,
      }));
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
    const url = `${siteUrl}/_api/web/lists/getbytitle('LocationMaster')/items`;
    const response = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-type": "application/json;odata=verbose",
      },
      body: JSON.stringify({ Title: title.trim(), IsActive: true }),
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
      const formattedDate = mealDateISO.split("T")[0];
      const url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items?$filter=Requester/Id eq ${userId} and MealType eq '${mealType}' and MealDate ge datetime'${formattedDate}T00:00:00Z' and MealDate le datetime'${formattedDate}T23:59:59Z'&$select=Id`;

      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
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
  ): Promise<{ success: boolean; item?: IMealOptOutRequestItem; error?: string }> {
    // PRE-WRITE GUARD (Barrier 2): Server/data-write revalidation immediately before saving
    const mealDateObj = new Date(payload.MealDate);
    const cutOffCheck = validateMealCutOff(mealDateObj, payload.MealType);

    if (!cutOffCheck.isValid) {
      return {
        success: false,
        error: cutOffCheck.errorMessage || "Submission closed for this meal cut-off.",
      };
    }

    // Check duplicate
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
        error: `You have already opted out of ${payload.MealType} on ${payload.MealDate.split("T")[0]}.`,
      };
    }

    try {
      const url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items`;
      const body = {
        Title: payload.Title,
        LocationId: payload.LocationId,
        MealType: payload.MealType,
        MealDate: payload.MealDate,
        RequesterId: payload.RequesterId,
      };

      const response: SPHttpClientResponse = await spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-type": "application/json;odata=verbose",
          },
          body: JSON.stringify(body),
        }
      );

      if (!response.ok) {
        const errorJson = await response.json().catch(() => null);
        const errMsg = errorJson?.error?.message?.value || `HTTP error ${response.status}: ${response.statusText}`;
        return { success: false, error: errMsg };
      }

      const json = await response.json();
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
   * 7. Fetch current user's submitted requests (My Meal Opt-Outs)
   */
  public static async getMyMealOptOuts(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    currentUserId: number,
    currentUserEmail: string
  ): Promise<IMealOptOutRequestItem[]> {
    try {
      // Query filtered by user ID on indexed Requester column
      const filter = `Requester/Id eq ${currentUserId}`;
      const url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items?$filter=${filter}&$select=Id,Title,MealType,MealDate,Created,Location/Id,Location/Title,Requester/Id,Requester/Title,Requester/EMail&$expand=Location,Requester&$orderby=MealDate desc&$top=500`;

      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );

      if (!response.ok) {
        console.warn(`[MealRequestService] getMyMealOptOuts error: ${response.statusText}`);
        return [];
      }

      const data = await response.json();
      return (data?.d?.results || []).map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        MealType: item.MealType,
        MealDate: item.MealDate,
        LocationId: item.Location?.Id ?? 0,
        Location: item.Location ? { Id: item.Location.Id, Title: item.Location.Title } : undefined,
        RequesterId: item.Requester?.Id ?? 0,
        Requester: item.Requester ? { Id: item.Requester.Id, Title: item.Requester.Title, EMail: item.Requester.EMail } : undefined,
        Created: item.Created,
      }));
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
        const top = options.pageSize || 50;

        url = `${siteUrl}/_api/web/lists/getbytitle('MealOptOutRequest')/items?$select=Id,Title,MealType,MealDate,Created,Location/Id,Location/Title,Requester/Id,Requester/Title,Requester/EMail&$expand=Location,Requester&$orderby=MealDate desc&$top=${top}${filterQuery}`;
      }

      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );

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
        LocationId: item.Location?.Id ?? 0,
        Location: item.Location ? { Id: item.Location.Id, Title: item.Location.Title } : undefined,
        RequesterId: item.Requester?.Id ?? 0,
        Requester: item.Requester ? { Id: item.Requester.Id, Title: item.Requester.Title, EMail: item.Requester.EMail } : undefined,
        Created: item.Created,
      }));

      return {
        items,
        nextUrl: nextLink,
      };
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
    // Check cut-off before permitting cancellation
    const mealDateObj = new Date(item.MealDate);
    const cutOffCheck = validateMealCutOff(mealDateObj, item.MealType);

    if (!cutOffCheck.isValid) {
      return {
        success: false,
        error: `Cannot cancel opt-out. Cut-off time for this meal has already passed.`,
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
