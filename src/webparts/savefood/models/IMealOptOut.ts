export type MealType = "Breakfast" | "Lunch" | "Dinner";

/**
 * Requester Person field representation in MealOptOutRequest
 */
export interface IRequesterPerson {
  Id: number;
  Title: string;
  EMail?: string;
  UserName?: string;
}

/**
 * Location Lookup field representation in MealOptOutRequest
 */
export interface ILocationLookup {
  Id: number;
  Title: string;
}

/**
 * Transactional Meal Opt-Out Request item in MealOptOutRequest list
 */
export interface IMealOptOutRequestItem {
  Id: number;
  Title: string; // Request Summary (e.g. "LQ - Lunch - 2026-11-18")
  LocationId: number;
  Location?: ILocationLookup;
  MealType: MealType;
  MealDate: string; // ISO date string (YYYY-MM-DD or YYYY-MM-DDT00:00:00Z)
  RequesterId: number;
  Requester?: IRequesterPerson;
  Created: string;
  Modified?: string;
}

/**
 * Payload for creating a new Meal Opt-Out Request item
 */
export interface ICreateMealOptOutPayload {
  Title: string;
  LocationId: number;
  LocationTitle?: string;
  MealType: MealType;
  MealDate: string;
  RequesterId: number;
  RequesterEmail?: string;
}
