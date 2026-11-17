/**
 * Model representing a Canteen / Dining Location item in LocationMaster
 */
export interface ILocationMasterItem {
  Id: number;
  Title: string; // Location name (e.g. LQ, MCR, Project Office)
  IsActive: boolean;
  Created?: string;
  Modified?: string;
}
