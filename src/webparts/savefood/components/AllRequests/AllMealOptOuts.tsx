import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem, MealType } from "../../models/IMealOptOut";
import { ILocationMasterItem } from "../../models/ILocation";
import { MealRequestService } from "../../services/mealRequestService";
import { formatISTDateToISO } from "../../services/cutOffService";
import { BackButton } from "../Common/BackButton";
import { useDynamicPagination } from "../../hooks/useDynamicPagination";
import { DataTablePagination } from "../Common/DataTablePagination";
import { AdvancedFilterPopover } from "../Common/AdvancedFilterPopover";

export interface IAllMealOptOutsProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  onBack: () => void;
  hideBackButton?: boolean;
  onFilteredDataChange?: (filteredItems: IMealOptOutRequestItem[], hasActiveFilters: boolean) => void;
  // Shared Filter State (Optional, for syncing with ReportsDashboard)
  sharedStartDate?: string | null;
  sharedEndDate?: string | null;
  onSharedStartDateChange?: (val: string | null) => void;
  onSharedEndDateChange?: (val: string | null) => void;
  sharedLocationIds?: number[];
  onSharedLocationIdsChange?: (ids: number[]) => void;
  sharedMealTypes?: MealType[];
  onSharedMealTypesChange?: (types: MealType[]) => void;
}

export const AllMealOptOuts: React.FC<IAllMealOptOutsProps> = ({
  context,
  spHttpClient,
  siteUrl,
  onBack,
  hideBackButton = false,
  onFilteredDataChange,
  sharedStartDate,
  sharedEndDate,
  onSharedStartDateChange,
  onSharedEndDateChange,
  sharedLocationIds,
  onSharedLocationIdsChange,
  sharedMealTypes,
  onSharedMealTypesChange,
}) => {
  const cachedInitial = React.useMemo(() => {
    return MealRequestService.getCachedAllMealOptOuts() || [];
  }, []);

  const [requests, setRequests] = React.useState<IMealOptOutRequestItem[]>(cachedInitial);
  const [locations, setLocations] = React.useState<ILocationMasterItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(cachedInitial.length === 0);

  // Date Presets & Custom Range Filter (Defaults to Today in IST)
  const todayISOString = React.useMemo(() => formatISTDateToISO(new Date()), []);

  // Controlled vs Uncontrolled Filter State
  const isControlled = sharedStartDate !== undefined;

  const [internalStartDate, setInternalStartDate] = React.useState<string | null>(todayISOString);
  const [internalEndDate, setInternalEndDate] = React.useState<string | null>(todayISOString);
  const [internalLocationIds, setInternalLocationIds] = React.useState<number[]>([]);
  const [internalMealTypes, setInternalMealTypes] = React.useState<MealType[]>([]);

  const startDate = isControlled ? sharedStartDate : internalStartDate;
  const endDate = isControlled ? sharedEndDate : internalEndDate;
  const setStartDate = isControlled && onSharedStartDateChange ? onSharedStartDateChange : setInternalStartDate;
  const setEndDate = isControlled && onSharedEndDateChange ? onSharedEndDateChange : setInternalEndDate;

  const selectedLocationIds = isControlled && sharedLocationIds ? sharedLocationIds : internalLocationIds;
  const setSelectedLocationIds = isControlled && onSharedLocationIdsChange ? onSharedLocationIdsChange : setInternalLocationIds;

  const selectedMealTypes = isControlled && sharedMealTypes ? sharedMealTypes : internalMealTypes;
  const setSelectedMealTypes = isControlled && onSharedMealTypesChange ? onSharedMealTypesChange : setInternalMealTypes;

  // Filter Popover state
  const [isFilterOpen, setIsFilterOpen] = React.useState<boolean>(false);
  const filterRef = React.useRef<HTMLDivElement>(null);

  // Active filter count for badge
  const isAllDatesActive = startDate === null && endDate === null;
  const isLocationFilterActive = selectedLocationIds.length > 0;
  const isMealTypeFilterActive = selectedMealTypes.length > 0 && selectedMealTypes.length < 3;

  const activeFilterCount = React.useMemo(() => {
    let count = 0;
    if (!isAllDatesActive) count++;
    if (isLocationFilterActive) count++;
    if (isMealTypeFilterActive) count++;
    return count;
  }, [isAllDatesActive, isLocationFilterActive, isMealTypeFilterActive]);

  const handleResetFilters = (): void => {
    setStartDate(null);
    setEndDate(null);
    setSelectedLocationIds([]);
    setSelectedMealTypes([]);
  };

  // Pagination
  const [nextUrl, setNextUrl] = React.useState<string | null>(null);
  const [isLoadingMore, setIsLoadingMore] = React.useState<boolean>(false);

  // 1. Load Locations on mount
  React.useEffect(() => {
    MealRequestService.getAllLocations(spHttpClient, siteUrl)
      .then((locs) => {
        setLocations(locs);
      })
      .catch((err: Error) => {
        console.error(err);
      });
  }, [spHttpClient, siteUrl]);

  // 2. Fetch Requests once with high page limit (in-memory filtering for instant UI response)
  const fetchRequests = React.useCallback(
    async (silent: boolean = false) => {
      if (!silent) {
        setIsLoading(true);
      }

      const paged = await MealRequestService.getAllMealOptOuts(spHttpClient, siteUrl, {
        dateISO: null,
        locationId: null,
        pageSize: 4800,
      });

      setRequests(paged.items);
      setNextUrl(paged.nextUrl);
      setIsLoading(false);
    },
    [spHttpClient, siteUrl]
  );

  React.useEffect(() => {
    const hasCache = cachedInitial.length > 0;
    fetchRequests(hasCache).catch((err: Error) => {
      console.error(err);
      setIsLoading(false);
    });
  }, [fetchRequests, cachedInitial]);

  // Load next page if needed
  const handleLoadMore = async (): Promise<void> => {
    if (!nextUrl || isLoadingMore) return;
    setIsLoadingMore(true);

    const paged = await MealRequestService.getAllMealOptOuts(spHttpClient, siteUrl, {
      nextUrl,
    });

    setRequests((prev) => [...prev, ...paged.items]);
    setNextUrl(paged.nextUrl);
    setIsLoadingMore(false);
  };

  // Instant in-memory filtering: Date Range, Location Multi-Select, Meal Type Multi-Select, and Requesters
  const filteredRequests = React.useMemo(() => {
    const effStart = startDate && endDate && startDate > endDate ? endDate : startDate;
    const effEnd = startDate && endDate && startDate > endDate ? startDate : endDate;

    return requests.filter((r) => {
      // 1. Date Range filter (exact YYYY-MM-DD string match, immune to UTC timezone shift)
      const itemDateStr = r.MealDate ? r.MealDate.split("T")[0] : "";
      if (effStart && itemDateStr < effStart) return false;
      if (effEnd && itemDateStr > effEnd) return false;

      // 2. Meal Type Multi-Select filter
      if (selectedMealTypes.length > 0 && selectedMealTypes.length < 3) {
        if (selectedMealTypes.indexOf(r.MealType) === -1) return false;
      }

      // 3. Location Multi-Select filter
      if (selectedLocationIds.length > 0) {
        const itemLocId = r.LocationId || r.Location?.Id;
        if (!itemLocId || selectedLocationIds.indexOf(itemLocId) === -1) return false;
      }

      return true;
    });
  }, [requests, startDate, endDate, selectedLocationIds, selectedMealTypes]);

  // Dynamic Viewport-Aware Pagination
  const {
    containerRef,
    currentPage,
    setCurrentPage,
    totalPages,
    totalItems,
    paginatedItems,
    startIndex,
    endIndex,
  } = useDynamicPagination<IMealOptOutRequestItem>({
    items: filteredRequests,
    viewType: "all-requests",
  });

  // Reset to first page when filters change
  React.useEffect(() => {
    setCurrentPage(1);
  }, [startDate, endDate, selectedLocationIds, selectedMealTypes, setCurrentPage]);

  // Notify parent component of current filtered dataset and active filter status
  React.useEffect(() => {
    if (onFilteredDataChange) {
      onFilteredDataChange(filteredRequests, activeFilterCount > 0);
    }
  }, [filteredRequests, activeFilterCount, onFilteredDataChange]);

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header */}
      <div className="mb-2 sm:mb-2.5 flex flex-col md:flex-row md:items-center md:justify-between gap-2.5">
        <div>
          {!hideBackButton && <BackButton onClick={onBack} className="mb-1" />}
          <div className="flex items-center gap-2">
            <h2 className="text-[22px] sm:text-[24px] font-bold text-[#0f172a] tracking-tight">
              All Meal Opt-Out Requests
            </h2>
          </div>
          <p className="text-[13.5px] text-[#475569] leading-snug mt-0.5">
            Consolidated kitchen catering headcount roster and location-wise opt-out records.
          </p>
        </div>

        {/* Action Controls: Filter Button & Records Count Pill */}
        <div className="flex items-center gap-2.5 flex-wrap self-start md:self-auto">
          {/* Filters Toggle Button & Dropdown Popover */}
          <div ref={filterRef} className="relative">
            <button
              type="button"
              onClick={() => setIsFilterOpen((prev) => !prev)}
              aria-expanded={isFilterOpen}
              className={`h-[40px] px-3.5 rounded-xl border inline-flex items-center gap-2 text-[14px] font-semibold transition cursor-pointer select-none whitespace-nowrap ${
                isFilterOpen || activeFilterCount > 0
                  ? "bg-blue-50/80 border-[#004B87] text-[#004B87] shadow-2xs ring-1 ring-[#004B87]/20"
                  : "bg-white border-slate-300 text-slate-700 hover:bg-slate-50"
              }`}
            >
              <svg viewBox="0 0 24 24" className="w-4 h-4 shrink-0 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
              </svg>
              <span>Filters</span>
              {activeFilterCount > 0 && (
                <span className="w-5 h-5 rounded-full bg-[#004B87] text-white text-[11px] font-bold flex items-center justify-center shrink-0">
                  {activeFilterCount}
                </span>
              )}
              <svg
                viewBox="0 0 24 24"
                className={`w-3.5 h-3.5 shrink-0 fill-none stroke-current transition-transform duration-200 ${
                  isFilterOpen ? "rotate-180" : ""
                }`}
                strokeWidth="2.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <polyline points="6 9 12 15 18 9" />
              </svg>
            </button>

            {/* Advanced Search Criteria Dropdown Popover */}
            <AdvancedFilterPopover
              isOpen={isFilterOpen}
              onClose={() => setIsFilterOpen(false)}
              startDate={startDate || null}
              endDate={endDate || null}
              onStartDateChange={setStartDate}
              onEndDateChange={setEndDate}
              selectedMealTypes={selectedMealTypes}
              onMealTypesChange={setSelectedMealTypes}
              locations={locations}
              selectedLocationIds={selectedLocationIds}
              onLocationsChange={setSelectedLocationIds}
              onResetAll={handleResetFilters}
              activeFilterCount={activeFilterCount}
            />
          </div>

          {/* Count Pill */}
          <div className="flex items-center gap-1.5 bg-slate-100 px-3.5 py-2 rounded-xl text-[13px] font-semibold text-slate-700">
            <span>Total Records:</span>
            <span className="text-[#004B87] font-bold text-[14px]">{filteredRequests.length}</span>
          </div>
        </div>
      </div>

      {/* Roster Table Card */}
      <div
        ref={containerRef}
        style={{ height: "auto", flex: "0 0 auto" }}
        className="bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm overflow-hidden"
      >
        {isLoading ? (
          <div className="p-12 text-center">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-4 border-[#004B87] border-t-transparent" />
            <p className="mt-3 text-xs sm:text-sm font-semibold text-slate-500">
              Querying catering roster...
            </p>
          </div>
        ) : filteredRequests.length === 0 ? (
          <div className="p-12 text-center">
            <div className="w-14 h-14 mx-auto rounded-full bg-slate-100 flex items-center justify-center text-slate-400 mb-3">
              <svg viewBox="0 0 24 24" className="w-7 h-7 fill-none stroke-current" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                <line x1="16" y1="2" x2="16" y2="6" />
                <line x1="8" y1="2" x2="8" y2="6" />
                <line x1="3" y1="10" x2="21" y2="10" />
              </svg>
            </div>
            <h3 className="text-base font-bold text-slate-800">No Opt-Outs for this Filter</h3>
            <p className="text-xs sm:text-sm text-slate-500 max-w-sm mx-auto mt-1">
              No employee meal cancellations recorded for the selected criteria.
            </p>
            {activeFilterCount > 0 && (
              <button
                type="button"
                onClick={handleResetFilters}
                className="mt-3.5 inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-[#004B87] hover:bg-[#003865] text-white text-xs font-bold transition shadow-2xs cursor-pointer"
              >
                <span>Reset All Filters</span>
              </button>
            )}
          </div>
        ) : (
          <>
            {/* Desktop Table View (>= 640px) */}
            <div className="hidden sm:block overflow-x-auto">
              <table className="w-full text-left text-[14.5px]">
                <thead className="bg-slate-50 border-b border-slate-200 text-[#475569] font-bold uppercase tracking-wider text-[12px] sm:text-[12.5px]">
                  <tr>
                    <th className="py-2.5 sm:py-2.5 px-4 sm:px-5">Meal Date</th>
                    <th className="py-2.5 sm:py-2.5 px-4">Meal Type</th>
                    <th className="py-2.5 sm:py-2.5 px-4">Location</th>
                    <th className="py-2.5 sm:py-2.5 px-4">Employee / Requester</th>
                    <th className="py-2.5 sm:py-2.5 px-4">Email</th>
                    <th className="py-2.5 sm:py-2.5 px-4 sm:px-5 text-right">Submitted Timestamp</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-medium text-slate-700">
                  {paginatedItems.map((item) => {
                    const mealDateObj = new Date(item.MealDate);
                    const formattedMealDate = mealDateObj.toLocaleDateString("en-IN", {
                      weekday: "short",
                      day: "2-digit",
                      month: "short",
                      year: "numeric",
                    });

                    const createdObj = item.Created ? new Date(item.Created) : null;
                    const formattedCreated = createdObj
                      ? createdObj.toLocaleDateString("en-IN", {
                          day: "2-digit",
                          month: "short",
                          hour: "2-digit",
                          minute: "2-digit",
                        })
                      : "—";

                    return (
                      <tr key={item.Id} className="hover:bg-slate-50/80 transition-colors">
                        <td className="py-2.5 sm:py-2.5 px-4 sm:px-5 font-bold text-slate-900 whitespace-nowrap">
                          {formattedMealDate}
                        </td>
                        <td className="py-2.5 sm:py-2.5 px-4 whitespace-nowrap">
                          <span
                            className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-[12px] font-semibold ${
                              item.MealType === "Breakfast"
                                ? "bg-amber-50 text-amber-800 border border-amber-200"
                                : item.MealType === "Lunch"
                                ? "bg-blue-50 text-blue-800 border border-blue-200"
                                : "bg-purple-50 text-purple-800 border border-purple-200"
                            }`}
                          >
                            <span>{item.MealType === "Breakfast" ? "🍳" : item.MealType === "Lunch" ? "🥗" : "🍲"}</span>
                            <span className="ml-1">{item.MealType}</span>
                          </span>
                        </td>
                        <td className="py-2.5 sm:py-2.5 px-4 font-semibold text-slate-800 whitespace-nowrap">
                          {item.Location?.Title || "—"}
                        </td>
                        <td className="py-2.5 sm:py-2.5 px-4 font-bold text-slate-900 whitespace-nowrap">
                          {item.Requester?.Title || "Employee"}
                        </td>
                        <td className="py-2.5 sm:py-2.5 px-4 text-slate-500 font-mono text-[12.5px] whitespace-nowrap">
                          {item.Requester?.EMail || "—"}
                        </td>
                        <td className="py-2.5 sm:py-2.5 px-4 sm:px-5 text-right text-slate-500 whitespace-nowrap text-[13px]">
                          {formattedCreated}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Mobile Card View (< 640px) */}
            <div className="sm:hidden divide-y divide-slate-100">
              {paginatedItems.map((item) => {
                const mealDateObj = new Date(item.MealDate);
                const formattedMealDate = mealDateObj.toLocaleDateString("en-IN", {
                  weekday: "short",
                  day: "2-digit",
                  month: "short",
                  year: "numeric",
                });

                const createdObj = item.Created ? new Date(item.Created) : null;
                const formattedCreated = createdObj
                  ? createdObj.toLocaleDateString("en-IN", {
                      day: "2-digit",
                      month: "short",
                      hour: "2-digit",
                      minute: "2-digit",
                    })
                  : "—";

                return (
                  <div
                    key={item.Id}
                    data-role="list-item"
                    className="p-4 hover:bg-slate-50/70 transition-colors flex flex-col gap-2.5"
                  >
                    <div className="flex items-center justify-between">
                      <span className="font-bold text-slate-900 text-[15px]">
                        {formattedMealDate}
                      </span>
                      <span
                        className={`inline-flex items-center px-3 py-1 rounded-full text-[13px] font-semibold ${
                          item.MealType === "Breakfast"
                            ? "bg-amber-50 text-amber-800 border border-amber-200"
                            : item.MealType === "Lunch"
                            ? "bg-blue-50 text-blue-800 border border-blue-200"
                            : "bg-purple-50 text-purple-800 border border-purple-200"
                        }`}
                      >
                        <span>{item.MealType === "Breakfast" ? "🍳" : item.MealType === "Lunch" ? "🥗" : "🍲"}</span>
                        <span className="ml-1">{item.MealType}</span>
                      </span>
                    </div>

                    <div className="flex items-center justify-between text-[14px]">
                      <span className="font-bold text-slate-900">
                        {item.Requester?.Title || "Employee"}
                      </span>
                      <span className="text-slate-600 font-semibold">
                        📍 {item.Location?.Title || "—"}
                      </span>
                    </div>

                    <div className="flex items-center justify-between text-[11px] text-slate-400 font-mono">
                      <span className="truncate max-w-[200px]">{item.Requester?.EMail || "—"}</span>
                      <span>{formattedCreated}</span>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Dynamic Pagination Footer */}
            <DataTablePagination
              currentPage={currentPage}
              totalPages={totalPages}
              totalItems={totalItems}
              startIndex={startIndex}
              endIndex={endIndex}
              onPageChange={setCurrentPage}
            />
          </>
        )}

        {/* Load More Pagination (for subsequent remote batches) */}
        {nextUrl && (
          <div className="p-3 bg-slate-50/80 border-t border-slate-200 text-center">
            <button
              type="button"
              onClick={handleLoadMore}
              disabled={isLoadingMore}
              className="px-5 py-2 rounded-full bg-[#004B87] hover:bg-[#003865] text-white text-xs font-bold transition cursor-pointer disabled:opacity-50"
            >
              {isLoadingMore ? "Loading Next Batch..." : "Fetch More Records From Server"}
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
