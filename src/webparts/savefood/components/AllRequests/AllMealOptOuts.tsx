import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem } from "../../models/IMealOptOut";
import { ILocationMasterItem } from "../../models/ILocation";
import { MealRequestService } from "../../services/mealRequestService";
import { formatISTDateToISO } from "../../services/cutOffService";
import { BackButton } from "../Common/BackButton";
import TagPeoplePicker, { IUser } from "../Common/TagPeoplePicker";
import { useDynamicPagination } from "../../hooks/useDynamicPagination";
import { DataTablePagination } from "../Common/DataTablePagination";

export interface IAllMealOptOutsProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  onBack: () => void;
  hideBackButton?: boolean;
  onFilteredDataChange?: (filteredItems: IMealOptOutRequestItem[], hasActiveFilters: boolean) => void;
}

export const AllMealOptOuts: React.FC<IAllMealOptOutsProps> = ({
  context,
  spHttpClient,
  siteUrl,
  onBack,
  hideBackButton = false,
  onFilteredDataChange,
}) => {
  const cachedInitial = React.useMemo(() => {
    return MealRequestService.getCachedAllMealOptOuts() || [];
  }, []);

  const [requests, setRequests] = React.useState<IMealOptOutRequestItem[]>(cachedInitial);
  const [locations, setLocations] = React.useState<ILocationMasterItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(cachedInitial.length === 0);

  // Date Presets & Custom Range Filter (Defaults to Today in IST)
  const todayISOString = React.useMemo(() => formatISTDateToISO(new Date()), []);
  const tomorrowISOString = React.useMemo(() => {
    const d = new Date();
    d.setDate(d.getDate() + 1);
    return formatISTDateToISO(d);
  }, []);

  // Default to Today on initial load as requested
  const [startDate, setStartDate] = React.useState<string | null>(todayISOString);
  const [endDate, setEndDate] = React.useState<string | null>(todayISOString);
  const [filterLocationId, setFilterLocationId] = React.useState<number | "all">("all");
  const [selectedRequesters, setSelectedRequesters] = React.useState<IUser[]>([]);

  // Preset Active Flags
  const isTodayActive = startDate === todayISOString && endDate === todayISOString;
  const isTomorrowActive = startDate === tomorrowISOString && endDate === tomorrowISOString;
  const isAllDatesActive = startDate === null && endDate === null;
  const isCustomRange =
    !isTodayActive &&
    !isTomorrowActive &&
    !isAllDatesActive &&
    (Boolean(startDate) || Boolean(endDate));

  // Date Range Change Handlers (guarantees To >= From)
  const handleStartDateChange = (val: string | null): void => {
    setStartDate(val);
    if (val && endDate && val > endDate) {
      setEndDate(val);
    }
  };

  const handleEndDateChange = (val: string | null): void => {
    if (val && startDate && val < startDate) {
      setEndDate(startDate);
    } else {
      setEndDate(val);
    }
  };

  // Filter Popover state & Click Outside Handler
  const [isFilterOpen, setIsFilterOpen] = React.useState<boolean>(false);
  const filterRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (filterRef.current && !filterRef.current.contains(event.target as Node)) {
        setIsFilterOpen(false);
      }
    };
    if (isFilterOpen) {
      document.addEventListener("mousedown", handleClickOutside);
    }
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [isFilterOpen]);

  // Active filter count for badge
  const activeFilterCount = React.useMemo(() => {
    let count = 0;
    if (!isAllDatesActive) count++;
    if (filterLocationId !== "all") count++;
    if (selectedRequesters.length > 0) count++;
    return count;
  }, [isAllDatesActive, filterLocationId, selectedRequesters]);

  const handleResetFilters = (): void => {
    setStartDate(null);
    setEndDate(null);
    setFilterLocationId("all");
    setSelectedRequesters([]);
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

  // Instant in-memory filtering: Date Range, Location, and Multi-Person Requesters
  const filteredRequests = React.useMemo(() => {
    const effStart = startDate && endDate && startDate > endDate ? endDate : startDate;
    const effEnd = startDate && endDate && startDate > endDate ? startDate : endDate;

    return requests.filter((r) => {
      // 1. Date Range filter (exact YYYY-MM-DD string match, immune to UTC timezone shift)
      const itemDateStr = r.MealDate ? r.MealDate.split("T")[0] : "";
      if (effStart && itemDateStr < effStart) return false;
      if (effEnd && itemDateStr > effEnd) return false;

      // 2. Location filter
      if (filterLocationId !== "all") {
        const itemLocId = r.LocationId || r.Location?.Id;
        if (itemLocId !== filterLocationId) return false;
      }

      // 3. Multi-Person Requester filter
      if (selectedRequesters.length > 0) {
        const reqEmail = (r.Requester?.EMail || "").toLowerCase().trim();
        const reqName = (r.Requester?.Title || "").toLowerCase().trim();

        const matches = selectedRequesters.some((u) => {
          const userEmail = (u.email || "").toLowerCase().trim();
          const userName = (u.title || "").toLowerCase().trim();

          // Match by email
          if (
            userEmail &&
            reqEmail &&
            (reqEmail === userEmail ||
              reqEmail.indexOf(userEmail) !== -1 ||
              userEmail.indexOf(reqEmail) !== -1)
          ) {
            return true;
          }
          // Match by name
          if (
            userName &&
            reqName &&
            (reqName === userName ||
              reqName.indexOf(userName) !== -1 ||
              userName.indexOf(reqName) !== -1)
          ) {
            return true;
          }
          return false;
        });

        if (!matches) {
          return false;
        }
      }

      return true;
    });
  }, [requests, startDate, endDate, filterLocationId, selectedRequesters]);

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

  // Reset to first page when filters or search change
  React.useEffect(() => {
    setCurrentPage(1);
  }, [startDate, endDate, filterLocationId, selectedRequesters, setCurrentPage]);

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
            {isFilterOpen && (
              <>
                {/* Responsive Backdrop (< lg screens) */}
                <div
                  className="fixed inset-0 bg-slate-900/40 z-40 lg:hidden backdrop-blur-xs"
                  onClick={() => setIsFilterOpen(false)}
                />

                <div
                  className="fixed inset-x-4 top-20 max-w-lg mx-auto lg:mx-0 lg:max-w-none lg:inset-x-auto lg:absolute lg:top-full lg:right-0 mt-0 lg:mt-2 w-auto lg:w-[460px] max-h-[85vh] overflow-y-auto bg-white rounded-2xl border-2 border-slate-200/90 shadow-2xl p-4 sm:p-5 z-50 animate-in fade-in zoom-in-95 duration-150"
                  style={{ filter: "drop-shadow(0 20px 25px -5px rgba(0, 0, 0, 0.2))" }}
                >
                  {/* Popover Header */}
                  <div className="flex items-center justify-between pb-3.5 mb-3.5 border-b border-slate-100">
                    <div className="flex items-center gap-2 text-[#0f172a]">
                      <div className="w-7 h-7 rounded-lg bg-blue-50 text-[#004B87] flex items-center justify-center">
                        <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                          <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
                        </svg>
                      </div>
                      <h3 className="text-[16px] font-bold tracking-tight">Advanced Filter Criteria</h3>
                    </div>
                    <button
                      type="button"
                      onClick={() => setIsFilterOpen(false)}
                      className="w-7 h-7 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 flex items-center justify-center transition cursor-pointer"
                      title="Close filters"
                      aria-label="Close filters"
                    >
                      <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18" />
                        <line x1="6" y1="6" x2="18" y2="18" />
                      </svg>
                    </button>
                  </div>

                  {/* Popover Form Fields */}
                  <div className="space-y-3.5">
                    {/* 1. Date Range Filter & Quick Presets */}
                    <div>
                      <div className="flex items-center justify-between mb-1.5">
                        <label className="block text-[12px] font-bold text-slate-600 uppercase tracking-wider">
                          MEAL DATE RANGE
                        </label>
                        <div className="flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              setStartDate(todayISOString);
                              setEndDate(todayISOString);
                            }}
                            className={`px-2.5 py-0.5 rounded-md text-[12px] font-semibold border transition cursor-pointer ${
                              isTodayActive
                                ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs"
                                : "bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100"
                            }`}
                          >
                            Today
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setStartDate(tomorrowISOString);
                              setEndDate(tomorrowISOString);
                            }}
                            className={`px-2.5 py-0.5 rounded-md text-[12px] font-semibold border transition cursor-pointer ${
                              isTomorrowActive
                                ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs"
                                : "bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100"
                            }`}
                          >
                            Tomorrow
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setStartDate(null);
                              setEndDate(null);
                            }}
                            className={`px-2.5 py-0.5 rounded-md text-[12px] font-semibold border transition cursor-pointer ${
                              isAllDatesActive
                                ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs"
                                : "bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100"
                            }`}
                          >
                            All Dates
                          </button>
                        </div>
                      </div>

                      <div className="flex items-center gap-2">
                        {/* From Date */}
                        <div className="flex items-center bg-white border border-slate-300 rounded-xl overflow-hidden focus-within:ring-2 focus-within:ring-[#004B87] focus-within:border-[#004B87] flex-1 min-h-[40px]">
                          <span className="px-2.5 text-[12.5px] font-bold text-slate-500 bg-slate-100 border-r border-slate-200 py-2 select-none shrink-0">
                            From
                          </span>
                          <input
                            type="date"
                            value={startDate || ""}
                            max={endDate || undefined}
                            onChange={(e) => handleStartDateChange(e.target.value || null)}
                            aria-label="From Date"
                            className="w-full h-[40px] px-2.5 text-[13.5px] font-medium text-slate-800 bg-white focus:outline-none"
                          />
                        </div>

                        <span className="text-slate-400 font-bold text-xs shrink-0 select-none">to</span>

                        {/* To Date */}
                        <div className="flex items-center bg-white border border-slate-300 rounded-xl overflow-hidden focus-within:ring-2 focus-within:ring-[#004B87] focus-within:border-[#004B87] flex-1 min-h-[40px]">
                          <span className="px-2.5 text-[12.5px] font-bold text-slate-500 bg-slate-100 border-r border-slate-200 py-2 select-none shrink-0">
                            To
                          </span>
                          <input
                            type="date"
                            value={endDate || ""}
                            min={startDate || undefined}
                            onChange={(e) => handleEndDateChange(e.target.value || null)}
                            aria-label="To Date"
                            className="w-full h-[40px] px-2.5 text-[13.5px] font-medium text-slate-800 bg-white focus:outline-none"
                          />
                        </div>

                        {/* Clear Range Button */}
                        {isCustomRange && (
                          <button
                            type="button"
                            onClick={() => {
                              setStartDate(null);
                              setEndDate(null);
                            }}
                            title="Clear Date Filter"
                            className="h-[40px] px-2.5 rounded-xl border border-slate-300 bg-slate-100 hover:bg-slate-200 text-slate-600 transition flex items-center justify-center shrink-0 cursor-pointer"
                          >
                            <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                              <line x1="18" y1="6" x2="6" y2="18" />
                              <line x1="6" y1="6" x2="18" y2="18" />
                            </svg>
                          </button>
                        )}
                      </div>
                    </div>

                    {/* 2. Canteen / Mess Location Filter */}
                    <div>
                      <label className="block text-[12px] font-bold text-slate-600 uppercase tracking-wider mb-1.5">
                        CANTEEN LOCATION
                      </label>
                      <select
                        value={filterLocationId}
                        onChange={(e) =>
                          setFilterLocationId(e.target.value === "all" ? "all" : Number(e.target.value))
                        }
                        className="w-full h-10 px-3 rounded-xl border border-slate-300 text-[14px] font-medium text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-[#004B87]"
                      >
                        <option value="all">All Mess Locations</option>
                        {locations.map((loc) => (
                          <option key={loc.Id} value={loc.Id}>
                            {loc.Title}
                          </option>
                        ))}
                      </select>
                    </div>

                    {/* 3. Multi-Person Requester Filter */}
                    <div>
                      <div className="flex items-center justify-between mb-1.5">
                        <label className="block text-[12px] font-bold text-slate-600 uppercase tracking-wider">
                          EMPLOYEE / REQUESTER
                        </label>
                        {selectedRequesters.length > 0 && (
                          <button
                            type="button"
                            onClick={() => setSelectedRequesters([])}
                            className="text-[11.5px] font-bold text-rose-600 hover:text-rose-700 transition cursor-pointer"
                          >
                            Clear ({selectedRequesters.length})
                          </button>
                        )}
                      </div>
                      <TagPeoplePicker
                        context={context}
                        spHttpClient={spHttpClient}
                        siteUrl={siteUrl}
                        selectedUsers={selectedRequesters}
                        onUsersSelected={setSelectedRequesters}
                      />
                    </div>
                  </div>

                  {/* Popover Footer */}
                  <div className="flex items-center justify-between pt-3.5 mt-4 border-t border-slate-100 gap-2">
                    <button
                      type="button"
                      onClick={handleResetFilters}
                      disabled={activeFilterCount === 0}
                      className="px-3.5 py-2 rounded-xl text-[13px] font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition cursor-pointer"
                    >
                      Reset All
                    </button>
                    <button
                      type="button"
                      onClick={() => setIsFilterOpen(false)}
                      className="px-5 py-2 rounded-xl bg-[#004B87] hover:bg-[#003865] text-white text-[13.5px] font-bold shadow-xs hover:shadow transition cursor-pointer"
                    >
                      Apply &amp; Close
                    </button>
                  </div>
                </div>
              </>
            )}
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
            {(startDate || endDate || filterLocationId !== "all" || selectedRequesters.length > 0) && (
              <button
                type="button"
                onClick={() => {
                  setStartDate(null);
                  setEndDate(null);
                  setFilterLocationId("all");
                  setSelectedRequesters([]);
                }}
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
