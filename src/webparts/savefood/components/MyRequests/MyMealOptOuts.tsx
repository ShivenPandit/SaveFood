import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem, MealType } from "../../models/IMealOptOut";
import { MealRequestService } from "../../services/mealRequestService";
import { validateRetractionCutOff } from "../../services/cutOffService";
import { BackButton } from "../Common/BackButton";
import { useDynamicPagination } from "../../hooks/useDynamicPagination";
import { DataTablePagination } from "../Common/DataTablePagination";

export interface IMyMealOptOutsProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  currentUserId: number;
  currentUserEmail: string;
  onNavigateCreate: () => void;
  onBack: () => void;
  onShowToast?: (message: string, type: "success" | "error" | "info") => void;
}

export const MyMealOptOuts: React.FC<IMyMealOptOutsProps> = ({
  context,
  spHttpClient,
  siteUrl,
  currentUserId,
  currentUserEmail,
  onNavigateCreate,
  onBack,
  onShowToast,
}) => {
  const cachedInitial = React.useMemo(() => {
    return MealRequestService.getCachedMyMealOptOuts(currentUserId) || [];
  }, [currentUserId]);

  const [requests, setRequests] = React.useState<IMealOptOutRequestItem[]>(cachedInitial);
  const [isLoading, setIsLoading] = React.useState<boolean>(cachedInitial.length === 0);
  const [cancellingId, setCancellingId] = React.useState<number | null>(null);

  // Filter State
  const [selectedMealFilter, setSelectedMealFilter] = React.useState<"ALL" | MealType>("ALL");
  const [searchLocation, setSearchLocation] = React.useState<string>("");
  const [filterStartDate, setFilterStartDate] = React.useState<string>("");
  const [filterEndDate, setFilterEndDate] = React.useState<string>("");
  const [filterStatus, setFilterStatus] = React.useState<"ALL" | "ACTIVE" | "LOCKED">("ALL");
  const [isFilterOpen, setIsFilterOpen] = React.useState<boolean>(false);
  const filterRef = React.useRef<HTMLDivElement>(null);

  // Close filter popover on outside click or ESC key
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (filterRef.current && !filterRef.current.contains(event.target as Node)) {
        setIsFilterOpen(false);
      }
    };
    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.key === "Escape") {
        setIsFilterOpen(false);
      }
    };
    if (isFilterOpen) {
      document.addEventListener("mousedown", handleClickOutside);
      document.addEventListener("keydown", handleKeyDown);
    }
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [isFilterOpen]);

  // Unique mess locations from requests for dropdown selection
  const uniqueLocations = React.useMemo(() => {
    const locs = new Set<string>();
    requests.forEach((r) => {
      if (r.Location?.Title) {
        locs.add(r.Location.Title);
      }
    });
    return Array.from(locs).sort();
  }, [requests]);

  // Active filter count for badge
  const activeFilterCount = React.useMemo(() => {
    let count = 0;
    if (selectedMealFilter !== "ALL") count++;
    if (searchLocation.trim() !== "") count++;
    if (filterStartDate !== "") count++;
    if (filterEndDate !== "") count++;
    if (filterStatus !== "ALL") count++;
    return count;
  }, [selectedMealFilter, searchLocation, filterStartDate, filterEndDate, filterStatus]);

  const handleResetAll = (): void => {
    setSelectedMealFilter("ALL");
    setSearchLocation("");
    setFilterStartDate("");
    setFilterEndDate("");
    setFilterStatus("ALL");
  };

  // Real-time 30-second tick to continuously evaluate meal cut-off deadlines
  const [, setTick] = React.useState<number>(0);
  React.useEffect(() => {
    const timer = setInterval(() => {
      setTick((t) => t + 1);
    }, 30000);
    return () => clearInterval(timer);
  }, []);

  const loadMyRequests = React.useCallback(async (silent: boolean = false) => {
    if (!silent) {
      setIsLoading(true);
    }
    const items = await MealRequestService.getMyMealOptOuts(
      spHttpClient,
      siteUrl,
      currentUserId,
      currentUserEmail
    );
    setRequests(items);
    setIsLoading(false);
  }, [spHttpClient, siteUrl, currentUserId, currentUserEmail]);

  React.useEffect(() => {
    const hasCache = cachedInitial.length > 0;
    loadMyRequests(hasCache).catch((err: Error) => {
      console.error(err);
      setIsLoading(false);
    });
  }, [loadMyRequests, cachedInitial]);

  // Handle Cancellation of an opt-out before deadline
  const handleCancel = async (item: IMealOptOutRequestItem): Promise<void> => {
    // Immediate pre-check against exact IST cut-off schedule (10 mins buffer)
    const preCheck = validateRetractionCutOff(item.MealDate, item.MealType);
    if (!preCheck.isValid) {
      if (onShowToast) {
        onShowToast(preCheck.errorMessage || "Cannot retract opt-out. Cut-off time has passed.", "error");
      }
      return;
    }

    const mealDateFormatted = item.MealDate ? item.MealDate.split("T")[0] : "";
    const confirmCancel = window.confirm(
      `Are you sure you want to retract your opt-out for ${item.MealType} on ${mealDateFormatted}?`
    );
    if (!confirmCancel) return;

    setCancellingId(item.Id);

    const result = await MealRequestService.cancelOptOut(spHttpClient, siteUrl, item, context);
    setCancellingId(null);

    if (result.success) {
      if (onShowToast) {
        onShowToast(`Opt-out for ${item.MealType} on ${mealDateFormatted} successfully retracted.`, "success");
      }
      loadMyRequests().catch((err: Error) => {
        console.error(err);
      });
    } else {
      if (onShowToast) {
        onShowToast(result.error || "Failed to retract opt-out.", "error");
      }
    }
  };

  // Filtered List
  const filteredRequests = React.useMemo(() => {
    return requests.filter((req) => {
      // 1. Meal Type Filter
      const matchMeal = selectedMealFilter === "ALL" || req.MealType === selectedMealFilter;

      // 2. Mess Location Filter
      const matchLoc =
        !searchLocation.trim() ||
        (req.Location?.Title || "").toLowerCase().indexOf(searchLocation.toLowerCase().trim()) !== -1;

      // 3. Date Range Filter
      let matchDate = true;
      if (req.MealDate) {
        const itemDateStr = req.MealDate.split("T")[0];
        if (filterStartDate && itemDateStr < filterStartDate) {
          matchDate = false;
        }
        if (filterEndDate && itemDateStr > filterEndDate) {
          matchDate = false;
        }
      }

      // 4. Cut-Off Status Filter
      let matchStatus = true;
      if (filterStatus !== "ALL") {
        const cutOffCheck = validateRetractionCutOff(req.MealDate, req.MealType);
        const isActive = cutOffCheck.isValid;
        if (filterStatus === "ACTIVE" && !isActive) {
          matchStatus = false;
        } else if (filterStatus === "LOCKED" && isActive) {
          matchStatus = false;
        }
      }

      return matchMeal && matchLoc && matchDate && matchStatus;
    });
  }, [requests, selectedMealFilter, searchLocation, filterStartDate, filterEndDate, filterStatus]);

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
    viewType: "my-requests",
  });

  // Reset to first page when search filter changes
  React.useEffect(() => {
    setCurrentPage(1);
  }, [selectedMealFilter, searchLocation, filterStartDate, filterEndDate, filterStatus, setCurrentPage]);

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header */}
      <div className="mb-2 sm:mb-2.5 flex flex-col md:flex-row md:items-center md:justify-between gap-3 w-full">
        <div>
          <BackButton onClick={onBack} className="mb-1.5" />
          <h2 className="text-[26px] sm:text-[28px] font-bold text-[#0f172a] tracking-tight">
            My Meal Opt-Out Requests
          </h2>
          <p className="text-[15px] text-[#475569] leading-relaxed mt-1">
            View and manage your submitted meal advance cancellations across all Vedanta canteens.
          </p>
        </div>

        {/* Action Controls Bar */}
        <div className="flex items-center gap-2.5 flex-wrap self-start md:self-auto">
          {/* Filters Toggle Button & Dropdown Popover */}
          <div ref={filterRef} className="relative">
            <button
              type="button"
              onClick={() => setIsFilterOpen((prev) => !prev)}
              aria-expanded={isFilterOpen}
              className={`h-[42px] px-4 rounded-xl border inline-flex items-center gap-2 text-[14.5px] font-semibold transition cursor-pointer select-none whitespace-nowrap ${
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
                  className="fixed inset-x-4 top-20 max-w-md mx-auto lg:mx-0 lg:max-w-none lg:inset-x-auto lg:absolute lg:top-full lg:right-0 mt-0 lg:mt-2 w-auto lg:w-[420px] max-h-[85vh] overflow-y-auto bg-white rounded-2xl border-2 border-slate-200/90 shadow-2xl p-4 sm:p-5 z-50 animate-in fade-in zoom-in-95 duration-150"
                  style={{ filter: "drop-shadow(0 20px 25px -5px rgba(0, 0, 0, 0.2))" }}
                >
                  {/* Popover Header */}
                  <div className="flex items-center justify-between pb-3.5 mb-4 border-b border-slate-100">
                    <div className="flex items-center gap-2 text-[#0f172a]">
                      <div className="w-7 h-7 rounded-lg bg-blue-50 text-[#004B87] flex items-center justify-center">
                        <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                          <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
                        </svg>
                      </div>
                      <h3 className="text-[16px] font-bold tracking-tight">Advanced Search Criteria</h3>
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
                  <div className="space-y-3 sm:space-y-3.5">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      {/* MEAL TYPE */}
                      <div>
                        <label className="block text-[12px] font-bold text-slate-600 uppercase tracking-wider mb-1.5">
                          MEAL TYPE
                        </label>
                        <select
                          value={selectedMealFilter}
                          onChange={(e) => setSelectedMealFilter(e.target.value as "ALL" | MealType)}
                          className="w-full h-10 px-3 rounded-xl border border-slate-300 text-[14px] font-medium text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-[#004B87]"
                        >
                          <option value="ALL">All Meals</option>
                          <option value="Breakfast">Breakfast</option>
                          <option value="Lunch">Lunch</option>
                          <option value="Dinner">Dinner</option>
                        </select>
                      </div>

                      {/* CANTEEN LOCATION */}
                      <div>
                        <label className="block text-[12px] font-bold text-slate-600 uppercase tracking-wider mb-1.5">
                          CANTEEN LOCATION
                        </label>
                        <select
                          value={searchLocation}
                          onChange={(e) => setSearchLocation(e.target.value)}
                          className="w-full h-10 px-3 rounded-xl border border-slate-300 text-[14px] font-medium text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-[#004B87]"
                        >
                          <option value="">All Mess Locations</option>
                          {uniqueLocations.map((loc) => (
                            <option key={loc} value={loc}>
                              {loc}
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>

                    {/* START DATE (FROM / TO) */}
                    <div>
                      <label className="block text-[12px] font-bold text-slate-600 uppercase tracking-wider mb-1.5">
                        START DATE (FROM / TO)
                      </label>
                      <div className="grid grid-cols-2 gap-2 sm:gap-2.5">
                        <div>
                          <input
                            type="date"
                            value={filterStartDate}
                            max={filterEndDate || undefined}
                            onChange={(e) => {
                              const val = e.target.value;
                              setFilterStartDate(val);
                              if (val && filterEndDate && val > filterEndDate) {
                                setFilterEndDate(val);
                              }
                            }}
                            className="w-full h-10 px-2 sm:px-3 rounded-xl border border-slate-300 text-[13px] sm:text-[13.5px] font-medium text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-[#004B87]"
                          />
                        </div>
                        <div>
                          <input
                            type="date"
                            value={filterEndDate}
                            min={filterStartDate || undefined}
                            onChange={(e) => {
                              const val = e.target.value;
                              if (val && filterStartDate && val < filterStartDate) {
                                setFilterEndDate(filterStartDate);
                              } else {
                                setFilterEndDate(val);
                              }
                            }}
                            className="w-full h-10 px-2 sm:px-3 rounded-xl border border-slate-300 text-[13px] sm:text-[13.5px] font-medium text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-[#004B87]"
                          />
                        </div>
                      </div>
                    </div>

                    {/* CUT-OFF STATUS */}
                    <div>
                      <label className="block text-[12px] font-bold text-slate-600 uppercase tracking-wider mb-1.5">
                        CUT-OFF STATUS
                      </label>
                      <select
                        value={filterStatus}
                        onChange={(e) => setFilterStatus(e.target.value as "ALL" | "ACTIVE" | "LOCKED")}
                        className="w-full h-10 px-3 rounded-xl border border-slate-300 text-[14px] font-medium text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-[#004B87]"
                      >
                        <option value="ALL">All Statuses</option>
                        <option value="ACTIVE">Active (Modifiable)</option>
                        <option value="LOCKED">Locked (Passed)</option>
                      </select>
                    </div>
                  </div>

                  {/* Popover Footer */}
                  <div className="pt-3.5 mt-4 border-t border-slate-100 flex items-center justify-between">
                    <button
                      type="button"
                      onClick={handleResetAll}
                      className="inline-flex items-center gap-1.5 text-rose-600 hover:text-rose-700 text-[14px] font-semibold transition cursor-pointer select-none"
                    >
                      <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                        <path d="M3 3v5h5" />
                      </svg>
                      <span>Reset All</span>
                    </button>

                    <button
                      type="button"
                      onClick={() => setIsFilterOpen(false)}
                      className="h-10 px-5 rounded-xl bg-[#004B87] hover:bg-[#003865] text-white text-[14px] font-semibold shadow-xs transition cursor-pointer whitespace-nowrap"
                    >
                      Apply &amp; Close
                    </button>
                  </div>
                </div>
              </>
            )}
          </div>

          {/* New Meal Opt-Out CTA Button */}
          <button
            type="button"
            onClick={onNavigateCreate}
            className="h-[42px] inline-flex items-center justify-center gap-2 bg-[#0066CC] hover:bg-[#0052A3] text-white text-[15px] font-semibold px-5 py-2 rounded-full shadow-2xs hover:shadow-xs transition cursor-pointer self-start sm:self-auto whitespace-nowrap"
          >
            <svg viewBox="0 0 24 24" className="w-4 h-4 shrink-0 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <line x1="12" y1="5" x2="12" y2="19" />
              <line x1="5" y1="12" x2="19" y2="12" />
            </svg>
            <span>New Meal Opt-Out</span>
          </button>
        </div>
      </div>

      {/* Main Table Card */}
      <div
        ref={containerRef}
        style={{ height: "auto", flex: "0 0 auto" }}
        className="w-full bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm overflow-hidden"
      >
        {isLoading ? (
          <div style={{ height: "auto", flex: "0 0 auto", padding: "32px" }} className="text-center">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-4 border-[#004B87] border-t-transparent" />
            <p className="mt-3 text-[14.5px] font-semibold text-slate-500">
              Loading your requests...
            </p>
          </div>
        ) : filteredRequests.length === 0 ? (
          <div style={{ height: "auto", flex: "0 0 auto", padding: "32px" }} className="text-center">
            <div className="w-14 h-14 mx-auto rounded-2xl bg-blue-50 text-blue-600 border border-blue-200 flex items-center justify-center mb-3 shadow-2xs">
              <svg viewBox="0 0 24 24" className="w-7 h-7 fill-none stroke-current" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                <line x1="16" y1="2" x2="16" y2="6" />
                <line x1="8" y1="2" x2="8" y2="6" />
                <line x1="3" y1="10" x2="21" y2="10" />
              </svg>
            </div>
            <h3 className="text-lg sm:text-xl font-bold text-slate-800">
              {requests.length === 0 ? "No Meal Opt-Out Requests Found" : "No Matching Requests Found"}
            </h3>
            <p className="text-[14.5px] text-slate-500 max-w-md mx-auto mt-1 mb-5 leading-relaxed">
              {requests.length === 0
                ? "You haven't submitted any advance meal cancellations yet."
                : "No requests match the selected filters. Try changing or clearing your search filter."}
            </p>
            {requests.length === 0 ? (
              <button
                type="button"
                onClick={onNavigateCreate}
                className="inline-flex items-center gap-2 bg-[#004B87] hover:bg-[#003865] text-white text-[15px] font-semibold px-6 py-2.5 rounded-full transition shadow-2xs cursor-pointer"
              >
                <span>Create Your First Opt-Out</span>
                <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="5" y1="12" x2="19" y2="12" />
                  <polyline points="12 5 19 12 12 19" />
                </svg>
              </button>
            ) : (
              activeFilterCount > 0 && (
                <button
                  type="button"
                  onClick={handleResetAll}
                  className="inline-flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-[14px] font-semibold px-5 py-2.5 rounded-xl transition cursor-pointer"
                >
                  <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                    <path d="M3 3v5h5" />
                  </svg>
                  <span>Reset All Filters</span>
                </button>
              )
            )}
          </div>
        ) : (
          <>
            {/* Desktop Table View (>= 640px) */}
            <div className="hidden sm:block w-full overflow-x-auto">
              <table className="w-full text-left text-[14.5px]">
                <thead className="bg-slate-50 border-b border-slate-200 text-[#475569] font-bold uppercase tracking-wider text-[12px] sm:text-[12.5px]">
                  <tr>
                    <th className="py-2.5 px-4 sm:px-5">Meal Date</th>
                    <th className="py-2.5 px-4">Meal Type</th>
                    <th className="py-2.5 px-4">Canteen Location</th>
                    <th className="py-2.5 px-4">Submitted On</th>
                    <th className="py-2.5 px-4 text-center">Cut-Off Status</th>
                    <th className="py-2.5 px-4 sm:px-5 text-right">Action</th>
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

                    const cutOffCheck = validateRetractionCutOff(item.MealDate, item.MealType);
                    const canCancel = cutOffCheck.isValid;

                    return (
                      <tr key={item.Id} className="hover:bg-slate-50/80 transition-colors">
                        <td className="py-2.5 px-4 sm:px-5 font-bold text-slate-900 whitespace-nowrap">
                          {formattedMealDate}
                        </td>
                        <td className="py-2.5 px-4 whitespace-nowrap">
                          <span
                            className={`inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[12.5px] font-semibold ${
                              item.MealType === "Breakfast"
                                ? "bg-amber-50 text-amber-800 border border-amber-200"
                                : item.MealType === "Lunch"
                                ? "bg-blue-50 text-blue-800 border border-blue-200"
                                : "bg-purple-50 text-purple-800 border border-purple-200"
                            }`}
                          >
                            <span>{item.MealType === "Breakfast" ? "🍳" : item.MealType === "Lunch" ? "🥗" : "🍲"}</span>
                            <span>{item.MealType}</span>
                          </span>
                        </td>
                        <td className="py-2.5 px-4 font-semibold text-slate-800 whitespace-nowrap">
                          {item.Location?.Title || "—"}
                        </td>
                        <td className="py-2.5 px-4 text-slate-500 whitespace-nowrap text-[13px]">
                          {formattedCreated}
                        </td>
                        <td className="py-2.5 px-4 text-center whitespace-nowrap">
                          {canCancel ? (
                            <span className="inline-flex items-center gap-1.5 text-[12.5px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2.5 py-0.5 rounded-full">
                              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                              Active (Modifiable)
                            </span>
                          ) : (
                            <span
                              title={cutOffCheck.errorMessage || "Past cut-off window"}
                              className="inline-flex items-center gap-1.5 text-[12.5px] font-medium text-slate-500 bg-slate-100 border border-slate-200/80 px-2.5 py-0.5 rounded-full"
                            >
                              <svg viewBox="0 0 24 24" className="w-3.5 h-3.5 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                                <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                                <path d="M7 11V7a5 5 0 0 1 10 0v4" />
                              </svg>
                              Locked (Passed)
                            </span>
                          )}
                        </td>
                        <td className="py-2.5 px-4 sm:px-5 text-right whitespace-nowrap">
                          {canCancel ? (
                            <button
                              type="button"
                              onClick={() => handleCancel(item)}
                              disabled={cancellingId === item.Id}
                              title="Retract Meal Opt-Out"
                              aria-label="Retract Meal Opt-Out"
                              className="h-8 w-8 inline-flex items-center justify-center rounded-xl border border-rose-200 bg-rose-50/70 text-rose-600 hover:text-white hover:bg-rose-600 hover:border-rose-600 shadow-2xs transition-all cursor-pointer disabled:opacity-50"
                            >
                              {cancellingId === item.Id ? (
                                <div className="w-4 h-4 border-2 border-rose-600 border-t-transparent rounded-full animate-spin" />
                              ) : (
                                <svg
                                  viewBox="0 0 24 24"
                                  className="w-4 h-4 shrink-0 fill-none stroke-current"
                                  strokeWidth="2.2"
                                  strokeLinecap="round"
                                  strokeLinejoin="round"
                                >
                                  <polyline points="1 4 1 10 7 10" />
                                  <path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10" />
                                </svg>
                              )}
                            </button>
                          ) : (
                            <div
                              title={cutOffCheck.errorMessage || "Past Cut-Off Window (Retraction Locked)"}
                              className="h-8 w-8 inline-flex items-center justify-center rounded-xl border border-slate-200/60 bg-slate-100/70 text-slate-400 cursor-not-allowed"
                            >
                              <svg
                                viewBox="0 0 24 24"
                                className="w-4 h-4 fill-none stroke-current"
                                strokeWidth="2"
                                strokeLinecap="round"
                                strokeLinejoin="round"
                              >
                                <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                                <path d="M7 11V7a5 5 0 0 1 10 0v4" />
                              </svg>
                            </div>
                          )}
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

                const cutOffCheck = validateRetractionCutOff(item.MealDate, item.MealType);
                const canCancel = cutOffCheck.isValid;

                return (
                  <div
                    key={item.Id}
                    data-role="list-item"
                    className="p-4 hover:bg-slate-50/70 transition-colors flex flex-col gap-3"
                  >
                    <div className="flex items-center justify-between">
                      <span className="font-bold text-slate-900 text-[15px]">
                        {formattedMealDate}
                      </span>
                      <span
                        className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[13px] font-semibold ${
                          item.MealType === "Breakfast"
                            ? "bg-amber-50 text-amber-800 border border-amber-200"
                            : item.MealType === "Lunch"
                            ? "bg-blue-50 text-blue-800 border border-blue-200"
                            : "bg-purple-50 text-purple-800 border border-purple-200"
                        }`}
                      >
                        <span>{item.MealType === "Breakfast" ? "🍳" : item.MealType === "Lunch" ? "🥗" : "🍲"}</span>
                        <span>{item.MealType}</span>
                      </span>
                    </div>

                    <div className="flex items-center justify-between text-[14px] text-slate-600">
                      <div className="flex items-center gap-1.5 font-medium">
                        <span className="text-slate-400">📍</span>
                        <span className="font-semibold text-slate-800">{item.Location?.Title || "—"}</span>
                      </div>
                      <span className="text-[13px] text-slate-400">{formattedCreated}</span>
                    </div>

                    <div className="flex items-center justify-between pt-2 border-t border-slate-100">
                      <div>
                        {canCancel ? (
                          <span className="inline-flex items-center gap-1.5 text-[13px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 px-3 py-1 rounded-full">
                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                            Active
                          </span>
                        ) : (
                          <span
                            title={cutOffCheck.errorMessage || "Past cut-off window"}
                            className="inline-flex items-center gap-1.5 text-[13px] font-medium text-slate-500 bg-slate-100 border border-slate-200/80 px-3 py-1 rounded-full"
                          >
                            <svg viewBox="0 0 24 24" className="w-3.5 h-3.5 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                              <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                              <path d="M7 11V7a5 5 0 0 1 10 0v4" />
                            </svg>
                            Locked (Passed)
                          </span>
                        )}
                      </div>

                      {canCancel ? (
                        <button
                          type="button"
                          onClick={() => handleCancel(item)}
                          disabled={cancellingId === item.Id}
                          title="Retract Meal Opt-Out"
                          aria-label="Retract Meal Opt-Out"
                          className="h-9 px-3.5 inline-flex items-center gap-2 text-[14px] font-semibold text-rose-600 bg-rose-50 border border-rose-200 hover:bg-rose-600 hover:text-white rounded-xl transition-colors cursor-pointer disabled:opacity-50"
                        >
                          {cancellingId === item.Id ? (
                            <div className="w-4 h-4 border-2 border-rose-600 border-t-transparent rounded-full animate-spin" />
                          ) : (
                            <svg
                              viewBox="0 0 24 24"
                              className="w-4 h-4 fill-none stroke-current"
                              strokeWidth="2.2"
                              strokeLinecap="round"
                              strokeLinejoin="round"
                            >
                              <polyline points="1 4 1 10 7 10" />
                              <path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10" />
                            </svg>
                          )}
                          <span>Retract</span>
                        </button>
                      ) : (
                        <div
                          title={cutOffCheck.errorMessage || "Past Cut-Off Window (Retraction Locked)"}
                          className="h-9 px-3.5 inline-flex items-center gap-2 text-[14px] font-medium text-slate-400 bg-slate-100 border border-slate-200/60 rounded-xl cursor-not-allowed"
                        >
                          <svg
                            viewBox="0 0 24 24"
                            className="w-4 h-4 fill-none stroke-current"
                            strokeWidth="2"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                            <path d="M7 11V7a5 5 0 0 1 10 0v4" />
                          </svg>
                          <span>Locked</span>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Dynamic Pagination Bar */}
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
      </div>
    </div>
  );
};
