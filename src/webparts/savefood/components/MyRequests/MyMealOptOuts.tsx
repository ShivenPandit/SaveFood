import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem, MealType } from "../../models/IMealOptOut";
import { MealRequestService } from "../../services/mealRequestService";
import { validateMealCutOff } from "../../services/cutOffService";
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
    // Immediate pre-check against exact IST cut-off schedule
    const preCheck = validateMealCutOff(item.MealDate, item.MealType);
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
      const matchMeal = selectedMealFilter === "ALL" || req.MealType === selectedMealFilter;
      const matchLoc =
        !searchLocation.trim() ||
        (req.Location?.Title || "").toLowerCase().indexOf(searchLocation.toLowerCase().trim()) !== -1;
      return matchMeal && matchLoc;
    });
  }, [requests, selectedMealFilter, searchLocation]);

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
  }, [selectedMealFilter, searchLocation, setCurrentPage]);

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header */}
      <div className="mb-2.5 sm:mb-3 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2.5 w-full">
        <div>
          <BackButton onClick={onBack} />
          <h2 className="text-lg sm:text-xl font-black text-slate-900 tracking-tight">
            My Meal Opt-Out Requests
          </h2>
          <p className="text-xs text-slate-500">
            View and manage your submitted meal advance cancellations across all Vedanta canteens.
          </p>
        </div>

        <button
          type="button"
          onClick={onNavigateCreate}
          className="inline-flex items-center gap-1.5 bg-[#0066CC] hover:bg-[#0052A3] text-white text-xs sm:text-sm font-bold px-4 py-2 rounded-full shadow-2xs hover:shadow-xs transition cursor-pointer self-start sm:self-auto"
        >
          <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <line x1="12" y1="5" x2="12" y2="19" />
            <line x1="5" y1="12" x2="19" y2="12" />
          </svg>
          <span>New Meal Opt-Out</span>
        </button>
      </div>

      {/* Filter and Search Bar Card */}
      <div
        style={{ height: "auto", flex: "0 0 auto" }}
        className="w-full bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm p-2.5 sm:px-4 sm:py-2.5 mb-2.5 sm:mb-3 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2.5"
      >
        {/* Meal Type Filter Pills */}
        <div className="flex items-center gap-1.5 overflow-x-auto pb-0.5 sm:pb-0">
          {(["ALL", "Breakfast", "Lunch", "Dinner"] as Array<"ALL" | MealType>).map((type) => {
            const isSelected = selectedMealFilter === type;
            return (
              <button
                key={type}
                type="button"
                onClick={() => setSelectedMealFilter(type)}
                className={`px-3 py-1 rounded-lg text-xs font-bold transition cursor-pointer whitespace-nowrap ${
                  isSelected
                    ? "bg-[#004B87] text-white shadow-2xs"
                    : "bg-slate-100 hover:bg-slate-200 text-slate-600"
                }`}
              >
                {type === "ALL" ? "All Meals" : type}
              </button>
            );
          })}
        </div>

        {/* Location Search Input */}
        <div className="relative min-w-[220px] sm:w-64">
          <input
            type="text"
            placeholder="Search by mess location..."
            value={searchLocation}
            onChange={(e) => setSearchLocation(e.target.value)}
            className="w-full h-8 pl-8 pr-3 rounded-lg border border-slate-300 text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#004B87]"
          />
          <svg
            viewBox="0 0 24 24"
            className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5 fill-none stroke-current"
            strokeWidth="2.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <circle cx="11" cy="11" r="8" />
            <line x1="21" y1="21" x2="16.65" y2="16.65" />
          </svg>
        </div>
      </div>

      {/* Main Table Card */}
      <div
        ref={containerRef}
        style={{ height: "auto", flex: "0 0 auto" }}
        className="w-full bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm overflow-hidden"
      >
        {isLoading ? (
          <div style={{ height: "auto", flex: "0 0 auto", padding: "24px" }} className="text-center">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-4 border-[#004B87] border-t-transparent" />
            <p className="mt-3 text-xs sm:text-sm font-semibold text-slate-500">
              Loading your requests...
            </p>
          </div>
        ) : filteredRequests.length === 0 ? (
          <div style={{ height: "auto", flex: "0 0 auto", padding: "24px" }} className="text-center">
            <div className="w-12 h-12 mx-auto rounded-xl bg-blue-50 text-blue-600 border border-blue-200 flex items-center justify-center mb-2.5 shadow-2xs">
              <svg viewBox="0 0 24 24" className="w-6 h-6 fill-none stroke-current" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                <line x1="16" y1="2" x2="16" y2="6" />
                <line x1="8" y1="2" x2="8" y2="6" />
                <line x1="3" y1="10" x2="21" y2="10" />
              </svg>
            </div>
            <h3 className="text-base sm:text-lg font-bold text-slate-800">
              {requests.length === 0 ? "No Meal Opt-Out Requests Found" : "No Matching Requests Found"}
            </h3>
            <p className="text-xs sm:text-sm text-slate-500 max-w-md mx-auto mt-1 mb-4">
              {requests.length === 0
                ? "You haven't submitted any advance meal cancellations yet."
                : "No requests match the selected filters. Try changing or clearing your search filter."}
            </p>
            {requests.length === 0 && (
              <button
                type="button"
                onClick={onNavigateCreate}
                className="inline-flex items-center gap-2 bg-[#004B87] hover:bg-[#003865] text-white text-xs sm:text-sm font-bold px-5 py-2.5 rounded-full transition shadow-2xs cursor-pointer"
              >
                <span>Create Your First Opt-Out</span>
                <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="5" y1="12" x2="19" y2="12" />
                  <polyline points="12 5 19 12 12 19" />
                </svg>
              </button>
            )}
          </div>
        ) : (
          <>
            {/* Desktop Table View (>= 640px) */}
            <div className="hidden sm:block w-full overflow-x-auto">
              <table className="w-full text-left text-xs sm:text-sm">
                <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider text-[10.5px] sm:text-xs">
                  <tr>
                    <th className="py-2 px-3 sm:px-5">Meal Date</th>
                    <th className="py-2 px-3">Meal Type</th>
                    <th className="py-2 px-3">Canteen Location</th>
                    <th className="py-2 px-3">Submitted On</th>
                    <th className="py-2 px-3 text-center">Cut-Off Status</th>
                    <th className="py-2 px-3 sm:px-5 text-right">Action</th>
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

                    const cutOffCheck = validateMealCutOff(item.MealDate, item.MealType);
                    const canCancel = cutOffCheck.isValid;

                    return (
                      <tr key={item.Id} className="hover:bg-slate-50/80 transition-colors">
                        <td className="py-2 px-3 sm:px-5 font-bold text-slate-900 whitespace-nowrap">
                          {formattedMealDate}
                        </td>
                        <td className="py-2 px-3 whitespace-nowrap">
                          <span
                            className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10.5px] font-bold ${
                              item.MealType === "Breakfast"
                                ? "bg-amber-50 text-amber-700 border border-amber-200"
                                : item.MealType === "Lunch"
                                ? "bg-blue-50 text-blue-700 border border-blue-200"
                                : "bg-purple-50 text-purple-700 border border-purple-200"
                            }`}
                          >
                            <span>{item.MealType === "Breakfast" ? "🍳" : item.MealType === "Lunch" ? "🥗" : "🍲"}</span>
                            <span>{item.MealType}</span>
                          </span>
                        </td>
                        <td className="py-2 px-3 font-semibold text-slate-800 whitespace-nowrap">
                          {item.Location?.Title || "—"}
                        </td>
                        <td className="py-2 px-3 text-slate-500 whitespace-nowrap text-[11px]">
                          {formattedCreated}
                        </td>
                        <td className="py-2 px-3 text-center whitespace-nowrap">
                          {canCancel ? (
                            <span className="inline-flex items-center gap-1 text-[10.5px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">
                              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                              Active (Modifiable)
                            </span>
                          ) : (
                            <span
                              title={cutOffCheck.errorMessage || "Past cut-off window"}
                              className="inline-flex items-center gap-1 text-[10.5px] font-semibold text-slate-500 bg-slate-100 border border-slate-200/80 px-2 py-0.5 rounded-full"
                            >
                              <svg viewBox="0 0 24 24" className="w-3 h-3 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                                <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                                <path d="M7 11V7a5 5 0 0 1 10 0v4" />
                              </svg>
                              Locked (Passed)
                            </span>
                          )}
                        </td>
                        <td className="py-2 px-3 sm:px-5 text-right whitespace-nowrap">
                          {canCancel ? (
                            <button
                              type="button"
                              onClick={() => handleCancel(item)}
                              disabled={cancellingId === item.Id}
                              title="Retract Meal Opt-Out"
                              aria-label="Retract Meal Opt-Out"
                              className="h-8 w-8 inline-flex items-center justify-center rounded-lg border border-rose-200 bg-rose-50/70 text-rose-600 hover:text-white hover:bg-rose-600 hover:border-rose-600 shadow-2xs transition-all cursor-pointer disabled:opacity-50"
                            >
                              {cancellingId === item.Id ? (
                                <div className="w-3.5 h-3.5 border-2 border-rose-600 border-t-transparent rounded-full animate-spin" />
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
                            </button>
                          ) : (
                            <div
                              title={cutOffCheck.errorMessage || "Past Cut-Off Window (Retraction Locked)"}
                              className="h-8 w-8 inline-flex items-center justify-center rounded-lg border border-slate-200/60 bg-slate-100/70 text-slate-400 cursor-not-allowed"
                            >
                              <svg
                                viewBox="0 0 24 24"
                                className="w-3.5 h-3.5 fill-none stroke-current"
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

                const cutOffCheck = validateMealCutOff(item.MealDate, item.MealType);
                const canCancel = cutOffCheck.isValid;

                return (
                  <div
                    key={item.Id}
                    data-role="list-item"
                    className="p-4 hover:bg-slate-50/70 transition-colors flex flex-col gap-2.5"
                  >
                    <div className="flex items-center justify-between">
                      <span className="font-extrabold text-slate-900 text-sm">
                        {formattedMealDate}
                      </span>
                      <span
                        className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold ${
                          item.MealType === "Breakfast"
                            ? "bg-amber-50 text-amber-700 border border-amber-200"
                            : item.MealType === "Lunch"
                            ? "bg-blue-50 text-blue-700 border border-blue-200"
                            : "bg-purple-50 text-purple-700 border border-purple-200"
                        }`}
                      >
                        <span>{item.MealType === "Breakfast" ? "🍳" : item.MealType === "Lunch" ? "🥗" : "🍲"}</span>
                        <span>{item.MealType}</span>
                      </span>
                    </div>

                    <div className="flex items-center justify-between text-xs text-slate-600">
                      <div className="flex items-center gap-1 font-medium">
                        <span className="text-slate-400">📍</span>
                        <span className="font-semibold text-slate-800">{item.Location?.Title || "—"}</span>
                      </div>
                      <span className="text-[11px] text-slate-400">{formattedCreated}</span>
                    </div>

                    <div className="flex items-center justify-between pt-1 border-t border-slate-100">
                      <div>
                        {canCancel ? (
                          <span className="inline-flex items-center gap-1 text-[10.5px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">
                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                            Active
                          </span>
                        ) : (
                          <span
                            title={cutOffCheck.errorMessage || "Past cut-off window"}
                            className="inline-flex items-center gap-1 text-[10.5px] font-semibold text-slate-500 bg-slate-100 border border-slate-200/80 px-2 py-0.5 rounded-full"
                          >
                            <svg viewBox="0 0 24 24" className="w-3 h-3 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
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
                          className="h-8 px-2.5 inline-flex items-center gap-1.5 text-xs font-bold text-rose-600 bg-rose-50 border border-rose-200 hover:bg-rose-600 hover:text-white rounded-lg transition-colors cursor-pointer disabled:opacity-50"
                        >
                          {cancellingId === item.Id ? (
                            <div className="w-3.5 h-3.5 border-2 border-rose-600 border-t-transparent rounded-full animate-spin" />
                          ) : (
                            <svg
                              viewBox="0 0 24 24"
                              className="w-3.5 h-3.5 fill-none stroke-current"
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
                          className="h-8 px-2.5 inline-flex items-center gap-1.5 text-xs font-semibold text-slate-400 bg-slate-100 border border-slate-200/60 rounded-lg cursor-not-allowed"
                        >
                          <svg
                            viewBox="0 0 24 24"
                            className="w-3.5 h-3.5 fill-none stroke-current"
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
