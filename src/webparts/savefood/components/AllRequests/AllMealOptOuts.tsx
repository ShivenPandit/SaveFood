import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem } from "../../models/IMealOptOut";
import { ILocationMasterItem } from "../../models/ILocation";
import { MealRequestService } from "../../services/mealRequestService";
import { getCurrentISTDate } from "../../services/cutOffService";
import { BackButton } from "../Common/BackButton";
import { useDynamicPagination } from "../../hooks/useDynamicPagination";
import { DataTablePagination } from "../Common/DataTablePagination";

export interface IAllMealOptOutsProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  onBack: () => void;
  hideBackButton?: boolean;
}

export const AllMealOptOuts: React.FC<IAllMealOptOutsProps> = ({
  spHttpClient,
  siteUrl,
  onBack,
  hideBackButton = false,
}) => {
  const [requests, setRequests] = React.useState<IMealOptOutRequestItem[]>([]);
  const [locations, setLocations] = React.useState<ILocationMasterItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  // Filters: Default to All Dates (null) so all meal records appear immediately
  const todayIST = getCurrentISTDate();
  const todayISOString = todayIST.toISOString().split("T")[0];

  const [filterDate, setFilterDate] = React.useState<string | null>(null);
  const [filterLocationId, setFilterLocationId] = React.useState<number | "all">("all");
  const [searchQuery, setSearchQuery] = React.useState<string>("");

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
    async () => {
      setIsLoading(true);

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
    fetchRequests().catch((err: Error) => {
      console.error(err);
    });
  }, [fetchRequests]);

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

  // Instant in-memory filtering: Date, Location, and Search query
  const filteredRequests = React.useMemo(() => {
    const query = searchQuery.trim().toLowerCase();

    return requests.filter((r) => {
      // 1. Date filter (exact YYYY-MM-DD string match, immune to UTC timezone shift)
      if (filterDate) {
        const itemDateStr = r.MealDate ? r.MealDate.split("T")[0] : "";
        if (itemDateStr !== filterDate) return false;
      }

      // 2. Location filter
      if (filterLocationId !== "all") {
        const itemLocId = r.LocationId || r.Location?.Id;
        if (itemLocId !== filterLocationId) return false;
      }

      // 3. Search text
      if (query) {
        const requesterName = (r.Requester?.Title || "").toLowerCase();
        const requesterEmail = (r.Requester?.EMail || "").toLowerCase();
        const summary = (r.Title || "").toLowerCase();
        const loc = (r.Location?.Title || "").toLowerCase();
        if (
          requesterName.indexOf(query) === -1 &&
          requesterEmail.indexOf(query) === -1 &&
          summary.indexOf(query) === -1 &&
          loc.indexOf(query) === -1
        ) {
          return false;
        }
      }

      return true;
    });
  }, [requests, filterDate, filterLocationId, searchQuery]);

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
    bufferBottom: 32,
    defaultRowHeight: 52,
  });

  // Reset to first page when filters or search change
  React.useEffect(() => {
    setCurrentPage(1);
  }, [filterDate, filterLocationId, searchQuery, setCurrentPage]);

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header */}
      <div className="mb-3.5 flex flex-col md:flex-row md:items-center md:justify-between gap-3">
        <div>
          {!hideBackButton && <BackButton onClick={onBack} />}
          <div className="flex items-center gap-2.5">
            <h2 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
              All Meal Opt-Out Requests
            </h2>
          </div>
          <p className="text-xs sm:text-sm text-slate-500 mt-0.5">
            Consolidated kitchen catering headcount roster and location-wise opt-out records.
          </p>
        </div>

        {/* Count Pill */}
        <div className="flex items-center gap-2 self-start md:self-auto bg-slate-100 px-3.5 py-1.5 rounded-full text-xs font-bold text-slate-700">
          <span>Total Records Shown:</span>
          <span className="text-[#004B87] text-sm">{filteredRequests.length}</span>
        </div>
      </div>

      {/* Filter Bar Card */}
      <div
        style={{ height: "auto", flex: "0 0 auto" }}
        className="bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm p-4 sm:p-5 mb-3.5"
      >
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4 items-end">
          {/* 1. Date Filter */}
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1">
              Filter by Meal Date
            </label>
            <div className="flex items-center gap-2">
              <input
                type="date"
                value={filterDate || ""}
                onChange={(e) => setFilterDate(e.target.value || null)}
                className="w-full h-10 px-3 rounded-xl border border-slate-300 text-xs sm:text-sm font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#004B87]"
              />
              <button
                type="button"
                onClick={() => setFilterDate(todayISOString)}
                title="Filter for Today"
                className={`px-3 h-10 rounded-xl text-xs font-bold border transition cursor-pointer shrink-0 ${
                  filterDate === todayISOString
                    ? "bg-[#004B87] text-white border-[#004B87]"
                    : "bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100"
                }`}
              >
                Today
              </button>
              <button
                type="button"
                onClick={() => setFilterDate(null)}
                title="View All Dates"
                className={`px-3 h-10 rounded-xl text-xs font-bold border transition cursor-pointer shrink-0 ${
                  filterDate === null
                    ? "bg-[#004B87] text-white border-[#004B87]"
                    : "bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100"
                }`}
              >
                All Dates
              </button>
            </div>
          </div>

          {/* 2. Location Filter */}
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1">
              Filter by Location
            </label>
            <select
              value={filterLocationId}
              onChange={(e) =>
                setFilterLocationId(e.target.value === "all" ? "all" : Number(e.target.value))
              }
              className="w-full h-10 px-3 rounded-xl border border-slate-300 text-xs sm:text-sm font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#004B87] bg-white"
            >
              <option value="all">All Mess Locations</option>
              {locations.map((loc) => (
                <option key={loc.Id} value={loc.Id}>
                  {loc.Title}
                </option>
              ))}
            </select>
          </div>

          {/* 3. Search Field */}
          <div className="lg:col-span-2">
            <label className="block text-xs font-bold text-slate-700 mb-1">
              Search by Requester / Email
            </label>
            <div className="relative">
              <input
                type="text"
                placeholder="Search employee name or email..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full h-10 pl-9 pr-3 rounded-xl border border-slate-300 text-xs sm:text-sm font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#004B87]"
              />
              <svg
                viewBox="0 0 24 24"
                className="w-4 h-4 text-slate-400 absolute left-3 top-3 pointer-events-none"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <circle cx="11" cy="11" r="8" />
                <line x1="21" y1="21" x2="16.65" y2="16.65" />
              </svg>
            </div>
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
              No employee meal cancellations recorded for the selected date and location.
            </p>
          </div>
        ) : (
          <>
            {/* Desktop Table View (>= 640px) */}
            <div className="hidden sm:block overflow-x-auto">
              <table className="w-full text-left text-xs sm:text-sm">
                <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider text-[11px]">
                  <tr>
                    <th className="py-3.5 px-4 sm:px-6">Meal Date</th>
                    <th className="py-3.5 px-4">Meal Type</th>
                    <th className="py-3.5 px-4">Location</th>
                    <th className="py-3.5 px-4">Employee / Requester</th>
                    <th className="py-3.5 px-4">Email</th>
                    <th className="py-3.5 px-4 sm:px-6 text-right">Submitted Timestamp</th>
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
                        <td className="py-3.5 px-4 sm:px-6 font-bold text-slate-900 whitespace-nowrap">
                          {formattedMealDate}
                        </td>
                        <td className="py-3.5 px-4 whitespace-nowrap">
                          <span
                            className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-bold ${
                              item.MealType === "Breakfast"
                                ? "bg-amber-50 text-amber-700 border border-amber-200"
                                : item.MealType === "Lunch"
                                ? "bg-blue-50 text-blue-700 border border-blue-200"
                                : "bg-purple-50 text-purple-700 border border-purple-200"
                            }`}
                          >
                            {item.MealType}
                          </span>
                        </td>
                        <td className="py-3.5 px-4 font-semibold text-slate-800 whitespace-nowrap">
                          {item.Location?.Title || "—"}
                        </td>
                        <td className="py-3.5 px-4 font-bold text-slate-900 whitespace-nowrap">
                          {item.Requester?.Title || "Employee"}
                        </td>
                        <td className="py-3.5 px-4 text-slate-500 font-mono text-xs whitespace-nowrap">
                          {item.Requester?.EMail || "—"}
                        </td>
                        <td className="py-3.5 px-4 sm:px-6 text-right text-slate-500 whitespace-nowrap">
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
                    className="p-4 hover:bg-slate-50/70 transition-colors flex flex-col gap-2"
                  >
                    <div className="flex items-center justify-between">
                      <span className="font-extrabold text-slate-900 text-sm">
                        {formattedMealDate}
                      </span>
                      <span
                        className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-bold ${
                          item.MealType === "Breakfast"
                            ? "bg-amber-50 text-amber-700 border border-amber-200"
                            : item.MealType === "Lunch"
                            ? "bg-blue-50 text-blue-700 border border-blue-200"
                            : "bg-purple-50 text-purple-700 border border-purple-200"
                        }`}
                      >
                        {item.MealType}
                      </span>
                    </div>

                    <div className="flex items-center justify-between text-xs">
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
