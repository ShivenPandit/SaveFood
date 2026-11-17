import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem } from "../../models/IMealOptOut";
import { ILocationMasterItem } from "../../models/ILocation";
import { MealRequestService } from "../../services/mealRequestService";
import { getCurrentISTDate } from "../../services/cutOffService";
import { BackButton } from "../Common/BackButton";

export interface IAllMealOptOutsProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  onBack: () => void;
}

export const AllMealOptOuts: React.FC<IAllMealOptOutsProps> = ({
  spHttpClient,
  siteUrl,
  onBack,
}) => {
  const [requests, setRequests] = React.useState<IMealOptOutRequestItem[]>([]);
  const [locations, setLocations] = React.useState<ILocationMasterItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  // Filters: Date defaults to today in IST
  const todayIST = getCurrentISTDate();
  const todayISOString = todayIST.toISOString().split("T")[0];

  const [filterDate, setFilterDate] = React.useState<string | null>(todayISOString);
  const [filterLocationId, setFilterLocationId] = React.useState<number | "all">("all");
  const [searchQuery, setSearchQuery] = React.useState<string>("");

  // Pagination
  const [nextUrl, setNextUrl] = React.useState<string | null>(null);
  const [isLoadingMore, setIsLoadingMore] = React.useState<boolean>(false);

  // 1. Load Locations on mount
  React.useEffect(() => {
    MealRequestService.getAllLocations(spHttpClient, siteUrl).then((locs) => {
      setLocations(locs);
    });
  }, [spHttpClient, siteUrl]);

  // 2. Fetch Requests based on active filters
  const fetchRequests = React.useCallback(
    async (resetPagination: boolean = true) => {
      if (resetPagination) setIsLoading(true);

      const paged = await MealRequestService.getAllMealOptOuts(spHttpClient, siteUrl, {
        dateISO: filterDate,
        locationId: filterLocationId === "all" ? null : Number(filterLocationId),
        pageSize: 50,
      });

      setRequests(paged.items);
      setNextUrl(paged.nextUrl);
      setIsLoading(false);
    },
    [spHttpClient, siteUrl, filterDate, filterLocationId]
  );

  React.useEffect(() => {
    fetchRequests(true);
  }, [fetchRequests]);

  // Load next page
  const handleLoadMore = async () => {
    if (!nextUrl || isLoadingMore) return;
    setIsLoadingMore(true);

    const paged = await MealRequestService.getAllMealOptOuts(spHttpClient, siteUrl, {
      nextUrl,
    });

    setRequests((prev) => [...prev, ...paged.items]);
    setNextUrl(paged.nextUrl);
    setIsLoadingMore(false);
  };

  // Filter in-memory for immediate instant search by name or summary
  const filteredRequests = React.useMemo(() => {
    if (!searchQuery.trim()) return requests;
    const query = searchQuery.toLowerCase();
    return requests.filter((r) => {
      const requesterName = (r.Requester?.Title || "").toLowerCase();
      const requesterEmail = (r.Requester?.EMail || "").toLowerCase();
      const summary = (r.Title || "").toLowerCase();
      const loc = (r.Location?.Title || "").toLowerCase();
      return (
        requesterName.indexOf(query) !== -1 ||
        requesterEmail.indexOf(query) !== -1 ||
        summary.indexOf(query) !== -1 ||
        loc.indexOf(query) !== -1
      );
    });
  }, [requests, searchQuery]);

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header */}
      <div className="mb-3.5 flex flex-col md:flex-row md:items-center md:justify-between gap-3">
        <div>
          <BackButton onClick={onBack} />
          <div className="flex items-center gap-2.5">
            <h2 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
              All Meal Opt-Out Requests
            </h2>
            <span className="px-2.5 py-0.5 rounded-full text-[10px] font-extrabold bg-[#004B87] text-white">
              Admin Roster
            </span>
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
                title="Reset to Today"
                className={`px-3 h-10 rounded-xl text-xs font-bold border transition cursor-pointer shrink-0 ${
                  filterDate === todayISOString
                    ? "bg-blue-50 border-blue-300 text-[#004B87]"
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
          <div className="overflow-x-auto">
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
                {filteredRequests.map((item) => {
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
        )}

        {/* Load More Pagination */}
        {nextUrl && (
          <div className="p-4 bg-slate-50 border-t border-slate-200 text-center">
            <button
              type="button"
              onClick={handleLoadMore}
              disabled={isLoadingMore}
              className="px-6 py-2 rounded-full bg-[#004B87] hover:bg-[#003865] text-white text-xs font-bold transition cursor-pointer disabled:opacity-50"
            >
              {isLoadingMore ? "Loading More..." : "Load Next Page (Paging Active)"}
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
