import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem, MealType } from "../../models/IMealOptOut";
import { MealRequestService } from "../../services/mealRequestService";
import { validateMealCutOff } from "../../services/cutOffService";
import { BackButton } from "../Common/BackButton";

export interface IMyMealOptOutsProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  currentUserId: number;
  currentUserEmail: string;
  onNavigateCreate: () => void;
  onBack: () => void;
}

export const MyMealOptOuts: React.FC<IMyMealOptOutsProps> = ({
  context,
  spHttpClient,
  siteUrl,
  currentUserId,
  currentUserEmail,
  onNavigateCreate,
  onBack,
}) => {
  const [requests, setRequests] = React.useState<IMealOptOutRequestItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [cancellingId, setCancellingId] = React.useState<number | null>(null);
  const [feedback, setFeedback] = React.useState<{ type: "success" | "error"; message: string } | null>(null);

  // Filter State
  const [selectedMealFilter, setSelectedMealFilter] = React.useState<"ALL" | MealType>("ALL");
  const [searchLocation, setSearchLocation] = React.useState<string>("");

  const loadMyRequests = React.useCallback(async () => {
    setIsLoading(true);
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
    loadMyRequests().catch((err: Error) => {
      console.error(err);
    });
  }, [loadMyRequests]);

  // Handle Cancellation of an opt-out before deadline
  const handleCancel = async (item: IMealOptOutRequestItem): Promise<void> => {
    const mealDateFormatted = item.MealDate ? item.MealDate.split("T")[0] : "";
    const confirmCancel = window.confirm(
      `Are you sure you want to retract your opt-out for ${item.MealType} on ${mealDateFormatted}?`
    );
    if (!confirmCancel) return;

    setCancellingId(item.Id);
    setFeedback(null);

    const result = await MealRequestService.cancelOptOut(spHttpClient, siteUrl, item, context);
    setCancellingId(null);

    if (result.success) {
      setFeedback({
        type: "success",
        message: `Opt-out for ${item.MealType} on ${mealDateFormatted} successfully retracted.`,
      });
      loadMyRequests().catch((err: Error) => {
        console.error(err);
      });
    } else {
      setFeedback({
        type: "error",
        message: result.error || "Failed to retract opt-out.",
      });
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

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header */}
      <div className="mb-3.5 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 w-full">
        <div>
          <BackButton onClick={onBack} />
          <h2 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
            My Meal Opt-Out Requests
          </h2>
          <p className="text-xs sm:text-sm text-slate-500">
            View and manage your submitted meal advance cancellations across all Vedanta canteens.
          </p>
        </div>

        <button
          type="button"
          onClick={onNavigateCreate}
          className="inline-flex items-center gap-1.5 bg-[#0066CC] hover:bg-[#0052A3] text-white text-xs sm:text-sm font-bold px-5 py-2.5 rounded-full shadow-2xs hover:shadow-xs transition cursor-pointer self-start sm:self-auto"
        >
          <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <line x1="12" y1="5" x2="12" y2="19" />
            <line x1="5" y1="12" x2="19" y2="12" />
          </svg>
          <span>New Meal Opt-Out</span>
        </button>
      </div>

      {/* Feedback Banner */}
      {feedback && (
        <div
          className={`mb-3.5 p-3.5 rounded-xl border flex items-center justify-between text-xs sm:text-sm font-medium ${
            feedback.type === "success"
              ? "bg-emerald-50 border-emerald-200 text-emerald-800"
              : "bg-rose-50 border-rose-200 text-rose-800"
          }`}
        >
          <span>{feedback.message}</span>
          <button
            type="button"
            onClick={() => setFeedback(null)}
            className="text-slate-400 hover:text-slate-600 font-bold ml-2 cursor-pointer"
          >
            ✕
          </button>
        </div>
      )}

      {/* Filter and Search Bar Card */}
      <div
        style={{ height: "auto", flex: "0 0 auto" }}
        className="w-full bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm p-4 mb-3.5 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3"
      >
        {/* Meal Type Filter Pills */}
        <div className="flex items-center gap-2 overflow-x-auto pb-1 sm:pb-0">
          {(["ALL", "Breakfast", "Lunch", "Dinner"] as Array<"ALL" | MealType>).map((type) => {
            const isSelected = selectedMealFilter === type;
            return (
              <button
                key={type}
                type="button"
                onClick={() => setSelectedMealFilter(type)}
                className={`px-4 py-1.5 rounded-xl text-xs sm:text-sm font-bold transition cursor-pointer whitespace-nowrap ${
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
        <div className="relative min-w-[240px] sm:w-72">
          <input
            type="text"
            placeholder="Search by mess location..."
            value={searchLocation}
            onChange={(e) => setSearchLocation(e.target.value)}
            className="w-full h-9 pl-9 pr-3 rounded-xl border border-slate-300 text-xs sm:text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#004B87]"
          />
          <svg
            viewBox="0 0 24 24"
            className="w-4 h-4 text-slate-400 absolute left-3 top-2.5 fill-none stroke-current"
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
          <div className="w-full overflow-x-auto">
            <table className="w-full text-left text-xs sm:text-sm">
              <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider text-[10.5px] sm:text-xs">
                <tr>
                  <th className="py-3 px-4 sm:px-6">Meal Date</th>
                  <th className="py-3 px-4">Meal Type</th>
                  <th className="py-3 px-4">Canteen Location</th>
                  <th className="py-3 px-4">Submitted On</th>
                  <th className="py-3 px-4 text-center">Cut-Off Status</th>
                  <th className="py-3 px-4 sm:px-6 text-right">Action</th>
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

                  const cutOffCheck = validateMealCutOff(mealDateObj, item.MealType);
                  const canCancel = cutOffCheck.isValid;

                  return (
                    <tr key={item.Id} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-2.5 px-4 sm:px-5 font-bold text-slate-900 whitespace-nowrap">
                        {formattedMealDate}
                      </td>
                      <td className="py-2.5 px-3 whitespace-nowrap">
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
                      <td className="py-2.5 px-3 font-semibold text-slate-800 whitespace-nowrap">
                        {item.Location?.Title || "—"}
                      </td>
                      <td className="py-2.5 px-3 text-slate-500 whitespace-nowrap text-[11px]">
                        {formattedCreated}
                      </td>
                      <td className="py-2.5 px-3 text-center whitespace-nowrap">
                        {canCancel ? (
                          <span className="inline-flex items-center gap-1 text-[10.5px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">
                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                            Active (Modifiable)
                          </span>
                        ) : (
                          <span className="inline-flex items-center text-[10.5px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
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
                            className="text-xs font-bold text-rose-600 hover:text-rose-800 hover:underline cursor-pointer disabled:opacity-50"
                          >
                            {cancellingId === item.Id ? "Cancelling..." : "Retract Opt-Out"}
                          </button>
                        ) : (
                          <span className="text-xs text-slate-400 font-normal">Past Cut-off</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
