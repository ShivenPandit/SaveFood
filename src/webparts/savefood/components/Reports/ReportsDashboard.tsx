import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem } from "../../models/IMealOptOut";
import { MealRequestService } from "../../services/mealRequestService";
import { BackButton } from "../Common/BackButton";
import { AllMealOptOuts } from "../AllRequests/AllMealOptOuts";

export type ReportTab = "dashboard" | "requests";

export interface IReportsDashboardProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  onBack: () => void;
  initialTab?: ReportTab;
}

export const ReportsDashboard: React.FC<IReportsDashboardProps> = ({
  context,
  spHttpClient,
  siteUrl,
  onBack,
  initialTab = "dashboard",
}) => {
  const cachedInitial = React.useMemo(() => {
    return MealRequestService.getCachedAllMealOptOuts() || [];
  }, []);

  const [activeTab, setActiveTab] = React.useState<ReportTab>(initialTab);
  const [data, setData] = React.useState<IMealOptOutRequestItem[]>(cachedInitial);
  const [isLoading, setIsLoading] = React.useState<boolean>(cachedInitial.length === 0);

  // Load all opt-outs for analytics
  React.useEffect(() => {
    let isMounted = true;
    MealRequestService.getAllMealOptOuts(spHttpClient, siteUrl, {
      dateISO: null, // all records
      pageSize: 4800,
    })
      .then((res) => {
        if (isMounted) {
          setData(res.items);
          setIsLoading(false);
        }
      })
      .catch((err: Error) => {
        console.error(err);
        if (isMounted) {
          setIsLoading(false);
        }
      });

    return () => {
      isMounted = false;
    };
  }, [spHttpClient, siteUrl]);

  // Aggregate Metrics
  const stats = React.useMemo(() => {
    const total = data.length;

    // Meal type counts
    const breakfast = data.filter((d) => d.MealType === "Breakfast").length;
    const lunch = data.filter((d) => d.MealType === "Lunch").length;
    const dinner = data.filter((d) => d.MealType === "Dinner").length;

    // Location counts
    const locationMap: Record<string, number> = {};
    data.forEach((d) => {
      const loc = d.Location?.Title || "Unspecified";
      locationMap[loc] = (locationMap[loc] || 0) + 1;
    });

    // Environmental metrics (approx. 400g edible food preserved per meal skipped)
    const kgSaved = Math.round(total * 0.4);
    // Approx 1.2 kg CO2 equivalent per kg food waste prevented
    const co2AvoidedKg = Math.round(kgSaved * 1.2);

    return {
      total,
      breakfast,
      lunch,
      dinner,
      locationMap,
      kgSaved,
      co2AvoidedKg,
    };
  }, [data]);

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header Row: BackButton on Left, Switcher Tabs on Right */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2.5 mb-2">
        <BackButton onClick={onBack} />

        {/* Modern Segmented Control Pill Tabs */}
        <div className="inline-flex items-center p-1 bg-slate-100/90 rounded-xl border border-slate-200/90 shadow-2xs self-start sm:self-auto">
          <button
            type="button"
            onClick={() => setActiveTab("dashboard")}
            className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-[13.5px] sm:text-[14px] font-bold transition-all duration-150 cursor-pointer select-none whitespace-nowrap ${
              activeTab === "dashboard"
                ? "bg-[#004B87] text-white shadow-xs"
                : "text-slate-600 hover:text-slate-900 hover:bg-slate-200/60"
            }`}
          >
            <svg
              viewBox="0 0 24 24"
              className="w-4 h-4 shrink-0 fill-none stroke-current"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <rect x="3" y="3" width="7" height="9" rx="1" />
              <rect x="14" y="3" width="7" height="5" rx="1" />
              <rect x="14" y="12" width="7" height="9" rx="1" />
              <rect x="3" y="16" width="7" height="5" rx="1" />
            </svg>
            <span>Report Dashboard</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab("requests")}
            className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-[13.5px] sm:text-[14px] font-bold transition-all duration-150 cursor-pointer select-none whitespace-nowrap ${
              activeTab === "requests"
                ? "bg-[#004B87] text-white shadow-xs"
                : "text-slate-600 hover:text-slate-900 hover:bg-slate-200/60"
            }`}
          >
            <svg
              viewBox="0 0 24 24"
              className="w-4 h-4 shrink-0 fill-none stroke-current"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <line x1="8" y1="6" x2="21" y2="6" />
              <line x1="8" y1="12" x2="21" y2="12" />
              <line x1="8" y1="18" x2="21" y2="18" />
              <line x1="3" y1="6" x2="3.01" y2="6" />
              <line x1="3" y1="12" x2="3.01" y2="12" />
              <line x1="3" y1="18" x2="3.01" y2="18" />
            </svg>
            <span>All Meal Opt-Out Requests</span>
          </button>
        </div>
      </div>

      {/* Tab 2 Content: All Meal Opt-Out Requests Roster */}
      {activeTab === "requests" && (
        <AllMealOptOuts
          context={context}
          spHttpClient={spHttpClient}
          siteUrl={siteUrl}
          onBack={onBack}
          hideBackButton={true}
        />
      )}

      {/* Tab 1 Content: Report Dashboard & Sustainability Analytics */}
      {activeTab === "dashboard" && (
        <>
          {/* Compact Single-Row Page Header & Summary Strip */}
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3 mb-3">
            {/* Left: Page Title & Subtitle */}
            <div>
              <h2 className="text-[22px] sm:text-[24px] font-bold text-[#0f172a] tracking-tight">
                Reports &amp; Sustainability Analytics
              </h2>
              <p className="text-[13.5px] text-[#475569] leading-snug mt-0.5">
                Real-time catering optimization, food waste prevention metrics, and location trends.
              </p>
            </div>

            {/* Right: Sleek Compact KPI Pill Card */}
            <div className="bg-white rounded-xl border border-slate-200/90 shadow-2xs px-3.5 py-2 flex items-center gap-3 sm:gap-4 flex-wrap self-start lg:self-auto shrink-0">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-lg bg-blue-50 text-[#004B87] flex items-center justify-center border border-blue-100 shrink-0">
                  <svg viewBox="0 0 24 24" className="w-5 h-5 fill-none stroke-current" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
                    <circle cx="9" cy="7" r="4" />
                    <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
                    <path d="M16 3.13a4 4 0 0 1 0 7.75" />
                  </svg>
                </div>
                <div>
                  <span className="text-[10.5px] font-bold text-slate-500 uppercase tracking-wider block leading-none">
                    TOTAL HEADCOUNT SAVED
                  </span>
                  <span className="text-[22px] sm:text-[24px] font-extrabold text-[#003865] tracking-tight leading-none">
                    {stats.total.toLocaleString()}
                  </span>
                </div>
              </div>

              <div className="h-6 w-px bg-slate-200 hidden sm:block" />

              {/* Inline Shift Chips */}
              <div className="flex items-center gap-1.5 flex-wrap">
                <span className="bg-amber-50 text-amber-800 text-[12px] font-semibold px-2.5 py-1 rounded-lg border border-amber-200/80">
                  🍳 Breakfast: {stats.breakfast}
                </span>
                <span className="bg-blue-50 text-blue-800 text-[12px] font-semibold px-2.5 py-1 rounded-lg border border-blue-200/80">
                  🥗 Lunch: {stats.lunch}
                </span>
                <span className="bg-purple-50 text-purple-800 text-[12px] font-semibold px-2.5 py-1 rounded-lg border border-purple-200/80">
                  🍲 Dinner: {stats.dinner}
                </span>
              </div>
            </div>
          </div>

          {isLoading ? (
            <div
              style={{ height: "auto", flex: "0 0 auto", padding: "32px" }}
              className="bg-white rounded-2xl border-2 border-slate-200/90 text-center shadow-sm"
            >
              <div className="inline-block animate-spin rounded-full h-8 w-8 border-4 border-[#004B87] border-t-transparent" />
              <p className="mt-3 text-[14.5px] font-semibold text-slate-500">
                Calculating sustainability analytics...
              </p>
            </div>
          ) : (
            /* Balanced Two-Column Bento Grid: Meal Distribution & Location Savings */
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-3.5 sm:gap-4 items-stretch">
              {/* Left Card: Opt-Out Distribution by Meal Type (5 cols on lg) */}
              <div className="lg:col-span-5 bg-white rounded-2xl border-2 border-slate-200/90 p-4 sm:p-4.5 shadow-xs flex flex-col justify-between">
                <div>
                  <div className="flex items-center justify-between mb-3">
                    <div>
                      <h4 className="text-[16px] sm:text-[17px] font-bold text-slate-900">
                        Opt-Out Distribution by Meal Type
                      </h4>
                      <p className="text-[12.5px] text-slate-500 mt-0.5">
                        Catering cancellations across operational shifts
                      </p>
                    </div>
                    <span className="text-[12px] font-semibold text-slate-600 bg-slate-100 px-2.5 py-0.5 rounded-md">
                      3 Shifts
                    </span>
                  </div>

                  {/* Segmented Proportional Progress Track */}
                  {stats.total > 0 && (
                    <div className="mb-3">
                      <div className="h-2.5 w-full bg-slate-100 rounded-full overflow-hidden flex">
                        <div
                          className="bg-amber-500 transition-all duration-500"
                          style={{ width: `${(stats.breakfast / stats.total) * 100}%` }}
                          title={`Breakfast: ${stats.breakfast}`}
                        />
                        <div
                          className="bg-blue-600 transition-all duration-500"
                          style={{ width: `${(stats.lunch / stats.total) * 100}%` }}
                          title={`Lunch: ${stats.lunch}`}
                        />
                        <div
                          className="bg-purple-600 transition-all duration-500"
                          style={{ width: `${(stats.dinner / stats.total) * 100}%` }}
                          title={`Dinner: ${stats.dinner}`}
                        />
                      </div>
                    </div>
                  )}

                  <div className="space-y-2">
                    {/* Breakfast */}
                    <div className="p-2.5 rounded-xl border border-amber-100 bg-amber-50/30">
                      <div className="flex items-center justify-between text-[13.5px] font-semibold mb-1">
                        <span className="text-amber-800 flex items-center gap-1.5">
                          <span>🍳</span> Breakfast
                        </span>
                        <span className="text-slate-900 font-bold">
                          {stats.breakfast} <span className="text-slate-500 font-normal text-[12.5px]">({stats.total > 0 ? Math.round((stats.breakfast / stats.total) * 100) : 0}%)</span>
                        </span>
                      </div>
                      <div className="h-2 w-full bg-amber-100/70 rounded-full overflow-hidden">
                        <div
                          className="h-full bg-amber-500 rounded-full transition-all duration-500"
                          style={{
                            width: `${stats.total > 0 ? (stats.breakfast / stats.total) * 100 : 0}%`,
                          }}
                        />
                      </div>
                    </div>

                    {/* Lunch */}
                    <div className="p-2.5 rounded-xl border border-blue-100 bg-blue-50/30">
                      <div className="flex items-center justify-between text-[13.5px] font-semibold mb-1">
                        <span className="text-blue-800 flex items-center gap-1.5">
                          <span>🥗</span> Lunch
                        </span>
                        <span className="text-slate-900 font-bold">
                          {stats.lunch} <span className="text-slate-500 font-normal text-[12.5px]">({stats.total > 0 ? Math.round((stats.lunch / stats.total) * 100) : 0}%)</span>
                        </span>
                      </div>
                      <div className="h-2 w-full bg-blue-100/70 rounded-full overflow-hidden">
                        <div
                          className="h-full bg-blue-600 rounded-full transition-all duration-500"
                          style={{
                            width: `${stats.total > 0 ? (stats.lunch / stats.total) * 100 : 0}%`,
                          }}
                        />
                      </div>
                    </div>

                    {/* Dinner */}
                    <div className="p-2.5 rounded-xl border border-purple-100 bg-purple-50/30">
                      <div className="flex items-center justify-between text-[13.5px] font-semibold mb-1">
                        <span className="text-purple-800 flex items-center gap-1.5">
                          <span>🍲</span> Dinner
                        </span>
                        <span className="text-slate-900 font-bold">
                          {stats.dinner} <span className="text-slate-500 font-normal text-[12.5px]">({stats.total > 0 ? Math.round((stats.dinner / stats.total) * 100) : 0}%)</span>
                        </span>
                      </div>
                      <div className="h-2 w-full bg-purple-100/70 rounded-full overflow-hidden">
                        <div
                          className="h-full bg-purple-600 rounded-full transition-all duration-500"
                          style={{
                            width: `${stats.total > 0 ? (stats.dinner / stats.total) * 100 : 0}%`,
                          }}
                        />
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Right Card: Headcount Savings by Location (7 cols on lg) */}
              <div className="lg:col-span-7 bg-white rounded-2xl border-2 border-slate-200/90 p-4 sm:p-4.5 shadow-xs">
                <div className="flex items-center justify-between mb-3">
                  <div>
                    <h4 className="text-[16px] sm:text-[17px] font-bold text-slate-900">
                      Headcount Savings by Location
                    </h4>
                    <p className="text-[12.5px] text-slate-500 mt-0.5">
                      Catering reduction ranked across facilities
                    </p>
                  </div>
                  <span className="text-[12px] font-semibold text-slate-600 bg-slate-100 px-2.5 py-0.5 rounded-md">
                    {Object.keys(stats.locationMap).length} Facilities
                  </span>
                </div>

                {Object.keys(stats.locationMap).length === 0 ? (
                  <div className="py-8 text-center text-slate-400 text-[14px]">
                    No location data available yet.
                  </div>
                ) : (
                  <div className="space-y-1.5">
                    {Object.keys(stats.locationMap)
                      .sort((a, b) => (stats.locationMap[b] || 0) - (stats.locationMap[a] || 0))
                      .map((loc, index) => {
                        const count = stats.locationMap[loc] || 0;
                        const pct = stats.total > 0 ? Math.round((count / stats.total) * 100) : 0;
                        return (
                          <div
                            key={loc}
                            className="px-3 py-1.5 rounded-xl border border-slate-200/80 bg-slate-50/50 flex items-center gap-2.5"
                          >
                            <span
                              className={`w-5 h-5 rounded-md flex items-center justify-center text-[11px] font-black shrink-0 ${
                                index === 0
                                  ? "bg-amber-100 text-amber-900 border border-amber-300"
                                  : index === 1
                                  ? "bg-slate-200 text-slate-800 border border-slate-300"
                                  : index === 2
                                  ? "bg-orange-100 text-orange-900 border border-orange-300"
                                  : "bg-slate-100 text-slate-600 border border-slate-200"
                              }`}
                            >
                              #{index + 1}
                            </span>
                            <span className="text-slate-900 font-bold text-[13px] w-24 sm:w-28 truncate shrink-0">
                              {loc}
                            </span>
                            <div className="flex-1 h-2 bg-slate-200/70 rounded-full overflow-hidden">
                              <div
                                className="h-full bg-gradient-to-r from-emerald-500 to-teal-600 rounded-full transition-all duration-500"
                                style={{ width: `${pct}%` }}
                              />
                            </div>
                            <span className="text-slate-800 font-bold text-[12.5px] shrink-0 text-right min-w-[55px]">
                              {count} <span className="text-slate-500 font-normal text-[11.5px]">({pct}%)</span>
                            </span>
                          </div>
                        );
                      })}
                  </div>
                )}
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
};
