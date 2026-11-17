import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IMealOptOutRequestItem } from "../../models/IMealOptOut";
import { MealRequestService } from "../../services/mealRequestService";
import { BackButton } from "../Common/BackButton";

export interface IReportsDashboardProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  onBack: () => void;
}

export const ReportsDashboard: React.FC<IReportsDashboardProps> = ({
  spHttpClient,
  siteUrl,
  onBack,
}) => {
  const [data, setData] = React.useState<IMealOptOutRequestItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  // Load all opt-outs for analytics
  React.useEffect(() => {
    let isMounted = true;
    MealRequestService.getAllMealOptOuts(spHttpClient, siteUrl, {
      dateISO: null, // all records
      pageSize: 500,
    }).then((res) => {
      if (isMounted) {
        setData(res.items);
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
      {/* Top Header */}
      <div className="mb-3.5 flex flex-col md:flex-row md:items-center md:justify-between gap-3">
        <div>
          <BackButton onClick={onBack} />
          <div className="flex items-center gap-2.5">
            <h2 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
              Reports &amp; Sustainability Analytics
            </h2>
            <span className="px-2.5 py-0.5 rounded-full text-[10px] font-extrabold bg-emerald-600 text-white">
              Admin Analytics
            </span>
          </div>
          <p className="text-xs sm:text-sm text-slate-500 mt-0.5">
            Real-time catering optimization, food waste prevention metrics, and location trends.
          </p>
        </div>
      </div>

      {isLoading ? (
        <div
          style={{ height: "auto", flex: "0 0 auto", padding: "24px" }}
          className="bg-white rounded-2xl border-2 border-slate-200/90 text-center shadow-sm"
        >
          <div className="inline-block animate-spin rounded-full h-8 w-8 border-4 border-[#004B87] border-t-transparent" />
          <p className="mt-3 text-xs sm:text-sm font-semibold text-slate-500">
            Calculating sustainability analytics...
          </p>
        </div>
      ) : (
        <div className="space-y-4">
          {/* 4 Metric Cards Strip */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            {/* 1. Total Opt-Outs */}
            <div className="bg-white rounded-2xl border-2 border-slate-200/90 p-5 shadow-xs">
              <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                Total Meals Saved
              </span>
              <p className="text-2xl sm:text-3xl font-black text-[#004B87] mt-1">
                {stats.total.toLocaleString()}
              </p>
              <p className="text-xs text-slate-500 mt-1">Direct catering cancellations</p>
            </div>

            {/* 2. Food Waste Prevented */}
            <div className="bg-white rounded-2xl border-2 border-slate-200/90 p-5 shadow-xs">
              <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                Food Waste Prevented
              </span>
              <p className="text-2xl sm:text-3xl font-black text-emerald-600 mt-1">
                {stats.kgSaved.toLocaleString()} kg
              </p>
              <p className="text-xs text-slate-500 mt-1">Based on ~400g/meal average</p>
            </div>

            {/* 3. CO2 Emissions Avoided */}
            <div className="bg-white rounded-2xl border-2 border-slate-200/90 p-5 shadow-xs">
              <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                CO₂e Footprint Reduced
              </span>
              <p className="text-2xl sm:text-3xl font-black text-teal-600 mt-1">
                {stats.co2AvoidedKg.toLocaleString()} kg
              </p>
              <p className="text-xs text-slate-500 mt-1">Environmental footprint saved</p>
            </div>

            {/* 4. Active Dining Locations */}
            <div className="bg-white rounded-2xl border-2 border-slate-200/90 p-5 shadow-xs">
              <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                Canteen Locations
              </span>
              <p className="text-2xl sm:text-3xl font-black text-purple-700 mt-1">
                {Object.keys(stats.locationMap).length}
              </p>
              <p className="text-xs text-slate-500 mt-1">Participating plant facilities</p>
            </div>
          </div>

          {/* Breakdown Grid: Meal Types & Locations */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Meal Type Breakdown Card */}
            <div className="bg-white rounded-2xl border-2 border-slate-200/90 p-5 sm:p-6 shadow-xs">
              <h4 className="text-base font-extrabold text-slate-900 mb-4">
                Opt-Out Distribution by Meal Type
              </h4>

              <div className="space-y-4">
                {/* Breakfast */}
                <div>
                  <div className="flex items-center justify-between text-xs sm:text-sm font-bold mb-1">
                    <span className="text-amber-800">🍳 Breakfast</span>
                    <span className="text-slate-800">
                      {stats.breakfast} ({stats.total > 0 ? Math.round((stats.breakfast / stats.total) * 100) : 0}%)
                    </span>
                  </div>
                  <div className="h-3 w-full bg-slate-100 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-amber-500 rounded-full transition-all duration-500"
                      style={{
                        width: `${stats.total > 0 ? (stats.breakfast / stats.total) * 100 : 0}%`,
                      }}
                    />
                  </div>
                </div>

                {/* Lunch */}
                <div>
                  <div className="flex items-center justify-between text-xs sm:text-sm font-bold mb-1">
                    <span className="text-blue-800">🥗 Lunch</span>
                    <span className="text-slate-800">
                      {stats.lunch} ({stats.total > 0 ? Math.round((stats.lunch / stats.total) * 100) : 0}%)
                    </span>
                  </div>
                  <div className="h-3 w-full bg-slate-100 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-blue-600 rounded-full transition-all duration-500"
                      style={{
                        width: `${stats.total > 0 ? (stats.lunch / stats.total) * 100 : 0}%`,
                      }}
                    />
                  </div>
                </div>

                {/* Dinner */}
                <div>
                  <div className="flex items-center justify-between text-xs sm:text-sm font-bold mb-1">
                    <span className="text-purple-800">🍲 Dinner</span>
                    <span className="text-slate-800">
                      {stats.dinner} ({stats.total > 0 ? Math.round((stats.dinner / stats.total) * 100) : 0}%)
                    </span>
                  </div>
                  <div className="h-3 w-full bg-slate-100 rounded-full overflow-hidden">
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

            {/* Location Breakdown Card */}
            <div className="bg-white rounded-2xl border-2 border-slate-200/90 p-5 sm:p-6 shadow-xs">
              <h4 className="text-base font-extrabold text-slate-900 mb-4">
                Headcount Savings by Location
              </h4>

              {Object.keys(stats.locationMap).length === 0 ? (
                <p className="text-xs text-slate-400">No location data available yet.</p>
              ) : (
                <div className="space-y-3.5">
                  {Object.keys(stats.locationMap).map((loc) => {
                    const count = stats.locationMap[loc];
                    const pct = stats.total > 0 ? Math.round((count / stats.total) * 100) : 0;
                    return (
                      <div key={loc}>
                        <div className="flex items-center justify-between text-xs sm:text-sm font-semibold mb-1">
                          <span className="text-slate-800">{loc}</span>
                          <span className="text-slate-600 font-bold">
                            {count} ({pct}%)
                          </span>
                        </div>
                        <div className="h-2.5 w-full bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-emerald-500 rounded-full transition-all duration-500"
                            style={{ width: `${pct}%` }}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
