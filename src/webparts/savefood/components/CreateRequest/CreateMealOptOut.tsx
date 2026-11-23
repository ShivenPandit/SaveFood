import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { ILocationMasterItem } from "../../models/ILocation";
import { MealType } from "../../models/IMealOptOut";
import { MealRequestService } from "../../services/mealRequestService";
import {
  validateMealCutOff,
  getCutOffDateTimeIST,
  CUT_OFF_RULES,
  getCurrentISTDate,
  formatISTDateToISO,
} from "../../services/cutOffService";
import { BackButton } from "../Common/BackButton";

export interface ICreateMealOptOutProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  currentUserId: number;
  currentUserName: string;
  currentUserEmail: string;
  onSuccess: (message?: string) => void;
  onCancel: () => void;
}

export const CreateMealOptOut: React.FC<ICreateMealOptOutProps> = ({
  spHttpClient,
  siteUrl,
  currentUserId,
  currentUserName,
  currentUserEmail,
  onSuccess,
  onCancel,
}) => {
  // Form State
  const [locations, setLocations] = React.useState<ILocationMasterItem[]>([]);
  const [selectedLocationId, setSelectedLocationId] = React.useState<number | "">("");
  const [selectedMealType, setSelectedMealType] = React.useState<MealType | "">("");

  // Date selection: Do not pre-select today's date; initialize as empty ("")
  const todayIST = getCurrentISTDate();
  const todayISOString = formatISTDateToISO(todayIST);
  const [selectedDate, setSelectedDate] = React.useState<string>("");

  // Status & Validation State
  const [isLoadingLocations, setIsLoadingLocations] = React.useState<boolean>(true);
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [serverError, setServerError] = React.useState<string | null>(null);

  // Field Touched / Submit Attempt tracking for inline validation
  const [touched, setTouched] = React.useState<{
    location?: boolean;
    date?: boolean;
    mealType?: boolean;
  }>({});
  const [hasAttemptedSubmit, setHasAttemptedSubmit] = React.useState<boolean>(false);

  // 1. Fetch active locations on load (no default selection)
  React.useEffect(() => {
    let isMounted = true;
    MealRequestService.getActiveLocations(spHttpClient, siteUrl)
      .then((items) => {
        if (isMounted) {
          setLocations(items);
          setIsLoadingLocations(false);
        }
      })
      .catch(() => {
        if (isMounted) setIsLoadingLocations(false);
      });

    return () => {
      isMounted = false;
    };
  }, [spHttpClient, siteUrl]);

  // 2. Evaluate cut-off window individually for each meal button (Breakfast, Lunch, Dinner) based on selectedDate
  const mealTypeOptions: MealType[] = ["Breakfast", "Lunch", "Dinner"];

  const mealValidation = React.useMemo(() => {
    const result: Record<MealType, { isExpired: boolean; message?: string }> = {
      Breakfast: { isExpired: false },
      Lunch: { isExpired: false },
      Dinner: { isExpired: false },
    };

    if (!selectedDate) {
      return result;
    }

    for (const type of mealTypeOptions) {
      const check = validateMealCutOff(selectedDate, type);
      if (!check.isValid) {
        result[type] = {
          isExpired: true,
          message: check.errorMessage,
        };
      }
    }

    return result;
  }, [selectedDate]);

  // Clear selectedMealType whenever the user picks a date whose cut-off has already passed for that meal
  React.useEffect(() => {
    if (selectedDate && selectedMealType && mealValidation[selectedMealType]?.isExpired) {
      setSelectedMealType("");
    }
  }, [selectedDate, selectedMealType, mealValidation]);

  // Selected Location Object
  const selectedLocationObj = React.useMemo(() => {
    return locations.find((l) => l.Id === Number(selectedLocationId));
  }, [locations, selectedLocationId]);

  // Compute current cut-off display info for the selected meal type and date
  const cutOffDisplay = React.useMemo(() => {
    if (!selectedDate || !selectedMealType) return null;
    const cutOff = getCutOffDateTimeIST(selectedDate, selectedMealType);
    return {
      deadline: cutOff.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit", hour12: true }),
      date: cutOff.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }),
      rule: CUT_OFF_RULES[selectedMealType].description,
    };
  }, [selectedDate, selectedMealType]);

  // Computed Inline Validation Errors
  const locationError =
    (hasAttemptedSubmit || touched.location) && !selectedLocationId
      ? "Please select a canteen / mess location."
      : null;

  const dateError =
    (hasAttemptedSubmit || touched.date) && !selectedDate
      ? "Please select a meal date."
      : null;

  const mealTypeError =
    (hasAttemptedSubmit || touched.mealType) && !selectedMealType
      ? "Please select a meal type (Breakfast, Lunch, or Dinner)."
      : null;

  // Handle Form Submission
  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    setHasAttemptedSubmit(true);
    setServerError(null);

    const isSelectedMealExpired = selectedMealType ? mealValidation[selectedMealType]?.isExpired : false;

    if (!selectedLocationId || !selectedDate || !selectedMealType || isSelectedMealExpired) {
      return;
    }

    setIsSubmitting(true);

    const locationName = selectedLocationObj ? selectedLocationObj.Title : "Mess";
    const titleSummary = `${locationName} - ${selectedMealType} - ${selectedDate}`;

    const result = await MealRequestService.createMealOptOut(spHttpClient, siteUrl, {
      Title: titleSummary,
      LocationId: Number(selectedLocationId),
      LocationTitle: locationName,
      MealType: selectedMealType,
      MealDate: `${selectedDate}T00:00:00Z`,
      RequesterId: currentUserId,
      RequesterEmail: currentUserEmail,
    });

    setIsSubmitting(false);

    if (result.success) {
      // Immediately transition to My Requests and display modern floating toast on destination screen
      const successMsg = `Opt-out submitted successfully for ${selectedMealType} on ${selectedDate}.`;
      onSuccess(successMsg);
    } else {
      setServerError(result.error || "Failed to submit request. Please try again.");
    }
  };

  const isCurrentSelectionExpired = !!selectedMealType && mealValidation[selectedMealType]?.isExpired;

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header */}
      <div className="mb-2.5 flex items-center justify-between w-full">
        <div>
          <BackButton onClick={onCancel} className="mb-1.5" />
          <h2 className="text-[26px] sm:text-[28px] font-bold text-[#0f172a] tracking-tight">
            Create Meal Opt-Out Request
          </h2>
          <p className="text-[15px] text-[#475569] leading-relaxed mt-1">
            Submit advance meal cancellations for upcoming plant shifts, official travel, or leave.
          </p>
        </div>
      </div>

      {/* Main Clean Form Card */}
      <div
        style={{ height: "auto", flex: "0 0 auto" }}
        className="w-full bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm p-6 sm:p-7"
      >
        {/* Server Error Alert */}
        {serverError && (
          <div className="mb-4 p-4 rounded-xl bg-rose-50 border border-rose-200 flex items-center gap-2.5 text-rose-800 text-[14px] font-medium">
            <svg viewBox="0 0 24 24" className="w-5 h-5 text-rose-600 shrink-0" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="12" cy="12" r="10" />
              <line x1="12" y1="8" x2="12" y2="12" />
              <line x1="12" y1="16" x2="12.01" y2="16" />
            </svg>
            <span>{serverError}</span>
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-5">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* 1. Location Selection (Top-Left) */}
            <div>
              <label className="block text-[14.5px] font-semibold text-[#1e293b] mb-2">
                Canteen / Mess Location <span className="text-rose-500">*</span>
              </label>
              {isLoadingLocations ? (
                <div className="min-h-[46px] h-[46px] w-full rounded-xl bg-slate-100 animate-pulse" />
              ) : (
                <select
                  value={selectedLocationId}
                  onChange={(e) => {
                    const val = e.target.value ? Number(e.target.value) : "";
                    setSelectedLocationId(val);
                    setTouched((prev) => ({ ...prev, location: true }));
                  }}
                  onBlur={() => setTouched((prev) => ({ ...prev, location: true }))}
                  disabled={isSubmitting}
                  className={`w-full min-h-[46px] h-[46px] px-4 rounded-xl border bg-white text-slate-800 text-[15px] font-medium focus:outline-none focus:ring-2 transition ${
                    locationError
                      ? "border-rose-500 focus:ring-rose-500/20 bg-rose-50/20"
                      : "border-slate-300 focus:ring-[#004B87] focus:border-transparent"
                  }`}
                >
                  <option value="">-- Select Canteen / Mess Location --</option>
                  {locations.map((loc) => (
                    <option key={loc.Id} value={loc.Id}>
                      {loc.Title}
                    </option>
                  ))}
                </select>
              )}
              {locationError ? (
                <p className="text-[13px] text-rose-600 font-semibold mt-1.5 flex items-center gap-1.5">
                  <svg viewBox="0 0 24 24" className="w-4 h-4 shrink-0 fill-none stroke-current" strokeWidth="2.5">
                    <circle cx="12" cy="12" r="10" />
                    <line x1="12" y1="8" x2="12" y2="12" />
                    <line x1="12" y1="16" x2="12.01" y2="16" />
                  </svg>
                  <span>{locationError}</span>
                </p>
              ) : (
                <p className="text-[13px] text-[#64748b] mt-1.5">
                  Select plant mess facility where meal is scheduled.
                </p>
              )}
            </div>

            {/* 2. Meal Type Selection (Top-Right) */}
            <div>
              <label className="block text-[14.5px] font-semibold text-[#1e293b] mb-2">
                Meal Type <span className="text-rose-500">*</span>
              </label>
              <div
                className={`grid grid-cols-3 gap-3 p-1 rounded-2xl transition-all ${
                  mealTypeError && !selectedMealType ? "border-2 border-rose-400 bg-rose-50/30" : ""
                }`}
              >
                {mealTypeOptions.map((type) => {
                  const isSelected = selectedMealType === type;
                  const status = mealValidation[type];
                  const isExpired = status?.isExpired;

                  return (
                    <div
                      key={type}
                      className="relative group w-full"
                    >
                      <button
                        type="button"
                        onClick={() => {
                          if (!isExpired) {
                            setSelectedMealType(type);
                            setTouched((prev) => ({ ...prev, mealType: true }));
                          }
                        }}
                        disabled={isSubmitting || isExpired}
                        style={
                          isExpired
                            ? {
                                opacity: 0.5,
                                backgroundColor: "#f1f5f9",
                                color: "#94a3b8",
                                borderColor: "#e2e8f0",
                                cursor: "not-allowed",
                              }
                            : undefined
                        }
                        className={`w-full min-h-[46px] h-[46px] rounded-xl flex items-center justify-center gap-2 text-[15px] font-semibold border-2 transition-all ${
                          isExpired
                            ? "cursor-not-allowed pointer-events-none"
                            : isSelected
                            ? "border-[#004B87] bg-blue-50/70 text-[#004B87] shadow-xs cursor-pointer"
                            : "border-slate-200 bg-white text-slate-700 hover:border-slate-300 hover:bg-slate-50 cursor-pointer"
                        }`}
                      >
                        <span className="text-lg">{type === "Breakfast" ? "🍳" : type === "Lunch" ? "🥗" : "🍲"}</span>
                        <span>{type}</span>
                      </button>

                      {/* Floating hover tooltip for expired meal button */}
                      {isExpired && status.message && (
                        <div
                          role="tooltip"
                          className="pointer-events-none absolute bottom-full left-1/2 -translate-x-1/2 mb-2 w-72 p-3 bg-slate-900/95 text-white text-[12px] leading-snug font-medium rounded-xl shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-150 z-30 text-center"
                        >
                          <div className="flex items-center gap-1.5 justify-center mb-1 text-amber-300 font-semibold text-[11.5px]">
                            <svg viewBox="0 0 24 24" className="w-4 h-4 shrink-0 fill-none stroke-current" strokeWidth="2.5">
                              <circle cx="12" cy="12" r="10" />
                              <line x1="12" y1="8" x2="12" y2="12" />
                              <line x1="12" y1="16" x2="12.01" y2="16" />
                            </svg>
                            <span>Cut-Off Passed</span>
                          </div>
                          <span>{status.message}</span>
                          <div className="absolute top-full left-1/2 -translate-x-1/2 border-4 border-transparent border-t-slate-900/95" />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
              {mealTypeError ? (
                <p className="text-[13px] text-rose-600 font-semibold mt-1.5 flex items-center gap-1.5">
                  <svg viewBox="0 0 24 24" className="w-4 h-4 shrink-0 fill-none stroke-current" strokeWidth="2.5">
                    <circle cx="12" cy="12" r="10" />
                    <line x1="12" y1="8" x2="12" y2="12" />
                    <line x1="12" y1="16" x2="12.01" y2="16" />
                  </svg>
                  <span>{mealTypeError}</span>
                </p>
              ) : cutOffDisplay ? (
                <p className="text-[13px] text-slate-500 mt-1.5">
                  Deadline: <span className="font-bold text-slate-700">{cutOffDisplay.date} at {cutOffDisplay.deadline} IST</span>
                </p>
              ) : (
                <p className="text-[13px] text-[#64748b] mt-1.5">
                  Select Breakfast, Lunch, or Dinner.
                </p>
              )}
            </div>

            {/* 3. Meal Date Picker (Bottom-Left) */}
            <div>
              <label className="block text-[14.5px] font-semibold text-[#1e293b] mb-2">
                Meal Date <span className="text-rose-500">*</span>
              </label>
              <input
                type="date"
                value={selectedDate}
                min={todayISOString}
                placeholder="dd-mm-yyyy"
                onChange={(e) => {
                  setSelectedDate(e.target.value);
                  setTouched((prev) => ({ ...prev, date: true }));
                }}
                onBlur={() => setTouched((prev) => ({ ...prev, date: true }))}
                disabled={isSubmitting}
                className={`w-full min-h-[46px] h-[46px] px-4 rounded-xl border bg-white text-slate-800 text-[15px] font-medium focus:outline-none focus:ring-2 transition ${
                  dateError
                    ? "border-rose-500 focus:ring-rose-500/20 bg-rose-50/20"
                    : "border-slate-300 focus:ring-[#004B87] focus:border-transparent"
                }`}
              />
              {dateError ? (
                <p className="text-[13px] text-rose-600 font-semibold mt-1.5 flex items-center gap-1.5">
                  <svg viewBox="0 0 24 24" className="w-4 h-4 shrink-0 fill-none stroke-current" strokeWidth="2.5">
                    <circle cx="12" cy="12" r="10" />
                    <line x1="12" y1="8" x2="12" y2="12" />
                    <line x1="12" y1="16" x2="12.01" y2="16" />
                  </svg>
                  <span>{dateError}</span>
                </p>
              ) : (
                <p className="text-[13px] text-[#64748b] mt-1.5">
                  Advance date for cancellation. Past dates are blocked.
                </p>
              )}
            </div>

            {/* 4. Requester (Bottom-Right, Read-Only) */}
            <div>
              <label className="block text-[14.5px] font-semibold text-[#1e293b] mb-2">
                Requester
              </label>
              <div className="min-h-[46px] h-[46px] px-4 rounded-xl border border-slate-200 bg-slate-50 flex items-center justify-between text-slate-700 text-[14.5px]">
                <span className="font-bold text-slate-900 truncate mr-2">{currentUserName}</span>
                <span className="text-slate-500 font-mono text-[13px] truncate">{currentUserEmail}</span>
              </div>
              <p className="text-[13px] text-[#64748b] mt-1.5">
                Automatically registered under signed-in user.
              </p>
            </div>
          </div>

          {/* Form Actions */}
          <div className="pt-5 mt-5 border-t border-slate-200 flex items-center justify-end gap-3.5">
            <button
              type="button"
              onClick={onCancel}
              disabled={isSubmitting}
              className="h-[46px] px-6 py-2.5 rounded-full border border-slate-300 text-slate-700 hover:bg-slate-100 text-[15px] font-semibold transition cursor-pointer whitespace-nowrap inline-flex items-center justify-center"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={isSubmitting || isCurrentSelectionExpired}
              className={`h-[46px] px-7 py-2.5 rounded-full text-white text-[15px] font-semibold shadow-sm transition-all inline-flex items-center justify-center gap-2.5 cursor-pointer whitespace-nowrap ${
                isSubmitting || isCurrentSelectionExpired
                  ? "bg-slate-300 text-slate-500 cursor-not-allowed"
                  : "bg-[#0066CC] hover:bg-[#0052A3] hover:shadow-md"
              }`}
            >
              {isSubmitting ? (
                <>
                  <svg className="animate-spin w-4 h-4 shrink-0 text-white" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" />
                  </svg>
                  <span>Submitting...</span>
                </>
              ) : (
                <>
                  <span>Submit Meal Opt-Out</span>
                  <svg viewBox="0 0 24 24" className="w-4 h-4 shrink-0 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                    <line x1="5" y1="12" x2="19" y2="12" />
                    <polyline points="12 5 19 12 12 19" />
                  </svg>
                </>
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
