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

  // Date selection: Default to today in YYYY-MM-DD
  const todayIST = getCurrentISTDate();
  const todayISOString = todayIST.toISOString().split("T")[0];
  const [selectedDate, setSelectedDate] = React.useState<string>(todayISOString);

  // Status & Validation State
  const [isLoadingLocations, setIsLoadingLocations] = React.useState<boolean>(true);
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [validationError, setValidationError] = React.useState<string | null>(null);
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

  // 2. Validate cut-off whenever date or meal type changes
  React.useEffect(() => {
    if (!selectedDate) {
      setValidationError("Please select a valid meal date.");
      return;
    }

    if (!selectedMealType) {
      setValidationError(null);
      return;
    }

    const result = validateMealCutOff(selectedDate, selectedMealType);

    if (!result.isValid) {
      setValidationError(result.errorMessage || "Submission closed for this meal.");
    } else {
      setValidationError(null);
    }
  }, [selectedDate, selectedMealType]);

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
      : validationError
      ? validationError
      : null;

  // Handle Form Submission
  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    setHasAttemptedSubmit(true);
    setServerError(null);

    if (!selectedLocationId || !selectedDate || !selectedMealType || !!validationError) {
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

  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header */}
      <div className="mb-3.5 flex items-center justify-between w-full">
        <div>
          <BackButton onClick={onCancel} />
          <h2 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
            Create Meal Opt-Out Request
          </h2>
          <p className="text-xs sm:text-sm text-slate-500">
            Submit advance meal cancellations for upcoming plant shifts, official travel, or leave.
          </p>
        </div>
      </div>

      {/* Main Clean Form Card */}
      <div
        style={{ height: "auto", flex: "0 0 auto" }}
        className="w-full bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm p-6"
      >
        {/* Server Error Alert */}
        {serverError && (
          <div className="mb-4 p-3.5 rounded-xl bg-rose-50 border border-rose-200 flex items-center gap-2.5 text-rose-800 text-xs sm:text-sm font-medium">
            <svg viewBox="0 0 24 24" className="w-4 h-4 text-rose-600 shrink-0" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="12" cy="12" r="10" />
              <line x1="12" y1="8" x2="12" y2="12" />
              <line x1="12" y1="16" x2="12.01" y2="16" />
            </svg>
            <span>{serverError}</span>
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5 sm:gap-6">
            {/* 1. Location Selection */}
            <div>
              <label className="block text-xs sm:text-sm font-bold text-slate-800 mb-1.5">
                Canteen / Mess Location <span className="text-rose-500">*</span>
              </label>
              {isLoadingLocations ? (
                <div className="h-11 w-full rounded-xl bg-slate-100 animate-pulse" />
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
                  className={`w-full h-11 px-4 rounded-xl border bg-white text-slate-800 text-xs sm:text-sm font-semibold focus:outline-none focus:ring-2 transition ${
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
                <p className="text-[11.5px] text-rose-600 font-semibold mt-1 flex items-center gap-1">
                  <svg viewBox="0 0 24 24" className="w-3.5 h-3.5 shrink-0 fill-none stroke-current" strokeWidth="2.5">
                    <circle cx="12" cy="12" r="10" />
                    <line x1="12" y1="8" x2="12" y2="12" />
                    <line x1="12" y1="16" x2="12.01" y2="16" />
                  </svg>
                  <span>{locationError}</span>
                </p>
              ) : (
                <p className="text-[11px] text-slate-400 mt-1">
                  Select plant mess facility where meal is scheduled.
                </p>
              )}
            </div>

            {/* 2. Meal Date Picker */}
            <div>
              <label className="block text-xs sm:text-sm font-bold text-slate-800 mb-1.5">
                Meal Date <span className="text-rose-500">*</span>
              </label>
              <input
                type="date"
                value={selectedDate}
                min={todayISOString}
                onChange={(e) => {
                  setSelectedDate(e.target.value);
                  setTouched((prev) => ({ ...prev, date: true }));
                }}
                onBlur={() => setTouched((prev) => ({ ...prev, date: true }))}
                disabled={isSubmitting}
                className={`w-full h-11 px-4 rounded-xl border bg-white text-slate-800 text-xs sm:text-sm font-semibold focus:outline-none focus:ring-2 transition ${
                  dateError
                    ? "border-rose-500 focus:ring-rose-500/20 bg-rose-50/20"
                    : "border-slate-300 focus:ring-[#004B87] focus:border-transparent"
                }`}
              />
              {dateError ? (
                <p className="text-[11.5px] text-rose-600 font-semibold mt-1 flex items-center gap-1">
                  <svg viewBox="0 0 24 24" className="w-3.5 h-3.5 shrink-0 fill-none stroke-current" strokeWidth="2.5">
                    <circle cx="12" cy="12" r="10" />
                    <line x1="12" y1="8" x2="12" y2="12" />
                    <line x1="12" y1="16" x2="12.01" y2="16" />
                  </svg>
                  <span>{dateError}</span>
                </p>
              ) : (
                <p className="text-[11px] text-slate-400 mt-1">
                  Advance date for cancellation. Past dates are blocked.
                </p>
              )}
            </div>

            {/* 3. Meal Type Selection (Radio Pills with No Default & Inline Validation) */}
            <div>
              <label className="block text-xs sm:text-sm font-bold text-slate-800 mb-1.5">
                Meal Type <span className="text-rose-500">*</span>
              </label>
              <div
                className={`grid grid-cols-3 gap-3 p-1 rounded-2xl transition-all ${
                  mealTypeError && !selectedMealType ? "border-2 border-rose-400 bg-rose-50/30" : ""
                }`}
              >
                {(["Breakfast", "Lunch", "Dinner"] as MealType[]).map((type) => {
                  const isSelected = selectedMealType === type;
                  return (
                    <button
                      key={type}
                      type="button"
                      onClick={() => {
                        setSelectedMealType(type);
                        setTouched((prev) => ({ ...prev, mealType: true }));
                      }}
                      disabled={isSubmitting}
                      className={`h-11 rounded-xl flex items-center justify-center gap-2 text-xs sm:text-sm font-bold border-2 transition-all cursor-pointer ${
                        isSelected
                          ? "border-[#004B87] bg-blue-50/70 text-[#004B87] shadow-xs"
                          : "border-slate-200 bg-white text-slate-600 hover:border-slate-300 hover:bg-slate-50"
                      }`}
                    >
                      <span className="text-base">{type === "Breakfast" ? "🍳" : type === "Lunch" ? "🥗" : "🍲"}</span>
                      <span>{type}</span>
                    </button>
                  );
                })}
              </div>
              {mealTypeError ? (
                <p className="text-[11.5px] text-rose-600 font-semibold mt-1.5 flex items-center gap-1">
                  <svg viewBox="0 0 24 24" className="w-3.5 h-3.5 shrink-0 fill-none stroke-current" strokeWidth="2.5">
                    <circle cx="12" cy="12" r="10" />
                    <line x1="12" y1="8" x2="12" y2="12" />
                    <line x1="12" y1="16" x2="12.01" y2="16" />
                  </svg>
                  <span>{mealTypeError}</span>
                </p>
              ) : cutOffDisplay ? (
                <p className="text-[11px] text-slate-500 mt-1">
                  Deadline: <span className="font-bold text-slate-700">{cutOffDisplay.date} at {cutOffDisplay.deadline} IST</span>
                </p>
              ) : (
                <p className="text-[11px] text-slate-400 mt-1">
                  Select Breakfast, Lunch, or Dinner.
                </p>
              )}
            </div>

            {/* 4. Requester (Read-Only) */}
            <div>
              <label className="block text-xs sm:text-sm font-bold text-slate-800 mb-1.5">
                Requester
              </label>
              <div className="h-11 px-4 rounded-xl border border-slate-200 bg-slate-50 flex items-center justify-between text-slate-700 text-xs sm:text-sm">
                <span className="font-bold text-slate-900 truncate mr-2">{currentUserName}</span>
                <span className="text-slate-500 font-mono text-[11px] truncate">{currentUserEmail}</span>
              </div>
              <p className="text-[11px] text-slate-400 mt-1">
                Automatically registered under signed-in user.
              </p>
            </div>
          </div>

          {/* Form Actions */}
          <div className="pt-4 mt-4 border-t border-slate-200 flex items-center justify-end gap-3">
            <button
              type="button"
              onClick={onCancel}
              disabled={isSubmitting}
              className="px-6 py-2.5 rounded-full border border-slate-300 text-slate-600 hover:bg-slate-100 text-xs sm:text-sm font-bold transition cursor-pointer"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={isSubmitting || (!!validationError && !!selectedMealType)}
              className={`px-7 py-2.5 rounded-full text-white text-xs sm:text-sm font-bold shadow-sm transition-all flex items-center gap-2 cursor-pointer ${
                isSubmitting || (!!validationError && !!selectedMealType)
                  ? "bg-slate-300 text-slate-500 cursor-not-allowed"
                  : "bg-[#0066CC] hover:bg-[#0052A3] hover:shadow-md"
              }`}
            >
              {isSubmitting ? (
                <>
                  <svg className="animate-spin w-4 h-4 text-white" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" />
                  </svg>
                  <span>Submitting...</span>
                </>
              ) : (
                <>
                  <span>Submit Meal Opt-Out</span>
                  <svg viewBox="0 0 24 24" className="w-4 h-4 fill-none stroke-current" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
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
