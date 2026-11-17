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
  onSuccess: () => void;
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
  const [selectedMealType, setSelectedMealType] = React.useState<MealType>("Lunch");

  // Date selection: Default to today in YYYY-MM-DD
  const todayIST = getCurrentISTDate();
  const todayISOString = todayIST.toISOString().split("T")[0];
  const [selectedDate, setSelectedDate] = React.useState<string>(todayISOString);

  // Status & Validation State
  const [isLoadingLocations, setIsLoadingLocations] = React.useState<boolean>(true);
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [validationError, setValidationError] = React.useState<string | null>(null);
  const [serverError, setServerError] = React.useState<string | null>(null);
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);

  // 1. Fetch active locations on load
  React.useEffect(() => {
    let isMounted = true;
    MealRequestService.getActiveLocations(spHttpClient, siteUrl)
      .then((items) => {
        if (isMounted) {
          setLocations(items);
          if (items.length > 0) {
            setSelectedLocationId(items[0].Id);
          }
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

    const dateObj = new Date(selectedDate);
    const result = validateMealCutOff(dateObj, selectedMealType);

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
    if (!selectedDate) return null;
    const parts = selectedDate.split("-");
    const dateObj = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    const cutOff = getCutOffDateTimeIST(dateObj, selectedMealType);
    return {
      deadline: cutOff.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit", hour12: true }),
      date: cutOff.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }),
      rule: CUT_OFF_RULES[selectedMealType].description,
    };
  }, [selectedDate, selectedMealType]);

  // Handle Form Submission
  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    setServerError(null);

    if (!selectedLocationId) {
      setServerError("Please select a dining mess location.");
      return;
    }

    if (!selectedDate) {
      setServerError("Please select a meal date.");
      return;
    }

    if (validationError) {
      setServerError(validationError);
      return;
    }

    setIsSubmitting(true);

    const locationName = selectedLocationObj ? selectedLocationObj.Title : "Mess";
    const titleSummary = `${locationName} - ${selectedMealType} - ${selectedDate}`;

    const result = await MealRequestService.createMealOptOut(spHttpClient, siteUrl, {
      Title: titleSummary,
      LocationId: Number(selectedLocationId),
      MealType: selectedMealType,
      MealDate: `${selectedDate}T00:00:00Z`,
      RequesterId: currentUserId,
    });

    setIsSubmitting(false);

    if (result.success) {
      setSuccessMessage(`Opt-out request submitted successfully for ${selectedMealType} on ${selectedDate}.`);
      setTimeout(() => {
        onSuccess();
      }, 1500);
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
        {/* Success Alert */}
        {successMessage && (
          <div className="mb-4 p-3.5 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center gap-2.5 text-emerald-800 text-xs sm:text-sm font-medium">
            <svg viewBox="0 0 24 24" className="w-4 h-4 text-emerald-600 shrink-0" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
              <polyline points="22 4 12 14.01 9 11.01" />
            </svg>
            <span>{successMessage}</span>
          </div>
        )}

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

        {/* Client Cut-off Validation Warning */}
        {validationError && (
          <div className="mb-4 p-3.5 rounded-xl bg-amber-50 border border-amber-200 flex items-center gap-2.5 text-amber-900 text-xs font-medium">
            <svg viewBox="0 0 24 24" className="w-4 h-4 text-amber-600 shrink-0" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
              <line x1="12" y1="9" x2="12" y2="13" />
              <line x1="12" y1="17" x2="12.01" y2="17" />
            </svg>
            <div>
              <span className="font-bold text-amber-950">Cut-Off Window Closed: </span>
              <span className="text-amber-800">{validationError}</span>
            </div>
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
                  onChange={(e) => setSelectedLocationId(Number(e.target.value))}
                  disabled={isSubmitting}
                  className="w-full h-11 px-4 rounded-xl border border-slate-300 bg-white text-slate-800 text-xs sm:text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-[#004B87] focus:border-transparent transition"
                >
                  {locations.length === 0 ? (
                    <option value="">No active locations found</option>
                  ) : (
                    locations.map((loc) => (
                      <option key={loc.Id} value={loc.Id}>
                        {loc.Title}
                      </option>
                    ))
                  )}
                </select>
              )}
              <p className="text-[11px] text-slate-400 mt-1">
                Select plant mess facility where meal is scheduled.
              </p>
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
                onChange={(e) => setSelectedDate(e.target.value)}
                disabled={isSubmitting}
                className="w-full h-11 px-4 rounded-xl border border-slate-300 bg-white text-slate-800 text-xs sm:text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-[#004B87] focus:border-transparent transition"
              />
              <p className="text-[11px] text-slate-400 mt-1">
                Advance date for cancellation. Past dates are blocked.
              </p>
            </div>

            {/* 3. Meal Type Selection (Radio Pills) */}
            <div>
              <label className="block text-xs sm:text-sm font-bold text-slate-800 mb-1.5">
                Meal Type <span className="text-rose-500">*</span>
              </label>
              <div className="grid grid-cols-3 gap-3">
                {(["Breakfast", "Lunch", "Dinner"] as MealType[]).map((type) => {
                  const isSelected = selectedMealType === type;
                  return (
                    <button
                      key={type}
                      type="button"
                      onClick={() => setSelectedMealType(type)}
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
              {cutOffDisplay && (
                <p className="text-[11px] text-slate-500 mt-1">
                  Deadline: <span className="font-bold text-slate-700">{cutOffDisplay.date} at {cutOffDisplay.deadline} IST</span>
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
              disabled={isSubmitting || !!validationError}
              className={`px-7 py-2.5 rounded-full text-white text-xs sm:text-sm font-bold shadow-sm transition-all flex items-center gap-2 cursor-pointer ${
                isSubmitting || !!validationError
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
