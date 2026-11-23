import * as React from "react";
import { MealType } from "../../models/IMealOptOut";
import { ILocationMasterItem } from "../../models/ILocation";
import { formatISTDateToISO } from "../../services/cutOffService";

export interface IAdvancedFilterPopoverProps {
  isOpen: boolean;
  onClose: () => void;
  // Date Range
  startDate: string | null;
  endDate: string | null;
  onStartDateChange: (val: string | null) => void;
  onEndDateChange: (val: string | null) => void;
  // Meal Type Multi-Select
  selectedMealTypes: MealType[];
  onMealTypesChange: (types: MealType[]) => void;
  // Canteen Location Multi-Select
  locations: ILocationMasterItem[];
  selectedLocationIds: number[];
  onLocationsChange: (ids: number[]) => void;
  // Optional legacy props kept for backward-compatibility without rendering
  showRequesterFilter?: boolean;
  selectedRequesters?: any[];
  onRequestersChange?: (users: any[]) => void;
  context?: any;
  spHttpClient?: any;
  siteUrl?: string;
  // Actions
  onResetAll: () => void;
  activeFilterCount: number;
}

export const AdvancedFilterPopover: React.FC<IAdvancedFilterPopoverProps> = ({
  isOpen,
  onClose,
  startDate,
  endDate,
  onStartDateChange,
  onEndDateChange,
  selectedMealTypes,
  onMealTypesChange,
  locations,
  selectedLocationIds,
  onLocationsChange,
  onResetAll,
  activeFilterCount,
}) => {
  const popoverRef = React.useRef<HTMLDivElement>(null);

  // Close on outside click
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (popoverRef.current && !popoverRef.current.contains(event.target as Node)) {
        onClose();
      }
    };
    if (isOpen) {
      document.addEventListener("mousedown", handleClickOutside);
    }
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [isOpen, onClose]);

  // Date Presets
  const todayISOString = React.useMemo(() => formatISTDateToISO(new Date()), []);
  const tomorrowISOString = React.useMemo(() => {
    const d = new Date();
    d.setDate(d.getDate() + 1);
    return formatISTDateToISO(d);
  }, []);

  const isTodayActive = startDate === todayISOString && endDate === todayISOString;
  const isTomorrowActive = startDate === tomorrowISOString && endDate === tomorrowISOString;
  const isAllDatesActive = startDate === null && endDate === null;
  const isCustomRange =
    !isTodayActive &&
    !isTomorrowActive &&
    !isAllDatesActive &&
    (Boolean(startDate) || Boolean(endDate));

  // Date Range Change Handlers (guarantees To >= From)
  const handleStartDateChange = (val: string | null): void => {
    onStartDateChange(val);
    if (val && endDate && val > endDate) {
      onEndDateChange(val);
    }
  };

  const handleEndDateChange = (val: string | null): void => {
    if (val && startDate && val < startDate) {
      onEndDateChange(startDate);
    } else {
      onEndDateChange(val);
    }
  };

  const isAllMealsSelected = selectedMealTypes.length === 0 || selectedMealTypes.length === 3;
  const isBreakfastSelected = !isAllMealsSelected && selectedMealTypes.indexOf("Breakfast") !== -1;
  const isLunchSelected = !isAllMealsSelected && selectedMealTypes.indexOf("Lunch") !== -1;
  const isDinnerSelected = !isAllMealsSelected && selectedMealTypes.indexOf("Dinner") !== -1;

  const handleToggleMealType = (type: MealType): void => {
    if (selectedMealTypes.length === 0) {
      onMealTypesChange([type]);
      return;
    }
    if (selectedMealTypes.indexOf(type) !== -1) {
      const next = selectedMealTypes.filter((t) => t !== type);
      onMealTypesChange(next);
    } else {
      const next = [...selectedMealTypes, type];
      if (next.length === 3) {
        onMealTypesChange([]);
      } else {
        onMealTypesChange(next);
      }
    }
  };

  const handleSelectAllMeals = (): void => {
    onMealTypesChange([]);
  };

  const isAllLocationsSelected =
    selectedLocationIds.length === 0 || selectedLocationIds.length === locations.length;

  const handleToggleLocation = (id: number): void => {
    if (isAllLocationsSelected) {
      onLocationsChange([id]);
      return;
    }
    if (selectedLocationIds.indexOf(id) !== -1) {
      const next = selectedLocationIds.filter((item) => item !== id);
      onLocationsChange(next);
    } else {
      const next = [...selectedLocationIds, id];
      if (next.length === locations.length) {
        onLocationsChange([]);
      } else {
        onLocationsChange(next);
      }
    }
  };

  const handleSelectAllLocations = (): void => {
    onLocationsChange([]);
  };

  if (!isOpen) return null;

  return (
    <>
      {/* Mobile Backdrop */}
      <div
        className="fixed inset-0 bg-slate-900/40 z-40 lg:hidden backdrop-blur-xs"
        onClick={onClose}
      />

      <div
        ref={popoverRef}
        className="fixed inset-x-4 top-20 max-w-lg mx-auto lg:mx-0 lg:max-w-none lg:inset-x-auto lg:absolute lg:top-full lg:right-0 mt-0 lg:mt-2 w-auto lg:w-[485px] max-h-[90vh] sm:max-h-none overflow-y-auto sm:overflow-visible bg-white rounded-2xl border-2 border-slate-200/90 shadow-2xl p-4 sm:p-5 z-50 animate-in fade-in zoom-in-95 duration-150"
        style={{ filter: "drop-shadow(0 20px 25px -5px rgba(0, 0, 0, 0.2))" }}
      >
        {/* Popover Header */}
        <div className="flex items-center justify-between pb-3 mb-3 border-b border-slate-100">
          <div className="flex items-center gap-2 text-[#0f172a]">
            <div className="w-7 h-7 rounded-lg bg-blue-50 text-[#004B87] flex items-center justify-center shrink-0">
              <svg
                width="16"
                height="16"
                viewBox="0 0 24 24"
                className="fill-none stroke-current"
                style={{ width: 16, height: 16, minWidth: 16, maxWidth: 16 }}
                strokeWidth="2.2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
              </svg>
            </div>
            <h3 className="text-[15.5px] font-bold tracking-tight">Advanced Filter Criteria</h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="w-7 h-7 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 flex items-center justify-center transition cursor-pointer"
            title="Close filters"
            aria-label="Close filters"
          >
            <svg
              width="16"
              height="16"
              viewBox="0 0 24 24"
              className="fill-none stroke-current"
              style={{ width: 16, height: 16, minWidth: 16, maxWidth: 16 }}
              strokeWidth="2.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>

        {/* Filter Controls Form - Compact & Zero-Scroll */}
        <div className="space-y-3.5">
          {/* 1. Meal Date Range Filter & Quick Presets */}
          <div>
            <div className="flex items-center justify-between mb-1">
              <label className="block text-[11.5px] font-bold text-slate-600 uppercase tracking-wider">
                MEAL DATE RANGE
              </label>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => {
                    onStartDateChange(todayISOString);
                    onEndDateChange(todayISOString);
                  }}
                  className={`px-2.5 py-0.5 rounded-md text-[11.5px] font-semibold border transition cursor-pointer ${
                    isTodayActive
                      ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs"
                      : "bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100"
                  }`}
                >
                  Today
                </button>
                <button
                  type="button"
                  onClick={() => {
                    onStartDateChange(tomorrowISOString);
                    onEndDateChange(tomorrowISOString);
                  }}
                  className={`px-2.5 py-0.5 rounded-md text-[11.5px] font-semibold border transition cursor-pointer ${
                    isTomorrowActive
                      ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs"
                      : "bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100"
                  }`}
                >
                  Tomorrow
                </button>
                <button
                  type="button"
                  onClick={() => {
                    onStartDateChange(null);
                    onEndDateChange(null);
                  }}
                  className={`px-2.5 py-0.5 rounded-md text-[11.5px] font-semibold border transition cursor-pointer ${
                    isAllDatesActive
                      ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs"
                      : "bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100"
                  }`}
                >
                  All Dates
                </button>
              </div>
            </div>

            <div className="flex items-center gap-2">
              {/* From Date */}
              <div className="flex items-center bg-white border border-slate-300 rounded-xl overflow-hidden focus-within:ring-2 focus-within:ring-[#004B87] focus-within:border-[#004B87] flex-1 min-h-[38px]">
                <span className="px-2.5 text-[12px] font-bold text-slate-500 bg-slate-100 border-r border-slate-200 py-1.5 select-none shrink-0">
                  From
                </span>
                <input
                  type="date"
                  value={startDate || ""}
                  max={endDate || undefined}
                  onChange={(e) => handleStartDateChange(e.target.value || null)}
                  aria-label="From Date"
                  className="w-full h-[38px] px-2 text-[13px] font-medium text-slate-800 bg-white focus:outline-none"
                />
              </div>

              <span className="text-slate-400 font-bold text-xs shrink-0 select-none">to</span>

              {/* To Date */}
              <div className="flex items-center bg-white border border-slate-300 rounded-xl overflow-hidden focus-within:ring-2 focus-within:ring-[#004B87] focus-within:border-[#004B87] flex-1 min-h-[38px]">
                <span className="px-2.5 text-[12px] font-bold text-slate-500 bg-slate-100 border-r border-slate-200 py-1.5 select-none shrink-0">
                  To
                </span>
                <input
                  type="date"
                  value={endDate || ""}
                  min={startDate || undefined}
                  onChange={(e) => handleEndDateChange(e.target.value || null)}
                  aria-label="To Date"
                  className="w-full h-[38px] px-2 text-[13px] font-medium text-slate-800 bg-white focus:outline-none"
                />
              </div>

              {/* Clear Range Button */}
              {isCustomRange && (
                <button
                  type="button"
                  onClick={() => {
                    onStartDateChange(null);
                    onEndDateChange(null);
                  }}
                  title="Clear Date Filter"
                  className="h-[38px] px-2.5 rounded-xl border border-slate-300 bg-slate-100 hover:bg-slate-200 text-slate-600 transition flex items-center justify-center shrink-0 cursor-pointer"
                >
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    className="fill-none stroke-current"
                    style={{ width: 14, height: 14, minWidth: 14, maxWidth: 14 }}
                    strokeWidth="2.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <line x1="18" y1="6" x2="6" y2="18" />
                    <line x1="6" y1="6" x2="18" y2="18" />
                  </svg>
                </button>
              )}
            </div>
          </div>

          {/* 2. Meal Type Multi-Select (Direct Toggle Pills) */}
          <div>
            <div className="flex items-center justify-between mb-1.5">
              <label className="block text-[11.5px] font-bold text-slate-600 uppercase tracking-wider">
                MEAL TYPE
              </label>
              {!isAllMealsSelected && (
                <button
                  type="button"
                  onClick={handleSelectAllMeals}
                  className="text-[11px] font-bold text-[#004B87] hover:underline cursor-pointer"
                >
                  Select All
                </button>
              )}
            </div>
            <div
              className="grid grid-cols-4 gap-1.5"
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(4, minmax(0, 1fr))",
                gap: "6px",
              }}
            >
              {/* All Meals Pill */}
              <button
                type="button"
                onClick={handleSelectAllMeals}
                className={`py-1.5 px-1 text-center rounded-xl text-[12px] font-bold border transition cursor-pointer select-none ${
                  isAllMealsSelected
                    ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs"
                    : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
                }`}
              >
                All Meals
              </button>

              {/* Breakfast Pill */}
              <button
                type="button"
                onClick={() => handleToggleMealType("Breakfast")}
                className={`py-1.5 px-1 text-center rounded-xl text-[12px] font-bold border transition cursor-pointer select-none flex items-center justify-center gap-1 ${
                  isBreakfastSelected
                    ? "bg-amber-500 text-white border-amber-600 shadow-2xs ring-1 ring-amber-400"
                    : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-amber-50/50 hover:text-amber-800"
                }`}
              >
                <span>🍳</span>
                <span>Breakfast</span>
              </button>

              {/* Lunch Pill */}
              <button
                type="button"
                onClick={() => handleToggleMealType("Lunch")}
                className={`py-1.5 px-1 text-center rounded-xl text-[12px] font-bold border transition cursor-pointer select-none flex items-center justify-center gap-1 ${
                  isLunchSelected
                    ? "bg-blue-600 text-white border-blue-700 shadow-2xs ring-1 ring-blue-400"
                    : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-blue-50/50 hover:text-blue-800"
                }`}
              >
                <span>🥗</span>
                <span>Lunch</span>
              </button>

              {/* Dinner Pill */}
              <button
                type="button"
                onClick={() => handleToggleMealType("Dinner")}
                className={`py-1.5 px-1 text-center rounded-xl text-[12px] font-bold border transition cursor-pointer select-none flex items-center justify-center gap-1 ${
                  isDinnerSelected
                    ? "bg-purple-600 text-white border-purple-700 shadow-2xs ring-1 ring-purple-400"
                    : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-purple-50/50 hover:text-purple-800"
                }`}
              >
                <span>🍲</span>
                <span>Dinner</span>
              </button>
            </div>
          </div>

          {/* 3. Canteen Location Multi-Select (Direct Toggle Chips - Zero Scrolling!) */}
          <div>
            <div className="flex items-center justify-between mb-1.5">
              <label className="block text-[11.5px] font-bold text-slate-600 uppercase tracking-wider">
                CANTEEN LOCATION
              </label>
              {!isAllLocationsSelected && (
                <button
                  type="button"
                  onClick={handleSelectAllLocations}
                  className="text-[11px] font-bold text-[#004B87] hover:underline cursor-pointer"
                >
                  Select All
                </button>
              )}
            </div>

            <div
              className="flex flex-wrap gap-1.5"
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: "6px",
              }}
            >
              {/* All Locations Chip */}
              <button
                type="button"
                onClick={handleSelectAllLocations}
                className={`py-1.5 px-3 rounded-xl text-[12px] font-bold border transition cursor-pointer select-none ${
                  isAllLocationsSelected
                    ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs"
                    : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
                }`}
              >
                All Mess Locations
              </button>

              {/* Dynamic Location Chips */}
              {locations.map((loc) => {
                const isSelected = !isAllLocationsSelected && selectedLocationIds.indexOf(loc.Id) !== -1;
                return (
                  <button
                    key={loc.Id}
                    type="button"
                    onClick={() => handleToggleLocation(loc.Id)}
                    className={`py-1.5 px-3 rounded-xl text-[12px] font-bold border transition cursor-pointer select-none flex items-center gap-1.5 ${
                      isSelected
                        ? "bg-[#004B87] text-white border-[#004B87] shadow-2xs ring-1 ring-[#004B87]/30"
                        : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
                    }`}
                  >
                    {isSelected && <span className="font-bold text-[11px] leading-none">✓</span>}
                    <span>{loc.Title}</span>
                  </button>
                );
              })}
            </div>
          </div>
        </div>

        {/* Popover Footer - Always Visible Without Scrolling */}
        <div className="flex items-center justify-between pt-3 mt-3.5 border-t border-slate-100 gap-2">
          <button
            type="button"
            onClick={onResetAll}
            disabled={activeFilterCount === 0}
            className="px-3.5 py-1.5 rounded-xl text-[12.5px] font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition cursor-pointer"
          >
            Reset All
          </button>
          <button
            type="button"
            onClick={onClose}
            className="px-5 py-2 rounded-xl bg-[#004B87] hover:bg-[#003865] text-white text-[13px] font-bold shadow-xs hover:shadow transition cursor-pointer"
          >
            Apply &amp; Close
          </button>
        </div>
      </div>
    </>
  );
};
