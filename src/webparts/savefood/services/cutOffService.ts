import { MealType } from "../models/IMealOptOut";

export interface ICutOffValidationResult {
  isValid: boolean;
  cutOffDateTimeIST: Date;
  currentDateTimeIST: Date;
  errorMessage?: string;
}

/**
 * Submission Cut-off schedule specification (India Standard Time):
 * Employees can submit meal opt-outs until this deadline:
 * - Breakfast: 08:00 PM (20:00) on the calendar day before MealDate
 * - Lunch: 10:00 AM (10:00) on MealDate
 * - Dinner: 06:00 PM (18:00) on MealDate
 */
export const CUT_OFF_RULES: Record<MealType, { hour: number; minute: number; daysBefore: number; description: string }> = {
  Breakfast: {
    hour: 20,
    minute: 0,
    daysBefore: 1,
    description: "08:00 PM (Prev Day)",
  },
  Lunch: {
    hour: 10,
    minute: 0,
    daysBefore: 0,
    description: "10:00 AM (Same Day)",
  },
  Dinner: {
    hour: 18,
    minute: 0,
    daysBefore: 0,
    description: "06:00 PM (Same Day)",
  },
};

/**
 * Retraction & Status Lock schedule specification (India Standard Time):
 * Opt-out requests lock (status = "Locked (Passed)" and retraction disabled)
 * exactly 10 minutes before the creation deadline / email dispatch:
 * - Breakfast: 07:50 PM (19:50) on the calendar day before MealDate
 * - Lunch: 09:50 AM (09:50) on MealDate
 * - Dinner: 05:50 PM (17:50) on MealDate
 */
export const RETRACTION_LOCK_RULES: Record<MealType, { hour: number; minute: number; daysBefore: number; description: string }> = {
  Breakfast: {
    hour: 19,
    minute: 50,
    daysBefore: 1,
    description: "07:50 PM (Prev Day)",
  },
  Lunch: {
    hour: 9,
    minute: 50,
    daysBefore: 0,
    description: "09:50 AM (Same Day)",
  },
  Dinner: {
    hour: 17,
    minute: 50,
    daysBefore: 0,
    description: "05:50 PM (Same Day)",
  },
};

/**
 * Formats a given Date into "YYYY-MM-DD" string in India Standard Time (Asia/Kolkata).
 */
export function formatISTDateToISO(date: Date = new Date()): string {
  try {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Kolkata",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(date);
  } catch (e) {
    const { year, month, day } = parseISTDateParts(date);
    const m = month < 9 ? "0" + (month + 1) : String(month + 1);
    const d = day < 10 ? "0" + day : String(day);
    return `${year}-${m}-${d}`;
  }
}

/**
 * Deterministically parses any date representation (ISO string, SharePoint UTC string, Date object)
 * into its calendar Year, Month (0-11), and Day in IST (Asia/Kolkata).
 */
export function parseISTDateParts(input: Date | string): { year: number; month: number; day: number } {
  if (typeof input === "string") {
    const trimmed = input.trim();
    // Fast path for ISO "YYYY-MM-DD" or "YYYY-MM-DDT00:00:00" strings
    const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(trimmed);
    if (match && (trimmed.length === 10 || trimmed.indexOf("T00:00:00") !== -1)) {
      return {
        year: parseInt(match[1], 10),
        month: parseInt(match[2], 10) - 1,
        day: parseInt(match[3], 10),
      };
    }
  }

  const d = typeof input === "string" ? new Date(input) : input;
  let year = d.getFullYear();
  let month = d.getMonth();
  let day = d.getDate();

  try {
    const formatter = new Intl.DateTimeFormat("en-US", {
      timeZone: "Asia/Kolkata",
      year: "numeric",
      month: "numeric",
      day: "numeric",
    });
    if (typeof (formatter as any).formatToParts === "function") {
      const parts: Array<{ type: string; value: string }> = (formatter as any).formatToParts(d);
      for (let i = 0; i < parts.length; i++) {
        const p = parts[i];
        if (p.type === "year") year = parseInt(p.value, 10);
        if (p.type === "month") month = parseInt(p.value, 10) - 1;
        if (p.type === "day") day = parseInt(p.value, 10);
      }
    } else {
      const istStr = d.toLocaleString("en-US", { timeZone: "Asia/Kolkata" });
      const parsed = new Date(istStr);
      if (!isNaN(parsed.getTime())) {
        year = parsed.getFullYear();
        month = parsed.getMonth();
        day = parsed.getDate();
      }
    }
  } catch (e) {
    // Fallback to local properties
  }

  return { year, month, day };
}

/**
 * Returns current Date in India Standard Time (Asia/Kolkata, UTC+05:30)
 * Uses Intl.DateTimeFormat parts for cross-browser, timezone-independent accuracy.
 */
export function getCurrentISTDate(date: Date = new Date()): Date {
  let year = 1970;
  let month = 0;
  let day = 1;
  let hour = 0;
  let minute = 0;
  let second = 0;

  try {
    const formatter = new Intl.DateTimeFormat("en-US", {
      timeZone: "Asia/Kolkata",
      year: "numeric",
      month: "numeric",
      day: "numeric",
      hour: "numeric",
      minute: "numeric",
      second: "numeric",
      hour12: false,
    });
    if (typeof (formatter as any).formatToParts === "function") {
      const parts: Array<{ type: string; value: string }> = (formatter as any).formatToParts(date);
      for (let i = 0; i < parts.length; i++) {
        const p = parts[i];
        if (p.type === "year") year = parseInt(p.value, 10);
        if (p.type === "month") month = parseInt(p.value, 10) - 1;
        if (p.type === "day") day = parseInt(p.value, 10);
        if (p.type === "hour") {
          const h = parseInt(p.value, 10);
          hour = h === 24 ? 0 : h;
        }
        if (p.type === "minute") minute = parseInt(p.value, 10);
        if (p.type === "second") second = parseInt(p.value, 10);
      }
      return new Date(year, month, day, hour, minute, second);
    }
  } catch (e) {
    // Fallback below
  }

  const istString = date.toLocaleString("en-US", { timeZone: "Asia/Kolkata" });
  return new Date(istString);
}

/**
 * Computes exact cut-off Date object in IST for a given MealDate and MealType
 */
export function getCutOffDateTimeIST(mealDate: Date | string, mealType: MealType): Date {
  const rule = CUT_OFF_RULES[mealType];
  const { year, month, day } = parseISTDateParts(mealDate);
  const cutOff = new Date(year, month, day, rule.hour, rule.minute, 0, 0);

  if (rule.daysBefore > 0) {
    cutOff.setDate(cutOff.getDate() - rule.daysBefore);
  }
  return cutOff;
}

/**
 * Buffer in milliseconds before the cut-off deadline when new submissions close.
 * 30 seconds buffer prevents race conditions with top-of-the-hour email dispatch / scheduled tasks.
 */
export const SUBMISSION_CUTOFF_BUFFER_MS = 30 * 1000;

/**
 * Double-barrier validation for Meal Opt-Out cut-offs:
 * Checks whether selected date is in the past and whether current local IST time has passed the deadline
 * (stops 30 seconds before cut-off to avoid race conditions with report dispatches).
 */
export function validateMealCutOff(
  selectedMealDate: Date | string,
  mealType: MealType,
  currentDate: Date = new Date()
): ICutOffValidationResult {
  const currentIST = getCurrentISTDate(currentDate);
  const { year, month, day } = parseISTDateParts(selectedMealDate);

  // Normalize meal date to midnight IST
  const mealDateMidnightIST = new Date(year, month, day, 0, 0, 0, 0);

  const todayMidnightIST = new Date(
    currentIST.getFullYear(),
    currentIST.getMonth(),
    currentIST.getDate(),
    0,
    0,
    0,
    0
  );

  // 1. Guard against past calendar dates
  if (mealDateMidnightIST.getTime() < todayMidnightIST.getTime()) {
    return {
      isValid: false,
      cutOffDateTimeIST: todayMidnightIST,
      currentDateTimeIST: currentIST,
      errorMessage: "Meal Date cannot be in the past.",
    };
  }

  // 2. Compute cut-off deadline in IST and apply 30-second pre-cutoff guard
  const cutOff = getCutOffDateTimeIST(selectedMealDate, mealType);
  const effectiveDeadlineMs = cutOff.getTime() - SUBMISSION_CUTOFF_BUFFER_MS;

  // 3. Evaluate boundary condition: reject when current time >= cut-off minus 30 seconds
  if (currentIST.getTime() >= effectiveDeadlineMs) {
    const formattedCutOffTime = cutOff.toLocaleTimeString("en-IN", {
      hour: "2-digit",
      minute: "2-digit",
      hour12: true,
    });
    const formattedCutOffDate = cutOff.toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
    const formattedMealDate = mealDateMidnightIST.toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
    });

    return {
      isValid: false,
      cutOffDateTimeIST: cutOff,
      currentDateTimeIST: currentIST,
      errorMessage: `Cut-off passed. Submissions for ${mealType} on ${formattedMealDate} closed on ${formattedCutOffDate} at ${formattedCutOffTime} IST (submissions close 30s before cut-off).`,
    };
  }

  return {
    isValid: true,
    cutOffDateTimeIST: cutOff,
    currentDateTimeIST: currentIST,
  };
}

/**
 * Computes exact retraction & lock cut-off Date object in IST for a given MealDate and MealType.
 * Locks 10 minutes before the creation deadline / email dispatch.
 */
export function getRetractionCutOffDateTimeIST(mealDate: Date | string, mealType: MealType): Date {
  const rule = RETRACTION_LOCK_RULES[mealType];
  const { year, month, day } = parseISTDateParts(mealDate);
  const cutOff = new Date(year, month, day, rule.hour, rule.minute, 0, 0);

  if (rule.daysBefore > 0) {
    cutOff.setDate(cutOff.getDate() - rule.daysBefore);
  }
  return cutOff;
}

/**
 * Double-barrier validation for Meal Opt-Out Retraction & Status Lock:
 * Evaluates whether an opt-out is still modifiable / retractable.
 * Locks (returns isValid = false) 10 minutes prior to email dispatch.
 */
export function validateRetractionCutOff(
  selectedMealDate: Date | string,
  mealType: MealType,
  currentDate: Date = new Date()
): ICutOffValidationResult {
  const currentIST = getCurrentISTDate(currentDate);
  const { year, month, day } = parseISTDateParts(selectedMealDate);

  // Normalize meal date to midnight IST
  const mealDateMidnightIST = new Date(year, month, day, 0, 0, 0, 0);

  const todayMidnightIST = new Date(
    currentIST.getFullYear(),
    currentIST.getMonth(),
    currentIST.getDate(),
    0,
    0,
    0,
    0
  );

  // 1. Guard against past calendar dates
  if (mealDateMidnightIST.getTime() < todayMidnightIST.getTime()) {
    return {
      isValid: false,
      cutOffDateTimeIST: todayMidnightIST,
      currentDateTimeIST: currentIST,
      errorMessage: "Meal Date has already passed.",
    };
  }

  // 2. Compute retraction cut-off deadline in IST (10 mins before cut-off)
  const cutOff = getRetractionCutOffDateTimeIST(selectedMealDate, mealType);

  // 3. Evaluate boundary condition: reject retraction when current time >= cut-off
  if (currentIST.getTime() >= cutOff.getTime()) {
    const formattedCutOffTime = cutOff.toLocaleTimeString("en-IN", {
      hour: "2-digit",
      minute: "2-digit",
      hour12: true,
    });
    const formattedCutOffDate = cutOff.toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
    const formattedMealDate = mealDateMidnightIST.toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
    });

    return {
      isValid: false,
      cutOffDateTimeIST: cutOff,
      currentDateTimeIST: currentIST,
      errorMessage: `Retraction closed. Retraction for ${mealType} on ${formattedMealDate} closed on ${formattedCutOffDate} at ${formattedCutOffTime} IST (10 minutes before meal cut-off).`,
    };
  }

  return {
    isValid: true,
    cutOffDateTimeIST: cutOff,
    currentDateTimeIST: currentIST,
  };
}
