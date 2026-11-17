import { MealType } from "../models/IMealOptOut";

export interface ICutOffValidationResult {
  isValid: boolean;
  cutOffDateTimeIST: Date;
  currentDateTimeIST: Date;
  errorMessage?: string;
}

/**
 * Cut-off schedule specification:
 * - Breakfast: 08:00 PM (20:00) on the calendar day before MealDate
 * - Lunch: 10:00 AM (10:00) on MealDate
 * - Dinner: 06:00 PM (18:00) on MealDate
 */
export const CUT_OFF_RULES: Record<MealType, { hour: number; minute: number; daysBefore: number; description: string }> = {
  Breakfast: {
    hour: 20,
    minute: 0,
    daysBefore: 1,
    description: "08:00 PM (Prior Day)",
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
 * Returns current Date in India Standard Time (Asia/Kolkata, UTC+05:30)
 */
export function getCurrentISTDate(date: Date = new Date()): Date {
  const istString = date.toLocaleString("en-US", { timeZone: "Asia/Kolkata" });
  return new Date(istString);
}

/**
 * Computes exact cut-off Date object in IST for a given MealDate and MealType
 */
export function getCutOffDateTimeIST(mealDate: Date, mealType: MealType): Date {
  const rule = CUT_OFF_RULES[mealType];
  const cutOff = new Date(mealDate.getFullYear(), mealDate.getMonth(), mealDate.getDate());

  if (rule.daysBefore > 0) {
    cutOff.setDate(cutOff.getDate() - rule.daysBefore);
  }
  cutOff.setHours(rule.hour, rule.minute, 0, 0);
  return cutOff;
}

/**
 * Double-barrier validation for Meal Opt-Out cut-offs:
 * Checks whether selected date is in the past and whether current local IST time has passed the deadline.
 */
export function validateMealCutOff(
  selectedMealDate: Date,
  mealType: MealType,
  currentDate: Date = new Date()
): ICutOffValidationResult {
  const currentIST = getCurrentISTDate(currentDate);

  // Normalize selected meal date to midnight IST
  const mealDateIST = new Date(
    selectedMealDate.getFullYear(),
    selectedMealDate.getMonth(),
    selectedMealDate.getDate(),
    0,
    0,
    0,
    0
  );

  const todayMidnightIST = new Date(
    currentIST.getFullYear(),
    currentIST.getMonth(),
    currentIST.getDate(),
    0,
    0,
    0,
    0
  );

  // 1. Guard against past dates
  if (mealDateIST.getTime() < todayMidnightIST.getTime()) {
    return {
      isValid: false,
      cutOffDateTimeIST: todayMidnightIST,
      currentDateTimeIST: currentIST,
      errorMessage: "Meal Date cannot be in the past.",
    };
  }

  // 2. Compute cut-off deadline in IST
  const cutOff = getCutOffDateTimeIST(mealDateIST, mealType);

  // 3. Evaluate boundary condition: reject when current time >= cut-off
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
    const formattedMealDate = mealDateIST.toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
    });

    return {
      isValid: false,
      cutOffDateTimeIST: cutOff,
      currentDateTimeIST: currentIST,
      errorMessage: `Submission closed. Cut-off for ${mealType} on ${formattedMealDate} was ${formattedCutOffDate} at ${formattedCutOffTime} IST.`,
    };
  }

  return {
    isValid: true,
    cutOffDateTimeIST: cutOff,
    currentDateTimeIST: currentIST,
  };
}
