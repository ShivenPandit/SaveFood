import * as XLSX from "xlsx";
import { IMealOptOutRequestItem } from "../models/IMealOptOut";

/**
 * Utility service to export meal opt-out records to a formatted Excel (.xlsx) spreadsheet
 *
 * @param items Array of opt-out items to export (full list or filtered subset)
 * @param isFiltered Whether active filters are applied in the UI
 */
export const exportMealOptOutsToExcel = (
  items: IMealOptOutRequestItem[],
  isFiltered: boolean = false
): void => {
  if (!items || items.length === 0) {
    return;
  }

  // 1. Map data rows to clear, human-readable column headers
  const rows = items.map((item, index) => {
    const mealDateObj = item.MealDate ? new Date(item.MealDate) : null;
    const formattedDate = mealDateObj
      ? mealDateObj.toLocaleDateString("en-IN", {
          day: "2-digit",
          month: "short",
          year: "numeric",
        })
      : "—";
    const dayOfWeek = mealDateObj
      ? mealDateObj.toLocaleDateString("en-IN", { weekday: "long" })
      : "—";

    const createdObj = item.Created ? new Date(item.Created) : null;
    const formattedCreated = createdObj
      ? createdObj.toLocaleDateString("en-IN", {
          day: "2-digit",
          month: "short",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        })
      : "—";

    return {
      "S.No": index + 1,
      "Meal Date": formattedDate,
      "Day": dayOfWeek,
      "Meal Type": item.MealType || "—",
      "Canteen Location": item.Location?.Title || "—",
      "Employee Name": item.Requester?.Title || "—",
      "Employee Email": item.Requester?.EMail || "—",
      "Submitted On": formattedCreated,
    };
  });

  // 2. Generate worksheet
  const worksheet = XLSX.utils.json_to_sheet(rows);

  // 3. Set custom column widths (wch) for pristine layout in Excel
  worksheet["!cols"] = [
    { wch: 8 },  // S.No
    { wch: 16 }, // Meal Date
    { wch: 14 }, // Day
    { wch: 14 }, // Meal Type
    { wch: 24 }, // Canteen Location
    { wch: 28 }, // Employee Name
    { wch: 34 }, // Employee Email
    { wch: 24 }, // Submitted On
  ];

  // 4. Create workbook and append sheet
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Meal Opt-Outs");

  // 5. Construct timestamped file name
  const todayStr = new Date().toISOString().split("T")[0];
  const filename = isFiltered
    ? `SaveFood_OptOuts_Filtered_${todayStr}.xlsx`
    : `SaveFood_OptOuts_All_${todayStr}.xlsx`;

  // 6. Generate binary array buffer and trigger download via browser Blob
  const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([excelBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  window.URL.revokeObjectURL(url);
};
