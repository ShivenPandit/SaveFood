import * as React from "react";

export interface IDataTablePaginationProps {
  currentPage: number;
  totalPages: number;
  totalItems: number;
  startIndex: number;
  endIndex: number;
  onPageChange: (page: number) => void;
  className?: string;
}

// Lucide-style crisp SVGs
const ChevronsLeftIcon: React.FC<{ className?: string }> = ({ className = "w-4 h-4" }) => (
  <svg
    viewBox="0 0 24 24"
    className={className}
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <polyline points="11 17 6 12 11 7" />
    <polyline points="18 17 13 12 18 7" />
  </svg>
);

const ChevronLeftIcon: React.FC<{ className?: string }> = ({ className = "w-4 h-4" }) => (
  <svg
    viewBox="0 0 24 24"
    className={className}
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <polyline points="15 18 9 12 15 6" />
  </svg>
);

const ChevronRightIcon: React.FC<{ className?: string }> = ({ className = "w-4 h-4" }) => (
  <svg
    viewBox="0 0 24 24"
    className={className}
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <polyline points="9 18 15 12 9 6" />
  </svg>
);

const ChevronsRightIcon: React.FC<{ className?: string }> = ({ className = "w-4 h-4" }) => (
  <svg
    viewBox="0 0 24 24"
    className={className}
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <polyline points="13 17 18 12 13 7" />
    <polyline points="6 17 11 12 6 7" />
  </svg>
);

/**
 * Generate smart pagination items with ellipses (...)
 */
function getPaginationRange(current: number, total: number): Array<number | "..."> {
  if (total <= 7) {
    const range: number[] = [];
    for (let i = 1; i <= total; i++) {
      range.push(i);
    }
    return range;
  }

  // Large page count: smart ellipsis
  if (current <= 4) {
    return [1, 2, 3, 4, 5, "...", total];
  }

  if (current >= total - 3) {
    return [1, "...", total - 4, total - 3, total - 2, total - 1, total];
  }

  return [1, "...", current - 1, current, current + 1, "...", total];
}

export const DataTablePagination: React.FC<IDataTablePaginationProps> = ({
  currentPage,
  totalPages,
  totalItems,
  startIndex,
  endIndex,
  onPageChange,
  className = "",
}) => {
  if (totalItems === 0) return null;

  const range = getPaginationRange(currentPage, totalPages);
  const startRecord = totalItems === 0 ? 0 : startIndex + 1;
  const endRecord = endIndex;

  return (
    <div
      className={`border-t border-slate-200/80 bg-white sm:bg-slate-50/50 px-4 py-1.5 sm:py-2 sm:px-6 flex flex-col sm:flex-row items-center justify-between gap-1.5 select-none ${className}`}
    >
      {/* Left Side: Summary Typography */}
      <div className="text-xs sm:text-sm text-slate-600 font-medium order-2 sm:order-1 text-center sm:text-left">
        Showing <span className="font-bold text-slate-900">{startRecord}</span>–
        <span className="font-bold text-slate-900">{endRecord}</span> of{" "}
        <span className="font-bold text-slate-900">{totalItems}</span> records
      </div>

      {/* Right Side: Page Controls */}
      <div className="flex items-center gap-1 sm:gap-1.5 order-1 sm:order-2 flex-wrap justify-center">
        {/* First Page */}
        <button
          type="button"
          onClick={() => onPageChange(1)}
          disabled={currentPage === 1}
          title="First Page"
          className="h-8 w-8 sm:h-8.5 sm:w-8.5 inline-flex items-center justify-center rounded-lg border border-slate-200/90 bg-white text-slate-600 hover:text-slate-950 hover:bg-slate-50 hover:border-slate-300 transition-all shadow-2xs disabled:opacity-30 disabled:pointer-events-none cursor-pointer"
        >
          <ChevronsLeftIcon className="w-4 h-4" />
        </button>

        {/* Previous Page */}
        <button
          type="button"
          onClick={() => onPageChange(currentPage - 1)}
          disabled={currentPage === 1}
          title="Previous Page"
          className="h-8 w-8 sm:h-8.5 sm:w-8.5 inline-flex items-center justify-center rounded-lg border border-slate-200/90 bg-white text-slate-600 hover:text-slate-950 hover:bg-slate-50 hover:border-slate-300 transition-all shadow-2xs disabled:opacity-30 disabled:pointer-events-none cursor-pointer"
        >
          <ChevronLeftIcon className="w-4 h-4" />
        </button>

        {/* Page Number Buttons / Ellipses */}
        <div className="flex items-center gap-1">
          {range.map((item, index) => {
            if (item === "...") {
              return (
                <span
                  key={`ellipsis-${index}`}
                  className="h-8 w-6 sm:h-8.5 sm:w-8 inline-flex items-center justify-center text-xs font-bold text-slate-400 select-none"
                >
                  …
                </span>
              );
            }

            const pageNum = item as number;
            const isActive = pageNum === currentPage;

            return (
              <button
                key={`page-${pageNum}`}
                type="button"
                onClick={() => onPageChange(pageNum)}
                aria-current={isActive ? "page" : undefined}
                className={`h-8 min-w-[32px] sm:h-8.5 sm:min-w-[34px] px-2 text-xs sm:text-[13px] font-bold rounded-lg transition-all cursor-pointer flex items-center justify-center ${
                  isActive
                    ? "bg-[#012964] text-white border border-[#012964] shadow-xs shadow-[#012964]/25 ring-1 ring-[#012964]/20"
                    : "bg-white text-slate-700 hover:text-slate-950 hover:bg-slate-50 hover:border-slate-300 border border-slate-200/90 shadow-2xs"
                }`}
              >
                {pageNum}
              </button>
            );
          })}
        </div>

        {/* Next Page */}
        <button
          type="button"
          onClick={() => onPageChange(currentPage + 1)}
          disabled={currentPage >= totalPages}
          title="Next Page"
          className="h-8 w-8 sm:h-8.5 sm:w-8.5 inline-flex items-center justify-center rounded-lg border border-slate-200/90 bg-white text-slate-600 hover:text-slate-950 hover:bg-slate-50 hover:border-slate-300 transition-all shadow-2xs disabled:opacity-30 disabled:pointer-events-none cursor-pointer"
        >
          <ChevronRightIcon className="w-4 h-4" />
        </button>

        {/* Last Page */}
        <button
          type="button"
          onClick={() => onPageChange(totalPages)}
          disabled={currentPage >= totalPages}
          title="Last Page"
          className="h-8 w-8 sm:h-8.5 sm:w-8.5 inline-flex items-center justify-center rounded-lg border border-slate-200/90 bg-white text-slate-600 hover:text-slate-950 hover:bg-slate-50 hover:border-slate-300 transition-all shadow-2xs disabled:opacity-30 disabled:pointer-events-none cursor-pointer"
        >
          <ChevronsRightIcon className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
};
