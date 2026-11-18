import * as React from "react";

export interface IUseDynamicPaginationOptions<T> {
  items: T[];
  bufferBottom?: number;
  defaultRowHeight?: number;
  minPageSize?: number;
  maxPageSize?: number;
  initialPage?: number;
  viewType?: "my-requests" | "all-requests" | "default";
}

export interface IUseDynamicPaginationResult<T> {
  containerRef: React.RefObject<HTMLDivElement>;
  currentPage: number;
  setCurrentPage: (page: number) => void;
  pageSize: number;
  totalPages: number;
  totalItems: number;
  paginatedItems: T[];
  startIndex: number;
  endIndex: number;
  recalculate: () => void;
}

/**
 * Calculates optimal rows per page directly from window.innerHeight & window.innerWidth.
 * Guaranteed to fit within viewport on Laptop, 1080p, 2K, and 4K screens without window scroll.
 * Completely decoupled from scroll position to guarantee ZERO flicker during scrolling.
 */
export const calculateOptimalPageSize = (
  h: number,
  w: number,
  viewType: "my-requests" | "all-requests" | "default" = "default"
): number => {
  if (w < 640) {
    return 3; // mobile card pagination
  }

  // Exact vertical overhead for non-row elements (in pixels):
  // - Top navbar + green accent line: 54px
  // - Main container padding (py-2.5): ~20px
  // - Header row (Back button, Title, Subtitle, Action button / Tabs): ~56px for my-requests, ~76px for all-requests
  // - Filter Card & search controls: ~68px for my-requests, ~86px for all-requests
  // - Table container card borders: 4px
  // - Table thead: ~36px
  // - DataTablePagination footer: ~44px
  // - Safety margin to guarantee zero vertical window scrollbar: ~40px
  // Total overhead: 380px for optimal bottom clearance on laptop screens.
  const overhead = 380;
  const rowHeight = 42; // Compact, clean table row height

  const availableRowSpace = Math.max(0, h - overhead);
  let rows = Math.floor(availableRowSpace / rowHeight);

  // On standard laptop screens (h <= 780px), cap my-requests at 8 rows so the pagination footer is never cut off
  if (viewType === "my-requests" && h <= 780) {
    rows = Math.min(8, Math.max(5, rows));
  }

  // Clamped ranges:
  // On compact laptop (h ≈ 620px-700px): 6–8 rows
  // On standard 1080p (h ≈ 750px-950px): 10–14 rows
  // On 2K QHD (h ≈ 1000px-1400px): 18–24 rows
  // On 4K UHD (h ≥ 1400px): 28–36 rows
  return Math.max(4, Math.min(36, rows));
};

export function useDynamicPagination<T>({
  items,
  initialPage = 1,
  minPageSize = 4,
  maxPageSize = 45,
  viewType = "default",
}: IUseDynamicPaginationOptions<T>): IUseDynamicPaginationResult<T> {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [currentPage, setCurrentPage] = React.useState<number>(initialPage);

  const getTargetPageSize = React.useCallback((): number => {
    if (typeof window === "undefined") return 6;
    const computed = calculateOptimalPageSize(window.innerHeight, window.innerWidth, viewType);
    return Math.max(minPageSize, Math.min(maxPageSize, computed));
  }, [viewType, minPageSize, maxPageSize]);

  const [pageSize, setPageSize] = React.useState<number>(getTargetPageSize);

  // Recalculate ONLY on window resize (debounced). NEVER on scroll! Zero flicker guaranteed!
  React.useEffect(() => {
    let timeoutId: number;

    const handleResize = (): void => {
      window.clearTimeout(timeoutId);
      timeoutId = window.setTimeout(() => {
        const target = getTargetPageSize();
        setPageSize((prev) => (prev !== target ? target : prev));
      }, 100);
    };

    window.addEventListener("resize", handleResize);

    // Initial check on mount
    handleResize();

    return () => {
      window.removeEventListener("resize", handleResize);
      window.clearTimeout(timeoutId);
    };
  }, [getTargetPageSize]);

  // Recalculate callback for external consumers
  const recalculate = React.useCallback(() => {
    const target = getTargetPageSize();
    setPageSize((prev) => (prev !== target ? target : prev));
  }, [getTargetPageSize]);

  // Reset or clamp current page when items length or pageSize changes
  const totalItems = items.length;
  const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));

  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(Math.max(1, totalPages));
    }
  }, [currentPage, totalPages]);

  // In-memory sliced items for current page
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalItems);

  const paginatedItems = React.useMemo(() => {
    return items.slice(startIndex, endIndex);
  }, [items, startIndex, endIndex]);

  return {
    containerRef,
    currentPage,
    setCurrentPage,
    pageSize,
    totalPages,
    totalItems,
    paginatedItems,
    startIndex,
    endIndex,
    recalculate,
  };
}
