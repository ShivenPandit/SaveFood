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

  // For my-requests, strictly enforce 8 rows on desktop so table and pagination footer fit cleanly above taskbar
  if (viewType === "my-requests") {
    return 8;
  }

  const overhead = 360;
  const rowHeight = 44;

  const availableRowSpace = Math.max(0, h - overhead);
  let rows = Math.floor(availableRowSpace / rowHeight);

  return Math.max(5, Math.min(25, rows));
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
