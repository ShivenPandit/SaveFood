import * as React from "react";

export interface IUseDynamicPaginationOptions<T> {
  items: T[];
  bufferBottom?: number;
  defaultRowHeight?: number;
  minPageSize?: number;
  maxPageSize?: number;
  initialPage?: number;
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
 * Initial viewport estimation based on standard screen resolutions
 */
const estimatePageSizeFromViewport = (): number => {
  if (typeof window === "undefined") return 10;

  const w = window.innerWidth;
  const h = window.innerHeight;

  // Mobile viewport (< 640px)
  if (w < 640) {
    return 3;
  }

  // 4K & Ultra-Tall Displays (>= 1800px vertical)
  if (h >= 1800) {
    return 36;
  }

  // 1440p QHD Displays (>= 1300px vertical)
  if (h >= 1300) {
    return 22;
  }

  // 1080p FHD Displays (>= 950px vertical)
  if (h >= 950) {
    return 14;
  }

  // 768p Laptop Displays (>= 650px vertical)
  if (h >= 650) {
    return 8;
  }

  // Compact netbooks / small height windows
  return 5;
};

export function useDynamicPagination<T>({
  items,
  bufferBottom = 32,
  defaultRowHeight = 48,
  minPageSize = 3,
  maxPageSize = 50,
  initialPage = 1,
}: IUseDynamicPaginationOptions<T>): IUseDynamicPaginationResult<T> {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [currentPage, setCurrentPage] = React.useState<number>(initialPage);
  const [pageSize, setPageSize] = React.useState<number>(estimatePageSizeFromViewport);

  const recalculate = React.useCallback(() => {
    if (typeof window === "undefined") return;

    const container = containerRef.current;
    if (!container) {
      setPageSize(estimatePageSizeFromViewport());
      return;
    }

    const containerRect = container.getBoundingClientRect();
    const isMobile = window.innerWidth < 640;

    // Available vertical height from container top to viewport bottom
    const availableHeight = window.innerHeight - containerRect.top - bufferBottom;

    // Measure overhead elements (headers, footers, borders, filter bars)
    let overhead = 0;

    // Table header (thead) overhead
    const thead = container.querySelector("thead");
    if (thead) {
      overhead += thead.getBoundingClientRect().height;
    }

    // Pagination bar placeholder overhead (~54px)
    overhead += 54;

    // Additional container padding/margins
    overhead += 16;

    const netRowSpace = Math.max(0, availableHeight - overhead);

    // Live Row Height Detection
    let liveRowHeight = defaultRowHeight;
    if (isMobile) {
      const mobileCard = container.querySelector('[data-role="list-item"]');
      if (mobileCard) {
        const measuredHeight = mobileCard.getBoundingClientRect().height;
        if (measuredHeight > 30) {
          liveRowHeight = measuredHeight + 8; // include card gap
        }
      } else {
        liveRowHeight = 115; // default mobile card estimate
      }
    } else {
      const row = container.querySelector("tbody tr");
      if (row) {
        const measuredHeight = row.getBoundingClientRect().height;
        if (measuredHeight > 20) {
          liveRowHeight = measuredHeight;
        }
      }
    }

    // Effective Page Size Calculation
    const calculatedRows = Math.floor(netRowSpace / liveRowHeight);
    const clampedSize = Math.min(maxPageSize, Math.max(minPageSize, calculatedRows));

    setPageSize((prev) => (prev !== clampedSize ? clampedSize : prev));
  }, [bufferBottom, defaultRowHeight, minPageSize, maxPageSize]);

  // Debounced Window Resize & Container ResizeObserver
  React.useEffect(() => {
    let timeoutId: number;

    const handleDebouncedResize = (): void => {
      window.clearTimeout(timeoutId);
      timeoutId = window.setTimeout(() => {
        recalculate();
      }, 120);
    };

    window.addEventListener("resize", handleDebouncedResize);

    // Initial measurement after DOM mount
    recalculate();

    // Use ResizeObserver on container if available
    let resizeObserver: ResizeObserver | null = null;
    if (typeof ResizeObserver !== "undefined" && containerRef.current) {
      resizeObserver = new ResizeObserver(() => {
        handleDebouncedResize();
      });
      resizeObserver.observe(containerRef.current);
    }

    return () => {
      window.removeEventListener("resize", handleDebouncedResize);
      window.clearTimeout(timeoutId);
      if (resizeObserver) {
        resizeObserver.disconnect();
      }
    };
  }, [recalculate]);

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
