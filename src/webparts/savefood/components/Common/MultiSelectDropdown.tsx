import * as React from "react";

export interface IMultiSelectOption<T extends number | string = number | string> {
  id: T;
  label: string;
}

export interface IMultiSelectDropdownProps<T extends number | string = number | string> {
  label?: string;
  allLabel?: string;
  itemName?: string;
  options: IMultiSelectOption<T>[];
  selectedIds: T[];
  onChange: (selectedIds: T[]) => void;
  className?: string;
}

export const MultiSelectDropdown = <T extends number | string = number | string>({
  allLabel = "All Mess Locations",
  itemName = "Locations",
  options,
  selectedIds,
  onChange,
  className = "",
}: IMultiSelectDropdownProps<T>): React.ReactElement => {
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const [searchQuery, setSearchQuery] = React.useState<string>("");
  const dropdownRef = React.useRef<HTMLDivElement>(null);

  // Close on outside click
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    if (isOpen) {
      document.addEventListener("mousedown", handleClickOutside);
    }
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [isOpen]);

  // Is all selected? (Empty array or all IDs selected = all)
  const isAllSelected = selectedIds.length === 0 || selectedIds.length === options.length;

  // Filter options by search query
  const filteredOptions = React.useMemo(() => {
    if (!searchQuery.trim()) return options;
    const q = searchQuery.toLowerCase().trim();
    return options.filter((opt) => opt.label.toLowerCase().indexOf(q) !== -1);
  }, [options, searchQuery]);

  // Handle toggling an individual option
  const handleToggle = (id: T): void => {
    if (isAllSelected) {
      // If currently all selected, clicking one selects ONLY that one
      onChange([id]);
      return;
    }

    if (selectedIds.indexOf(id) !== -1) {
      const next = selectedIds.filter((item) => item !== id);
      // If deselecting leaves 0 selected, reset to all
      onChange(next);
    } else {
      const next = [...selectedIds, id];
      if (next.length === options.length) {
        onChange([]); // all selected represented as []
      } else {
        onChange(next);
      }
    }
  };

  const handleSelectAll = (): void => {
    onChange([]); // [] means all locations selected
  };

  const handleClearAll = (): void => {
    onChange([]);
  };

  // Compute trigger button label
  const triggerLabel = React.useMemo(() => {
    if (isAllSelected || selectedIds.length === 0) {
      return allLabel;
    }
    if (selectedIds.length === 1) {
      const match = options.find((o) => o.id === selectedIds[0]);
      return match ? match.label : allLabel;
    }
    if (selectedIds.length === 2) {
      const labels = selectedIds
        .map((id) => options.find((o) => o.id === id)?.label)
        .filter(Boolean);
      return labels.join(", ");
    }
    return `${selectedIds.length} ${itemName} Selected`;
  }, [isAllSelected, selectedIds, options, allLabel, itemName]);

  return (
    <div ref={dropdownRef} className={`relative w-full ${className}`}>
      {/* Trigger Button - Minimal and Compact */}
      <button
        type="button"
        onClick={() => setIsOpen((prev) => !prev)}
        className={`w-full h-10 px-3.5 rounded-xl border text-[13.5px] font-medium flex items-center justify-between transition cursor-pointer select-none text-left ${
          isOpen
            ? "border-[#004B87] ring-2 ring-[#004B87]/20 bg-white text-slate-900"
            : !isAllSelected && selectedIds.length > 0
            ? "border-[#004B87] bg-blue-50/40 text-[#004B87]"
            : "border-slate-300 bg-white text-slate-800 hover:border-slate-400"
        }`}
        aria-haspopup="listbox"
        aria-expanded={isOpen}
      >
        <span className="truncate pr-1">{triggerLabel}</span>

        <div className="flex items-center gap-1.5 shrink-0">
          {!isAllSelected && selectedIds.length > 0 && (
            <span className="w-5 h-5 rounded-full bg-[#004B87] text-white text-[11px] font-bold flex items-center justify-center">
              {selectedIds.length}
            </span>
          )}
          <svg
            viewBox="0 0 24 24"
            className={`w-4 h-4 text-slate-400 transition-transform duration-150 ${
              isOpen ? "rotate-180 text-[#004B87]" : ""
            }`}
            fill="none"
            stroke="currentColor"
            strokeWidth="2.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <polyline points="6 9 12 15 18 9" />
          </svg>
        </div>
      </button>

      {/* Dropdown Menu */}
      {isOpen && (
        <div className="absolute left-0 right-0 top-full mt-1.5 bg-white rounded-xl border border-slate-200 shadow-xl z-50 overflow-hidden animate-in fade-in zoom-in-95 duration-100">
          {/* Quick Actions Row */}
          <div className="p-2 border-b border-slate-100 bg-slate-50/70 flex items-center justify-between gap-2 text-[12px]">
            <span className="font-semibold text-slate-500">
              {isAllSelected ? "All Selected" : `${selectedIds.length} of ${options.length} Selected`}
            </span>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleSelectAll}
                className="font-bold text-[#004B87] hover:underline cursor-pointer"
              >
                Select All
              </button>
              <span className="text-slate-300">|</span>
              <button
                type="button"
                onClick={handleClearAll}
                className="font-bold text-slate-500 hover:text-slate-700 hover:underline cursor-pointer"
              >
                Reset
              </button>
            </div>
          </div>

          {/* Search Input if more than 5 options */}
          {options.length > 5 && (
            <div className="p-2 border-b border-slate-100">
              <div className="relative">
                <input
                  type="text"
                  placeholder="Search canteen location..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full h-8 pl-7 pr-2 text-[12.5px] rounded-lg border border-slate-200 bg-slate-50 focus:bg-white focus:outline-none focus:ring-1 focus:ring-[#004B87]"
                />
                <svg
                  viewBox="0 0 24 24"
                  className="w-3.5 h-3.5 text-slate-400 absolute left-2 top-2.5 fill-none stroke-current"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <circle cx="11" cy="11" r="8" />
                  <line x1="21" y1="21" x2="16.65" y2="16.65" />
                </svg>
              </div>
            </div>
          )}

          {/* Options Checklist */}
          <div className="max-h-48 overflow-y-auto p-1.5 space-y-0.5">
            {filteredOptions.length === 0 ? (
              <div className="p-3 text-center text-[12.5px] text-slate-400 font-medium">
                No option matches &quot;{searchQuery}&quot;
              </div>
            ) : (
              filteredOptions.map((opt) => {
                const isChecked = isAllSelected || selectedIds.indexOf(opt.id) !== -1;
                return (
                  <label
                    key={String(opt.id)}
                    onClick={() => handleToggle(opt.id)}
                    className={`flex items-center gap-2.5 px-2.5 py-1.5 rounded-lg text-[13px] font-medium cursor-pointer transition select-none ${
                      isChecked
                        ? "bg-blue-50/60 text-[#004B87]"
                        : "text-slate-700 hover:bg-slate-100/70"
                    }`}
                  >
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={() => {}} // handled by parent label onClick
                      className="w-4 h-4 rounded border-slate-300 text-[#004B87] focus:ring-[#004B87] cursor-pointer"
                    />
                    <span className="flex-1 truncate">{opt.label}</span>
                  </label>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
};
