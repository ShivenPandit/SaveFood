import * as React from "react";

export interface IBackButtonProps {
  onClick: () => void;
  label?: string;
  className?: string;
}

export const BackButton: React.FC<IBackButtonProps> = ({
  onClick,
  label = "Back to Dashboard",
  className = "",
}) => {
  return (
    <button
      type="button"
      onClick={onClick}
      style={{
        boxShadow: "0 1px 2px rgba(0,0,0,0.04)",
      }}
      className={`group inline-flex items-center gap-2 px-4 py-2 text-[14.5px] font-semibold text-[#1e293b] bg-white border border-[#cbd5e1] rounded-xl hover:bg-[#f1f5f9] hover:border-[#94a3b8] transition-all duration-150 cursor-pointer select-none ${className}`}
    >
      <svg
        viewBox="0 0 24 24"
        className="w-4 h-4 text-slate-500 group-hover:text-[#1e293b] transform group-hover:-translate-x-0.5 transition-transform duration-150 fill-none stroke-current"
        strokeWidth="2.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      >
        <line x1="19" y1="12" x2="5" y2="12" />
        <polyline points="12 19 5 12 12 5" />
      </svg>
      <span>{label}</span>
    </button>
  );
};
