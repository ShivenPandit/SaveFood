import * as React from "react";
import { BackButton } from "../Common/BackButton";

export interface IUserManualProps {
  onBack: () => void;
  onNavigateCreate: () => void;
}

export const UserManual: React.FC<IUserManualProps> = ({ onBack }) => {
  return (
    <div className="w-full select-none" style={{ height: "auto", flex: "0 0 auto" }}>
      {/* Top Header with Back button and Heading */}
      <div className="mb-2.5 flex items-center justify-between w-full">
        <div>
          <BackButton onClick={onBack} className="mb-1.5" />
          <h2 className="text-[26px] sm:text-[28px] font-bold text-[#0f172a] tracking-tight">
            User Manual
          </h2>
          <p className="text-[15px] text-[#475569] leading-relaxed mt-1">
            Operating procedures and cancellation guidelines.
          </p>
        </div>
      </div>

      {/* Full-Bleed Card Container */}
      <div
        style={{ height: "auto", flex: "0 0 auto", padding: "32px 24px" }}
        className="w-full bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm text-center"
      >
        <div className="w-14 h-14 mx-auto rounded-2xl bg-blue-50 text-[#004B87] flex items-center justify-center text-3xl mb-3 shadow-2xs">
          📖
        </div>
        <h3 className="text-lg sm:text-xl font-bold text-slate-800">Save Food Guidelines & Standard Operating Procedures</h3>
        <p className="text-[15px] text-[#475569] max-w-lg mx-auto mt-2 leading-relaxed">
          Process documentation, cut-off schedules, and caterer coordination manuals will be published here.
        </p>
      </div>
    </div>
  );
};
