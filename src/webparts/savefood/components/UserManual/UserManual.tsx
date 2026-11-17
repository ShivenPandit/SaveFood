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
      <div className="mb-3.5 flex items-center justify-between w-full">
        <div>
          <BackButton onClick={onBack} />
          <h2 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
            User Manual
          </h2>
          <p className="text-xs sm:text-sm text-slate-500">
            Operating procedures and guidelines.
          </p>
        </div>
      </div>

      {/* Full-Bleed Card Container */}
      <div
        style={{ height: "auto", flex: "0 0 auto", padding: "24px" }}
        className="w-full bg-white rounded-2xl border-2 border-slate-200/90 shadow-sm text-center"
      >
        <div className="w-12 h-12 mx-auto rounded-xl bg-blue-50 text-[#004B87] flex items-center justify-center text-2xl mb-2.5 shadow-2xs">
          📖
        </div>
        <h3 className="text-base sm:text-lg font-bold text-slate-800">User Manual</h3>
        <p className="text-xs sm:text-sm text-slate-500 max-w-md mx-auto mt-1">
          Process documentation, caterer policies, and user manuals will be published here.
        </p>
      </div>
    </div>
  );
};
