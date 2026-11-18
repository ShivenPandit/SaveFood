import * as React from "react";

export interface IToastProps {
  message: string;
  type?: "success" | "error" | "info";
  onClose: () => void;
  duration?: number;
}

export const Toast: React.FC<IToastProps> = ({
  message,
  type = "success",
  onClose,
  duration = 3500,
}) => {
  const [isVisible, setIsVisible] = React.useState<boolean>(false);

  React.useEffect(() => {
    // Trigger smooth slide/fade entrance
    const enterTimer = setTimeout(() => {
      setIsVisible(true);
    }, 10);

    // Auto-dismiss after duration
    const dismissTimer = setTimeout(() => {
      setIsVisible(false);
      setTimeout(onClose, 250); // wait for fade out animation
    }, duration);

    return () => {
      clearTimeout(enterTimer);
      clearTimeout(dismissTimer);
    };
  }, [duration, onClose]);

  const handleManualDismiss = (): void => {
    setIsVisible(false);
    setTimeout(onClose, 250);
  };

  const isSuccess = type === "success";
  const isError = type === "error";

  return (
    <div
      role="status"
      aria-live="polite"
      style={{
        position: "fixed",
        top: "20px",
        left: "50%",
        transform: isVisible
          ? "translateX(-50%) translateY(0)"
          : "translateX(-50%) translateY(-16px)",
        opacity: isVisible ? 1 : 0,
        transition: "all 0.25s cubic-bezier(0.16, 1, 0.3, 1)",
        zIndex: 9999,
        pointerEvents: "auto",
      }}
      className="select-none flex items-center gap-3 px-4 py-2 sm:px-5 sm:py-2.5 rounded-full shadow-xl shadow-slate-950/20 backdrop-blur-md bg-slate-900/95 text-white border border-slate-700/60 max-w-[90vw] sm:max-w-md"
    >
      {/* Icon */}
      <div className="shrink-0 flex items-center justify-center">
        {isSuccess && (
          <div className="w-5 h-5 rounded-full bg-emerald-500/20 text-emerald-400 flex items-center justify-center border border-emerald-500/40">
            <svg
              viewBox="0 0 24 24"
              className="w-3.5 h-3.5"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.8"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <polyline points="20 6 9 17 4 12" />
            </svg>
          </div>
        )}
        {isError && (
          <div className="w-5 h-5 rounded-full bg-rose-500/20 text-rose-400 flex items-center justify-center border border-rose-500/40">
            <svg
              viewBox="0 0 24 24"
              className="w-3.5 h-3.5"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.8"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <circle cx="12" cy="12" r="10" />
              <line x1="12" y1="8" x2="12" y2="12" />
              <line x1="12" y1="16" x2="12.01" y2="16" />
            </svg>
          </div>
        )}
        {!isSuccess && !isError && (
          <div className="w-5 h-5 rounded-full bg-sky-500/20 text-sky-400 flex items-center justify-center border border-sky-500/40">
            <svg
              viewBox="0 0 24 24"
              className="w-3.5 h-3.5"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.8"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <circle cx="12" cy="12" r="10" />
              <line x1="12" y1="16" x2="12" y2="12" />
              <line x1="12" y1="8" x2="12.01" y2="8" />
            </svg>
          </div>
        )}
      </div>

      {/* Message Text */}
      <span className="text-[14px] sm:text-[14.5px] font-semibold text-slate-100 tracking-tight truncate">
        {message}
      </span>

      {/* Dismiss Button */}
      <button
        type="button"
        onClick={handleManualDismiss}
        className="ml-1 text-slate-400 hover:text-white transition-colors cursor-pointer p-0.5 rounded-full hover:bg-white/10 shrink-0"
        title="Dismiss"
        aria-label="Dismiss notification"
      >
        <svg
          viewBox="0 0 24 24"
          className="w-3.5 h-3.5"
          fill="none"
          stroke="currentColor"
          strokeWidth="2.5"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          <line x1="18" y1="6" x2="6" y2="18" />
          <line x1="6" y1="6" x2="18" y2="18" />
        </svg>
      </button>
    </div>
  );
};
