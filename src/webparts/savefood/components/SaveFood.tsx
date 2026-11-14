import * as React from "react";
import { ISaveFoodProps } from "./ISaveFoodProps";
import { VEDANTA_LOGO_BASE64 } from "../assets/logoBase64";
import {
  CreateOptOutIllustration,
  MyRequestsIllustration,
  AllRequestsIllustration,
  ReportsIllustration,
  UserManualIllustration,
  AdminConfigIllustration,
} from "./HomeIllustrations";

interface IHomeActionCard {
  id: string;
  title: string;
  description: string;
  actionLabel: string;
  isPrimary?: boolean;
  illustration: React.FC<{ className?: string }>;
}

const SaveFood: React.FC<ISaveFoodProps> = (props) => {
  const siteAssetLogoUrl = `${props.newsiteUrl}/SiteAssets/Vedanta1.png`;
  const [logoSrc, setLogoSrc] = React.useState<string>(siteAssetLogoUrl);

  // Container width measurement via ResizeObserver for responsive grid
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [containerWidth, setContainerWidth] = React.useState<number>(1200);
  const [isDesktop, setIsDesktop] = React.useState<boolean>(
    typeof window !== "undefined" ? window.innerWidth >= 768 : true
  );

  React.useEffect(() => {
    const handleResize = (): void => {
      if (typeof window !== "undefined") {
        setIsDesktop(window.innerWidth >= 768);
      }
    };

    window.addEventListener("resize", handleResize);

    const el = containerRef.current;
    if (el && el.offsetWidth > 0) {
      setContainerWidth(el.offsetWidth);
    }

    let observer: ResizeObserver | null = null;
    if (el && typeof ResizeObserver !== "undefined") {
      observer = new ResizeObserver((entries) => {
        for (const entry of entries) {
          if (entry.contentRect.width > 0) {
            setContainerWidth(entry.contentRect.width);
          }
        }
      });
      observer.observe(el);
    }

    return () => {
      window.removeEventListener("resize", handleResize);
      if (observer) observer.disconnect();
    };
  }, []);

  // Suppress outer scroll on desktop and lock viewport to 100vh with zero bottom gap
  React.useEffect(() => {
    if (isDesktop) {
      document.documentElement.classList.add("savefood-dashboard-no-scroll");
      document.body.classList.add("savefood-dashboard-no-scroll");
    } else {
      document.documentElement.classList.remove("savefood-dashboard-no-scroll");
      document.body.classList.remove("savefood-dashboard-no-scroll");
    }

    return () => {
      document.documentElement.classList.remove("savefood-dashboard-no-scroll");
      document.body.classList.remove("savefood-dashboard-no-scroll");
    };
  }, [isDesktop]);

  // Determine column count based on available container width
  const cols = containerWidth >= 940 ? 3 : containerWidth >= 580 ? 2 : 1;

  // Fallback to embedded base64 Vedanta logo if SiteAssets image fails
  const handleLogoError = (): void => {
    if (logoSrc !== VEDANTA_LOGO_BASE64) {
      setLogoSrc(VEDANTA_LOGO_BASE64);
    }
  };

  // Format current date matching enterprise pattern: "Thu, Nov 12, 2026"
  const formattedDate = React.useMemo(() => {
    try {
      const options: Intl.DateTimeFormatOptions = {
        weekday: "short",
        day: "numeric",
        month: "short",
        year: "numeric",
      };
      return new Date().toLocaleDateString("en-US", options);
    } catch {
      return "";
    }
  }, []);

  // Card click handler without any toast banner
  const handleCardClick = (cardTitle: string): void => {
    console.log(`[SaveFood] Action selected: ${cardTitle}`);
  };

  // User Initials
  const cleanName = (props.newcurrentUserName || "User").replace(/\s*\([^)]*\)/g, "").trim();
  const initials = cleanName
    .split(" ")
    .map((n) => n[0])
    .join("")
    .substring(0, 2)
    .toUpperCase();

  // 6 Action Cards matching reference design: Highlighted Primary Card + 5 Clean Cards with 3D Vectors
  const actionCards: IHomeActionCard[] = [
    {
      id: "create-optout",
      title: "Create Meal Opt-Out Request",
      description:
        "Submit advance meal cancellations for upcoming shifts, leave, or plant travel.",
      actionLabel: "Create Opt-Out",
      isPrimary: true,
      illustration: CreateOptOutIllustration,
    },
    {
      id: "my-optouts",
      title: "My Meal Opt-Out Requests",
      description:
        "Track submitted meal cancellations and monitor advance notice status.",
      actionLabel: "View Requests",
      isPrimary: false,
      illustration: MyRequestsIllustration,
    },
    {
      id: "all-optouts",
      title: "All Meal Opt-Out Requests",
      description:
        "Consolidated kitchen roster, caterer headcounts, and location-wise records.",
      actionLabel: "View Roster",
      isPrimary: false,
      illustration: AllRequestsIllustration,
    },
    {
      id: "reports",
      title: "Reports & Dashboard",
      description:
        "Real-time food savings metrics, waste reduction trends, and caterer accuracy.",
      actionLabel: "View Reports",
      isPrimary: false,
      illustration: ReportsIllustration,
    },
    {
      id: "admin-configuration",
      title: "Admin Configuration",
      description:
        "Configure plant canteen units, manage meal cut-off windows, and system roles.",
      actionLabel: "Open Console",
      isPrimary: false,
      illustration: AdminConfigIllustration,
    },
    {
      id: "manual",
      title: "User Manual",
      description:
        "View process guidelines, meal booking cut-off policies, FAQs, and support.",
      actionLabel: "View Guide",
      isPrimary: false,
      illustration: UserManualIllustration,
    },
  ];

  return (
    <div
      style={{
        height: isDesktop ? "100vh" : "auto",
        minHeight: isDesktop ? "100vh" : "100%",
        maxHeight: isDesktop ? "100vh" : undefined,
        overflowY: isDesktop ? "hidden" : "auto",
        overflowX: "hidden",
        backgroundColor: "#F8FAFC",
      }}
      className="savefood-canvas-breakout w-full max-w-none flex-1 flex flex-col justify-between text-slate-800 antialiased font-sans select-none box-border"
    >
      {/* 1. Top Navbar: Vedanta Logo & User Profile (Edge-to-Edge White Header matching CER) */}
      <header className="h-12 sm:h-13 lg:h-14 bg-white border-b border-slate-200/90 shrink-0 shadow-2xs z-30 sticky top-0 w-full select-none px-4 sm:px-6 lg:px-8 xl:px-10 2xl:px-12 flex items-center justify-between">
        {/* Left Side: Brand Logo + System Identity */}
        <div className="flex items-center gap-2.5 sm:gap-3.5 min-w-0">
          <img
            src={logoSrc}
            alt="Vedanta Logo"
            className="h-7 sm:h-7.5 lg:h-8 w-auto object-contain transition-transform"
            onError={handleLogoError}
          />
          <div className="h-5 w-px bg-slate-200 hidden xs:block" />
          <div className="flex items-center min-w-0">
            <span className="text-xs sm:text-sm lg:text-base font-extrabold text-[#003865] tracking-tight truncate">
              Save Food Management System
            </span>
          </div>
        </div>

        {/* Right Side: Date & User Profile Avatar Pill */}
        <div className="flex items-center gap-2 sm:gap-3">
          <span className="text-[11px] sm:text-xs text-slate-500 font-medium hidden md:block">
            {formattedDate}
          </span>
          <div className="flex items-center gap-2 bg-slate-50 border border-slate-200/80 rounded-full py-1 px-1.5 sm:px-2.5 shadow-2xs">
            <div className="w-6.5 h-6.5 sm:w-7 sm:h-7 rounded-full bg-[#004B87] text-white flex items-center justify-center font-bold text-[10px] sm:text-[11px] shrink-0">
              {initials}
            </div>
            <div className="text-left hidden sm:block pr-1">
              <p className="text-[11.5px] font-bold text-slate-900 leading-tight truncate max-w-[160px]">
                {props.newcurrentUserName || "User"}
              </p>
              <p className="text-[9.5px] text-slate-500 font-medium leading-none mt-0.5">
                {props.newisSiteAdmin ? "Administrator" : "Employee / Initiator"}
              </p>
            </div>
          </div>
        </div>
      </header>

      {/* Signature Green Accent Line */}
      <div className="w-full h-[2px] bg-[#009639] shrink-0" />

      {/* 2. Full-Bleed Contained Hero Banner with Angled Geometric Ribbons */}
      <section className="relative w-full overflow-hidden bg-gradient-to-r from-[#002D56] via-[#004B87] to-[#046A38] text-white select-none shadow-md shrink-0 py-4 sm:py-5 lg:py-6 px-4 sm:px-6 lg:px-8 xl:px-10 2xl:px-12">
        {/* Deep Navy Gradient Mask */}
        <div className="absolute inset-0 bg-gradient-to-r from-[#002D56] via-[#004B87]/90 via-55% to-transparent pointer-events-none z-10" />

        {/* Angled Translucent Geometric Overlay Slashes */}
        <div className="absolute inset-y-0 left-1/3 md:left-2/5 lg:left-1/2 w-96 pointer-events-none z-10 overflow-hidden hidden sm:block">
          {/* Sky Blue Translucent Diagonal Ribbon */}
          <div
            className="absolute -top-36 bottom-[-140px] w-12 lg:w-14 bg-sky-400/25 transform -skew-x-[26deg] backdrop-blur-[1px]"
            style={{ left: "10%" }}
          />
          {/* Gold/Yellow Accent Diagonal Ribbon */}
          <div
            className="absolute -top-36 bottom-[-140px] w-4 lg:w-5 bg-[#F59E0B]/80 transform -skew-x-[26deg] shadow-lg"
            style={{ left: "26%" }}
          />
          {/* Royal Blue Translucent Diagonal Ribbon */}
          <div
            className="absolute -top-36 bottom-[-140px] w-20 lg:w-24 bg-[#1E3A8A]/45 transform -skew-x-[26deg]"
            style={{ left: "34%" }}
          />
        </div>

        {/* Main Banner Content */}
        <div className="relative z-20 flex flex-col md:flex-row md:items-center justify-between gap-3 sm:gap-4 min-w-0">
          {/* Left Column: Greeting & Core System Messaging */}
          <div className="max-w-2xl xl:max-w-3xl min-w-0">
            {/* Gold Accent Bar */}
            <div className="w-10 sm:w-12 h-1.5 bg-[#F59E0B] rounded-full mb-1.5 sm:mb-2" />

            <p className="text-white/90 text-xs sm:text-sm font-semibold tracking-wider uppercase">
              Vedanta Sustainability Platform
            </p>

            <h1 className="font-extrabold text-white tracking-tight leading-tight text-lg sm:text-2xl lg:text-3xl mt-1">
              VOGL Save Food{" "}
              <span className="text-emerald-400 font-black whitespace-nowrap">
                Management System
              </span>
            </h1>

            <p className="text-amber-300 font-extrabold tracking-widest text-xs sm:text-sm mt-1 uppercase">
              Inform Early. Prepare Smart. Waste Less.
            </p>

            <p className="text-blue-100/90 leading-relaxed text-xs sm:text-[13px] lg:text-sm mt-1.5 max-w-xl lg:max-w-2xl 2xl:max-w-3xl">
              A centralized digital platform to manage advance meal opt-out requests, optimize catering planning, minimize food wastage, and promote sustainability across VOGL locations.
            </p>
          </div>

          {/* Right Column: Clean Informational Badge */}
          <div className="hidden lg:flex flex-col items-end text-right justify-center pr-2 shrink-0">
            <div className="bg-white/10 backdrop-blur-md border border-white/20 rounded-2xl p-3.5 sm:p-4 max-w-[230px] text-center shadow-lg">
              <span className="text-[10px] font-bold tracking-wider text-emerald-300 uppercase block">
                Responsible Living
              </span>
              <p className="text-xs font-semibold text-white mt-1 leading-snug">
                Every advance opt-out preserves food &amp; energy.
              </p>
              <div className="w-8 h-0.5 bg-[#F59E0B] rounded-full mx-auto mt-2" />
            </div>
          </div>
        </div>
      </section>

      {/* 3. 6 Crisp White Action Cards with "Pop Up & Outline" on Hover (Zero Pastel Fills) */}
      <div
        ref={containerRef}
        className="w-full flex-1 min-h-0 flex flex-col justify-center px-4 sm:px-6 lg:px-8 xl:px-10 2xl:px-12 py-3 sm:py-4 lg:py-5 overflow-visible md:overflow-hidden"
      >
        <div
          className={`grid ${
            cols === 3
              ? "grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 lg:grid-rows-2"
              : cols === 2
              ? "grid-cols-1 sm:grid-cols-2"
              : "grid-cols-1"
          } gap-3 sm:gap-4 lg:gap-5 h-full flex-1 min-h-0 min-w-0`}
        >
          {actionCards.map((card) => {
            const Illustration = card.illustration;
            const isPrimary = card.isPrimary;

            return (
              <div
                key={card.id}
                onClick={() => handleCardClick(card.title)}
                className={`group relative flex flex-col justify-between p-4 sm:p-5 lg:p-5.5 xl:p-6 rounded-2xl ${
                  isPrimary
                    ? "bg-gradient-to-br from-[#EBFBF0] via-[#E2F7E9] to-[#D1F2DD] border-2 border-[#86EFAC] hover:border-[#10B981] hover:shadow-emerald-900/15"
                    : "bg-white border-2 border-slate-200/90 hover:border-[#0066CC] hover:shadow-slate-300/40"
                } shadow-xs hover:shadow-xl hover:-translate-y-1.5 transition-all duration-300 ease-out cursor-pointer overflow-hidden h-full min-h-[145px] sm:min-h-[160px] lg:min-h-[175px] select-none min-w-0`}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => {
                  if (e.key === "Enter" || e.key === " ") {
                    e.preventDefault();
                    handleCardClick(card.title);
                  }
                }}
              >
                {/* Decorative Bottom-Right Fluid Wave Auras */}
                {isPrimary ? (
                  <>
                    <div className="absolute -bottom-8 -right-8 w-44 h-44 rounded-full bg-emerald-200/40 blur-xl pointer-events-none" />
                    <div className="absolute -bottom-4 -right-4 w-32 h-32 rounded-full bg-teal-200/35 blur-lg pointer-events-none" />
                  </>
                ) : (
                  <>
                    <div className="absolute -bottom-8 -right-8 w-44 h-44 rounded-full bg-blue-100/45 blur-xl pointer-events-none" />
                    <div className="absolute -bottom-4 -right-4 w-32 h-32 rounded-full bg-sky-100/50 blur-lg pointer-events-none" />
                  </>
                )}

                {/* Main Content Layout: Heading + Description on Left, 3D Vector Icon on Right */}
                <div className="flex items-start justify-between gap-2.5 sm:gap-4 z-10 min-w-0">
                  {/* Left Column: Heading + 2-Line Subtitle Description */}
                  <div className="flex-1 min-w-0 pr-1 sm:pr-2">
                    <h3
                      className={`font-extrabold text-slate-900 tracking-tight leading-snug ${
                        !isPrimary ? "group-hover:text-[#004B87]" : ""
                      } transition-colors text-sm sm:text-base lg:text-[17px] xl:text-[18px]`}
                    >
                      {card.title}
                    </h3>
                    <p
                      className={`${
                        isPrimary ? "text-slate-600" : "text-slate-500"
                      } mt-1 sm:mt-1.5 leading-relaxed text-xs sm:text-[12.5px] xl:text-[13px] line-clamp-2`}
                    >
                      {card.description}
                    </p>
                  </div>

                  {/* Right Column: 3D Vector Illustration (Floating with soft ambient glow) */}
                  <div className="shrink-0 flex items-center justify-end pointer-events-none select-none pl-1 self-start -mt-1 sm:-mt-2">
                    <Illustration className="w-16 h-16 sm:w-20 sm:h-20 lg:w-22 lg:h-22 xl:w-24 xl:h-24 transition-transform duration-300 group-hover:scale-105" />
                  </div>
                </div>

                {/* Bottom Section: Action Button on Bottom-Left */}
                <div className="mt-auto pt-3 sm:pt-4 z-10">
                  {isPrimary ? (
                    <span className="inline-flex items-center gap-2 bg-[#0066CC] hover:bg-[#0052A3] text-white text-xs sm:text-[13px] font-bold px-4 py-1.5 sm:py-2 rounded-full shadow-sm group-hover:shadow-md transition-all">
                      <span>{card.actionLabel}</span>
                      <svg
                        viewBox="0 0 24 24"
                        className="w-4 h-4 transform group-hover:translate-x-1 transition-transform"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <line x1="5" y1="12" x2="19" y2="12" />
                        <polyline points="12 5 19 12 12 19" />
                      </svg>
                    </span>
                  ) : (
                    <span className="inline-flex items-center gap-1.5 sm:gap-2 bg-white hover:bg-blue-50/80 border border-blue-200/90 text-[#0066CC] text-xs sm:text-[13px] font-bold px-3.5 sm:px-4 py-1.5 rounded-full shadow-2xs group-hover:border-blue-400 group-hover:shadow-xs transition-all">
                      <span>{card.actionLabel}</span>
                      <svg
                        viewBox="0 0 24 24"
                        className="w-3.5 h-3.5 sm:w-4 sm:h-4 transform group-hover:translate-x-1 transition-transform"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <line x1="5" y1="12" x2="19" y2="12" />
                        <polyline points="12 5 19 12 12 19" />
                      </svg>
                    </span>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* 4. Subtle Corporate Footer Line Anchored at Bottom (Zero Bottom Gap) */}
      <footer
        style={{ backgroundColor: "#F8FAFC", borderTop: "1px solid #E2E8F0" }}
        className="py-2.5 sm:py-3 px-4 sm:px-6 lg:px-8 xl:px-10 2xl:px-12 text-center select-none shrink-0 mt-auto"
      >
        <div className="w-full flex flex-col sm:flex-row items-center justify-between text-[11px] sm:text-xs text-slate-500 font-medium">
          <span>© Vedanta Resources Ltd. • Save Food Management System</span>
          <span className="hidden sm:inline">Transforming For Good • Inform Early. Prepare Smart. Waste Less.</span>
        </div>
      </footer>
    </div>
  );
};

export default SaveFood;
