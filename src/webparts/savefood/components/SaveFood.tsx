import * as React from "react";
import { ISaveFoodProps } from "./ISaveFoodProps";
import { VEDANTA_LOGO_BASE64 } from "../assets/logoBase64";
import {
  CreateOptOutIllustration,
  MyRequestsIllustration,
  ReportsIllustration,
  UserManualIllustration,
} from "./HomeIllustrations";
import { resolveUserRole } from "../services/roleService";
import { CreateMealOptOut } from "./CreateRequest/CreateMealOptOut";
import { MyMealOptOuts } from "./MyRequests/MyMealOptOuts";
import { ReportsDashboard } from "./Reports/ReportsDashboard";
import { UserManual } from "./UserManual/UserManual";

export type ActiveView =
  | "home"
  | "create"
  | "my-requests"
  | "all-requests"
  | "reports"
  | "manual";

interface IHomeActionCard {
  id: string;
  title: string;
  description: string;
  actionLabel: string;
  isPrimary?: boolean;
  adminOnly?: boolean;
  view: ActiveView;
  illustration: React.FC<{ className?: string }>;
}

// Helper: Parse active view from current browser URL hash
const getViewFromHash = (): ActiveView => {
  if (typeof window === "undefined") return "home";
  const hash = (window.location.hash || "").replace(/^#\/?/, "").toLowerCase().trim();
  if (hash === "create") return "create";
  if (hash === "my-requests" || hash === "my") return "my-requests";
  if (hash === "all-requests" || hash === "all") return "all-requests";
  if (hash === "reports") return "reports";
  if (hash === "manual") return "manual";
  return "home";
};

const SaveFood: React.FC<ISaveFoodProps> = (props) => {
  const siteAssetLogoUrl = `${props.newsiteUrl}/SiteAssets/Vedanta1.png`;
  const [logoSrc, setLogoSrc] = React.useState<string>(siteAssetLogoUrl);

  // Active View Navigation State (synced with browser URL hash)
  const [currentView, setCurrentView] = React.useState<ActiveView>(getViewFromHash);

  // Role Resolution State (Default: Employee)
  const [isAdmin, setIsAdmin] = React.useState<boolean>(false);
  const [isResolvingRole, setIsResolvingRole] = React.useState<boolean>(true);

  // 1. Resolve User Role against RoleMatrix list
  React.useEffect(() => {
    let isMounted = true;
    resolveUserRole(
      props.newspHttpClient,
      props.newsiteUrl,
      props.newcurrentUserEmail,
      props.newcurrentUserLoginName
    )
      .then((res) => {
        if (isMounted) {
          setIsAdmin(res.isAdmin);
          setIsResolvingRole(false);
        }
      })
      .catch((err) => {
        console.warn("[SaveFood] Role lookup failed, fallback to Employee:", err);
        if (isMounted) {
          setIsAdmin(false);
          setIsResolvingRole(false);
        }
      });

    return () => {
      isMounted = false;
    };
  }, [props.newspHttpClient, props.newsiteUrl, props.newcurrentUserEmail, props.newcurrentUserLoginName]);

  // 2. Hash Router Listener with Route Guard
  React.useEffect(() => {
    const onHashChange = (): void => {
      const newView = getViewFromHash();
      // Route guard: if trying to access admin view without admin role, redirect to home
      if (["all-requests", "reports"].indexOf(newView) !== -1 && !isAdmin) {
        if (window.location.hash !== "#/" && window.location.hash !== "") {
          window.location.hash = "#/";
        }
        setCurrentView("home");
      } else {
        setCurrentView(newView);
      }
    };

    window.addEventListener("hashchange", onHashChange);
    // Sync view immediately
    onHashChange();

    return () => {
      window.removeEventListener("hashchange", onHashChange);
    };
  }, [isAdmin]);

  // Lock dashboard scrollbar when on the dashboard to ensure the entire page fits cleanly
  React.useEffect(() => {
    if (currentView === "home") {
      document.body.classList.add("savefood-dashboard-no-scroll");
      document.documentElement.classList.add("savefood-dashboard-no-scroll");
    } else {
      document.body.classList.remove("savefood-dashboard-no-scroll");
      document.documentElement.classList.remove("savefood-dashboard-no-scroll");
    }
    return () => {
      document.body.classList.remove("savefood-dashboard-no-scroll");
      document.documentElement.classList.remove("savefood-dashboard-no-scroll");
    };
  }, [currentView]);

  const handleLogoError = (): void => {
    if (logoSrc !== VEDANTA_LOGO_BASE64) {
      setLogoSrc(VEDANTA_LOGO_BASE64);
    }
  };

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

  // Safe navigation that updates the browser URL hash for every view
  const handleNavigate = (view: ActiveView): void => {
    if (["all-requests", "reports"].indexOf(view) !== -1 && !isAdmin) {
      console.warn("[SaveFood] Route guard blocked unauthorized view navigation:", view);
      return;
    }
    const targetHash = view === "home" ? "#/" : `#/${view}`;
    if (window.location.hash !== targetHash) {
      window.location.hash = targetHash;
    }
    setCurrentView(view);
  };

  const cleanName = (props.newcurrentUserName || "User").replace(/\s*\([^)]*\)/g, "").trim();
  const initials = cleanName
    .split(" ")
    .map((n) => n[0])
    .join("")
    .substring(0, 2)
    .toUpperCase();

  // 6 Action Cards definition
  const actionCards: IHomeActionCard[] = [
    {
      id: "create-optout",
      title: "Create Meal Opt-Out Request",
      description:
        "Submit advance meal cancellations for upcoming shifts, leave, or plant travel.",
      actionLabel: "Create Opt-Out",
      isPrimary: true,
      adminOnly: false,
      view: "create",
      illustration: CreateOptOutIllustration,
    },
    {
      id: "my-optouts",
      title: "My Meal Opt-Out Requests",
      description:
        "Track submitted meal cancellations and monitor advance notice status.",
      actionLabel: "View Requests",
      isPrimary: false,
      adminOnly: false,
      view: "my-requests",
      illustration: MyRequestsIllustration,
    },
    {
      id: "reports",
      title: "Reports & Dashboard",
      description:
        "Real-time food savings metrics, waste reduction trends, and caterer accuracy.",
      actionLabel: "View Reports",
      isPrimary: false,
      adminOnly: true,
      view: "reports",
      illustration: ReportsIllustration,
    },
    {
      id: "manual",
      title: "User Manual",
      description:
        "View process guidelines, meal booking cut-off policies, FAQs, and support.",
      actionLabel: "View Guide",
      isPrimary: false,
      adminOnly: false,
      view: "manual",
      illustration: UserManualIllustration,
    },
  ];

  // Non-admins (Employees) see Employee accessible cards (3 cards). Admins see all 5 cards.
  const visibleCards = actionCards.filter((c) => !c.adminOnly || isAdmin);

  return (
    <div
      style={{
        minHeight: "100vh",
        backgroundColor: "#F8FAFC",
      }}
      className="savefood-canvas-breakout w-full max-w-none flex flex-col min-h-screen text-slate-800 antialiased font-sans select-none box-border"
    >
      {/* 1. Top Navbar: Vedanta Logo, Title, Role Pill & User Profile */}
      <header className="h-12 sm:h-13 lg:h-14 bg-white border-b border-slate-200/90 shrink-0 shadow-2xs z-30 sticky top-0 w-full select-none px-6 lg:px-8 flex items-center justify-between">
        {/* Left Side: Brand Logo + System Identity */}
        <div className="flex items-center gap-2.5 sm:gap-3.5 min-w-0">
          <img
            src={logoSrc}
            alt="Vedanta Logo"
            className="h-7 sm:h-7.5 lg:h-8 w-auto object-contain transition-transform cursor-pointer"
            onClick={() => handleNavigate("home")}
            onError={handleLogoError}
          />
          <div className="h-5 w-px bg-slate-200 hidden xs:block" />
          <div className="flex items-center gap-2 min-w-0">
            <span
              onClick={() => handleNavigate("home")}
              className="text-xs sm:text-sm lg:text-base font-extrabold text-[#003865] tracking-tight truncate cursor-pointer hover:text-[#004B87] transition-colors"
            >
              Save Food Management System
            </span>
          </div>
        </div>

        {/* Right Side: Role Badge, Date & User Profile */}
        <div className="flex items-center gap-2 sm:gap-3">
          {/* User Role Indicator Pill */}
          {!isResolvingRole && (
            <span
              className={`px-2.5 py-0.5 rounded-full text-[10.5px] font-extrabold tracking-wide uppercase ${
                isAdmin
                  ? "bg-[#004B87] text-white shadow-2xs"
                  : "bg-slate-100 text-slate-700 border border-slate-200"
              }`}
            >
              {isAdmin ? "Admin" : "Employee"}
            </span>
          )}

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
                {isAdmin ? "Administrator" : "Employee"}
              </p>
            </div>
          </div>
        </div>
      </header>

      {/* Signature Green Accent Line */}
      <div className="w-full h-[2px] bg-[#009639] shrink-0" />

      {/* 2. Main Content Router */}
      {currentView === "home" && (
        <main
          style={{
            display: "flex",
            flexDirection: "column",
            justifyContent: "flex-start",
            width: "100%",
            flex: "1 0 auto",
            boxSizing: "border-box",
          }}
          className="w-full box-border"
        >
          {/* Edge-to-Edge Full-Bleed Hero Banner */}
          <section
            style={{
              flex: "0 0 auto",
              minHeight: "175px",
              padding: "24px 36px",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              width: "100%",
              margin: 0,
              borderRadius: 0,
              boxSizing: "border-box",
            }}
            className="relative w-full overflow-hidden bg-gradient-to-r from-[#002D56] via-[#004B87] to-[#046A38] text-white select-none shadow-md shrink-0 rounded-none"
          >
            <div className="absolute inset-0 bg-gradient-to-r from-[#002D56] via-[#004B87]/90 via-55% to-transparent pointer-events-none z-10" />

            <div className="absolute inset-y-0 left-1/3 md:left-2/5 lg:left-1/2 w-96 pointer-events-none z-10 overflow-hidden hidden sm:block">
              <div
                className="absolute -top-36 bottom-[-140px] w-12 lg:w-14 bg-sky-400/25 transform -skew-x-[26deg] backdrop-blur-[1px]"
                style={{ left: "10%" }}
              />
              <div
                className="absolute -top-36 bottom-[-140px] w-4 lg:w-5 bg-[#F59E0B]/80 transform -skew-x-[26deg] shadow-lg"
                style={{ left: "26%" }}
              />
              <div
                className="absolute -top-36 bottom-[-140px] w-20 lg:w-24 bg-[#1E3A8A]/45 transform -skew-x-[26deg]"
                style={{ left: "34%" }}
              />
            </div>

            <div className="relative z-20 flex flex-col md:flex-row md:items-center justify-between gap-4 w-full min-w-0">
              <div className="max-w-2xl xl:max-w-3xl min-w-0">
                <div className="w-12 h-1.5 bg-[#F59E0B] rounded-full mb-2" />
                <p className="text-white/90 text-sm sm:text-[14px] font-semibold tracking-wider uppercase">
                  Vedanta Sustainability Platform
                </p>
                <h1 className="font-extrabold text-white tracking-tight leading-tight text-2xl sm:text-3xl lg:text-[32px] mt-1">
                  VOGL Save Food{" "}
                  <span className="text-emerald-400 font-black whitespace-nowrap">
                    Management System
                  </span>
                </h1>
                <p className="text-amber-300 font-extrabold tracking-widest text-xs sm:text-sm mt-1 uppercase">
                  Inform Early. Prepare Smart. Waste Less.
                </p>
                <p className="text-blue-100/90 leading-relaxed text-xs sm:text-sm mt-2 max-w-xl lg:max-w-2xl 2xl:max-w-3xl">
                  A centralized digital platform to manage advance meal opt-out requests, optimize catering planning, minimize food wastage, and promote sustainability across VOGL locations.
                </p>
              </div>

              <div className="hidden lg:flex flex-col items-end text-right justify-center pr-2 shrink-0">
                <div
                  style={{ padding: "18px 24px" }}
                  className="bg-white/10 backdrop-blur-md border border-white/20 rounded-2xl max-w-[240px] text-center shadow-lg"
                >
                  <span className="text-[11px] font-bold tracking-wider text-emerald-300 uppercase block">
                    Responsible Living
                  </span>
                  <p className="text-xs sm:text-[13px] font-semibold text-white mt-1 leading-snug">
                    Every advance opt-out preserves food &amp; energy.
                  </p>
                  <div className="w-8 h-0.5 bg-[#F59E0B] rounded-full mx-auto mt-2" />
                </div>
              </div>
            </div>
          </section>

          {/* Card Grid: Padded Container Underneath Banner (2 Columns for both Admin & Employee) */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(2, 1fr)",
              gap: "20px",
              padding: "20px 32px 16px 32px",
              width: "100%",
              boxSizing: "border-box",
            }}
            className="w-full"
          >
            {visibleCards.map((card) => {
              const Illustration = card.illustration;
              const isPrimary = card.isPrimary;

              return (
                <div
                  key={card.id}
                  onClick={() => handleNavigate(card.view)}
                  style={{
                    minHeight: "190px",
                    maxHeight: "225px",
                    padding: "26px 30px",
                    display: "flex",
                    flexDirection: "column",
                    justifyContent: "space-between",
                    boxSizing: "border-box",
                  }}
                  className={`group relative rounded-2xl ${
                    isPrimary
                      ? "bg-gradient-to-br from-[#EBFBF0] via-[#E2F7E9] to-[#D1F2DD] border-2 border-[#86EFAC] hover:border-[#10B981] hover:shadow-emerald-900/15"
                      : "bg-white border-2 border-slate-200/90 hover:border-[#0066CC] hover:shadow-slate-300/40"
                  } shadow-xs hover:shadow-lg hover:-translate-y-0.5 transition-all duration-200 ease-out cursor-pointer overflow-hidden select-none min-w-0`}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" || e.key === " ") {
                      e.preventDefault();
                      handleNavigate(card.view);
                    }
                  }}
                >
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

                  <div className="flex items-start justify-between gap-4 z-10 min-w-0">
                    <div className="flex-1 min-w-0 pr-2">
                      <h3
                        className={`font-black text-slate-900 tracking-tight leading-snug ${
                          !isPrimary ? "group-hover:text-[#004B87]" : ""
                        } transition-colors text-[17px] sm:text-[18px]`}
                      >
                        {card.title}
                      </h3>
                      <p
                        className={`${
                          isPrimary ? "text-slate-600" : "text-slate-500"
                        } mt-2 text-[13px] sm:text-[13.5px] leading-relaxed line-clamp-2`}
                      >
                        {card.description}
                      </p>
                    </div>

                    <div className="shrink-0 flex items-center justify-end pointer-events-none select-none pl-1 self-start">
                      <Illustration className="w-14 h-14 sm:w-16 sm:h-16 transition-transform duration-300 group-hover:scale-105" />
                    </div>
                  </div>

                  <div className="mt-auto pt-3 z-10">
                    {isPrimary ? (
                      <span className="inline-flex items-center gap-2 bg-[#0066CC] hover:bg-[#0052A3] text-white text-xs sm:text-sm font-bold px-4 py-2 rounded-full shadow-2xs group-hover:shadow-xs transition-all">
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
                      <span className="inline-flex items-center gap-2 bg-white hover:bg-blue-50/80 border border-blue-200/90 text-[#0066CC] text-xs sm:text-sm font-bold px-4 py-2 rounded-full shadow-2xs group-hover:border-blue-400 group-hover:shadow-xs transition-all">
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
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </main>
      )}

      {/* Child Views */}
      {currentView === "create" && (
        <main
          style={{ flex: "1 0 auto", display: "flex", flexDirection: "column", padding: "16px 32px" }}
          className="w-full box-border"
        >
          <CreateMealOptOut
            context={props.newcontext}
            spHttpClient={props.newspHttpClient}
            siteUrl={props.newsiteUrl}
            currentUserId={props.newcurrentUserId}
            currentUserName={props.newcurrentUserName}
            currentUserEmail={props.newcurrentUserEmail}
            onSuccess={() => handleNavigate("my-requests")}
            onCancel={() => handleNavigate("home")}
          />
        </main>
      )}

      {currentView === "my-requests" && (
        <main
          style={{ flex: "1 0 auto", display: "flex", flexDirection: "column", padding: "16px 32px" }}
          className="w-full box-border"
        >
          <MyMealOptOuts
            context={props.newcontext}
            spHttpClient={props.newspHttpClient}
            siteUrl={props.newsiteUrl}
            currentUserId={props.newcurrentUserId}
            currentUserEmail={props.newcurrentUserEmail}
            onNavigateCreate={() => handleNavigate("create")}
            onBack={() => handleNavigate("home")}
          />
        </main>
      )}

      {currentView === "all-requests" && isAdmin && (
        <main
          style={{ flex: "1 0 auto", display: "flex", flexDirection: "column", padding: "16px 32px" }}
          className="w-full box-border"
        >
          <ReportsDashboard
            context={props.newcontext}
            spHttpClient={props.newspHttpClient}
            siteUrl={props.newsiteUrl}
            initialTab="requests"
            onBack={() => handleNavigate("home")}
          />
        </main>
      )}

      {currentView === "reports" && isAdmin && (
        <main
          style={{ flex: "1 0 auto", display: "flex", flexDirection: "column", padding: "16px 32px" }}
          className="w-full box-border"
        >
          <ReportsDashboard
            context={props.newcontext}
            spHttpClient={props.newspHttpClient}
            siteUrl={props.newsiteUrl}
            initialTab="dashboard"
            onBack={() => handleNavigate("home")}
          />
        </main>
      )}

      {currentView === "manual" && (
        <main
          style={{ flex: "1 0 auto", display: "flex", flexDirection: "column", padding: "16px 32px" }}
          className="w-full box-border"
        >
          <UserManual
            onBack={() => handleNavigate("home")}
            onNavigateCreate={() => handleNavigate("create")}
          />
        </main>
      )}

      {/* 4. Compact Single-Row Corporate Footer Pinned Flush to Bottom */}
      <footer
        style={{ backgroundColor: "#ffffff", borderTop: "1px solid #e2e8f0", padding: "8px 32px" }}
        className="select-none shrink-0 mt-auto w-full"
      >
        <div className="w-full flex flex-col sm:flex-row items-center justify-between text-[11px] sm:text-xs text-slate-500 font-medium gap-1.5 sm:gap-0">
          <span>© Vedanta Resources Ltd. • Save Food Management System</span>
          <div className="flex items-center gap-1.5 sm:gap-2 flex-wrap justify-center sm:justify-end text-[10.5px] sm:text-[11.5px]">
            <span className="font-extrabold text-[#003865]">Cut-Off Timings (IST):</span>
            <span className="inline-flex items-center gap-1 bg-amber-50 text-amber-800 border border-amber-200/90 px-2 py-0.5 rounded-full font-bold">
              <span>🍳</span> Breakfast: 08:00 PM (Prev Day)
            </span>
            <span className="inline-flex items-center gap-1 bg-blue-50 text-blue-800 border border-blue-200/90 px-2 py-0.5 rounded-full font-bold">
              <span>🥗</span> Lunch: 10:00 AM
            </span>
            <span className="inline-flex items-center gap-1 bg-purple-50 text-purple-800 border border-purple-200/90 px-2 py-0.5 rounded-full font-bold">
              <span>🍲</span> Dinner: 06:00 PM
            </span>
          </div>
        </div>
      </footer>
    </div>
  );
};

export default SaveFood;
