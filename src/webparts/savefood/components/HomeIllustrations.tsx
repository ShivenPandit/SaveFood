import * as React from "react";

/**
 * Modern 3D-styled Vector Illustrations for VOGL Save Food Management System.
 * Matched precisely to the enterprise dining dashboard aesthetic.
 */

// 1. Create Meal Opt-Out Request (Blue Calendar + Green Cutlery Opt-Out Badge)
export const CreateOptOutIllustration: React.FC<{ className?: string }> = ({ className = "" }) => (
  <svg
    viewBox="0 0 110 110"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={`shrink-0 select-none ${className}`}
    aria-label="Create Opt-Out Calendar Icon"
  >
    <defs>
      <filter id="calShadow" x="-20%" y="-15%" width="140%" height="140%">
        <feDropShadow dx="0" dy="5" stdDeviation="5" floodColor="#0F385A" floodOpacity="0.18" />
      </filter>
      <filter id="badgeGlow" x="-25%" y="-25%" width="150%" height="150%">
        <feDropShadow dx="0" dy="4" stdDeviation="4" floodColor="#16A34A" floodOpacity="0.35" />
      </filter>
      <linearGradient id="calHeader" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#0284C7" />
        <stop offset="100%" stopColor="#0052CC" />
      </linearGradient>
      <linearGradient id="optOutGreen" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#34D399" />
        <stop offset="100%" stopColor="#10B981" />
      </linearGradient>
    </defs>

    {/* Ambient Soft Glow Behind Illustration */}
    <circle cx="55" cy="55" r="46" fill="#86EFAC" fillOpacity="0.22" />

    {/* Calendar Body */}
    <g filter="url(#calShadow)">
      {/* Calendar Base Plate / Page */}
      <rect x="22" y="24" width="58" height="56" rx="10" fill="#FFFFFF" />
      {/* Subtle border */}
      <rect x="22" y="24" width="58" height="56" rx="10" stroke="#CBD5E1" strokeWidth="1.2" />

      {/* Calendar Header Bar */}
      <path
        d="M22 34C22 28.4772 26.4772 24 32 24H70C75.5228 24 80 28.4772 80 34V40H22V34Z"
        fill="url(#calHeader)"
      />

      {/* White / Silver Binder Spiral Rings */}
      <rect x="34" y="18" width="5" height="11" rx="2.5" fill="#E2E8F0" />
      <rect x="35" y="19" width="3" height="9" rx="1.5" fill="#FFFFFF" />

      <rect x="63" y="18" width="5" height="11" rx="2.5" fill="#E2E8F0" />
      <rect x="64" y="19" width="3" height="9" rx="1.5" fill="#FFFFFF" />

      {/* Calendar Grid Date Dots / Blocks */}
      {/* Row 1 */}
      <rect x="30" y="47" width="8" height="6" rx="1.8" fill="#E2E8F0" />
      <rect x="42" y="47" width="8" height="6" rx="1.8" fill="#E2E8F0" />
      <rect x="54" y="47" width="8" height="6" rx="1.8" fill="#E2E8F0" />
      <rect x="66" y="47" width="8" height="6" rx="1.8" fill="#BAE6FD" />

      {/* Row 2 */}
      <rect x="30" y="57" width="8" height="6" rx="1.8" fill="#E2E8F0" />
      <rect x="42" y="57" width="8" height="6" rx="1.8" fill="#E2E8F0" />
      <rect x="54" y="57" width="8" height="6" rx="1.8" fill="#E2E8F0" />

      {/* Row 3 */}
      <rect x="30" y="67" width="8" height="6" rx="1.8" fill="#E2E8F0" />
      <rect x="42" y="67" width="8" height="6" rx="1.8" fill="#E2E8F0" />
    </g>

    {/* Green Opt-Out Badge (Crossed Cutlery Dining Opt-Out) */}
    <g filter="url(#badgeGlow)">
      <circle cx="76" cy="72" r="17" fill="url(#optOutGreen)" stroke="#FFFFFF" strokeWidth="3.2" />

      {/* Inner Cutlery Symbols */}
      {/* Fork (Left angle) */}
      <path
        d="M69 66L79 78"
        stroke="#FFFFFF"
        strokeWidth="2.4"
        strokeLinecap="round"
      />
      {/* Fork tines */}
      <path
        d="M68 64L71 67M70.5 62.5L73.5 65.5"
        stroke="#FFFFFF"
        strokeWidth="1.8"
        strokeLinecap="round"
      />

      {/* Knife (Right angle) */}
      <path
        d="M79 66L69 78"
        stroke="#FFFFFF"
        strokeWidth="2.4"
        strokeLinecap="round"
      />
      {/* Knife Blade curved tip */}
      <path
        d="M80 64C78.5 62.5 76 64 76 64"
        stroke="#FFFFFF"
        strokeWidth="2"
        strokeLinecap="round"
      />
    </g>
  </svg>
);

// 2. My Meal Opt-Out Requests (Document Sheet + Clock Badge)
export const MyRequestsIllustration: React.FC<{ className?: string }> = ({ className = "" }) => (
  <svg
    viewBox="0 0 110 110"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={`shrink-0 select-none ${className}`}
    aria-label="My Requests Document Icon"
  >
    <defs>
      <filter id="docShadow" x="-20%" y="-15%" width="140%" height="140%">
        <feDropShadow dx="0" dy="5" stdDeviation="5" floodColor="#0F385A" floodOpacity="0.16" />
      </filter>
      <filter id="clockShadow" x="-25%" y="-25%" width="150%" height="150%">
        <feDropShadow dx="0" dy="4" stdDeviation="4" floodColor="#0052CC" floodOpacity="0.32" />
      </filter>
      <linearGradient id="docGrad" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#FFFFFF" />
        <stop offset="100%" stopColor="#F0F7FF" />
      </linearGradient>
      <linearGradient id="clockGrad" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#0284C7" />
        <stop offset="100%" stopColor="#0066CC" />
      </linearGradient>
    </defs>

    {/* Ambient Soft Glow Behind Illustration */}
    <circle cx="55" cy="55" r="46" fill="#BAE6FD" fillOpacity="0.3" />

    {/* Document Body */}
    <g filter="url(#docShadow)">
      {/* Document Sheet */}
      <rect x="22" y="16" width="56" height="70" rx="9" fill="url(#docGrad)" stroke="#BAE6FD" strokeWidth="1.5" />

      {/* Folded Top-Right Corner */}
      <path
        d="M66 16V26C66 27.6569 67.3431 29 69 29H78"
        stroke="#93C5FD"
        strokeWidth="1.5"
        strokeLinecap="round"
      />
      <path
        d="M66 16L78 29H69C67.3431 29 66 27.6569 66 26V16Z"
        fill="#E0F2FE"
      />

      {/* Document Header Line (Blue) */}
      <rect x="30" y="26" width="28" height="5" rx="2.5" fill="#0066CC" />

      {/* Document Content Text Lines */}
      <rect x="30" y="38" width="40" height="4" rx="2" fill="#38BDF8" />
      <rect x="30" y="47" width="34" height="3.5" rx="1.75" fill="#94A3B8" />
      <rect x="30" y="55" width="26" height="3.5" rx="1.75" fill="#CBD5E1" />
      <rect x="30" y="63" width="22" height="3.5" rx="1.75" fill="#CBD5E1" />
    </g>

    {/* Blue Clock Badge (Cut-off Window & Schedule) */}
    <g filter="url(#clockShadow)">
      <circle cx="76" cy="74" r="16" fill="url(#clockGrad)" stroke="#FFFFFF" strokeWidth="3" />

      {/* Clock Face Details */}
      <circle cx="76" cy="74" r="2" fill="#FFFFFF" />
      {/* Hour Hand (pointing towards 10 o'clock) */}
      <line x1="76" y1="74" x2="71" y2="69" stroke="#FFFFFF" strokeWidth="2.4" strokeLinecap="round" />
      {/* Minute Hand (pointing towards 2 o'clock) */}
      <line x1="76" y1="74" x2="82" y2="70" stroke="#FFFFFF" strokeWidth="2" strokeLinecap="round" />
    </g>
  </svg>
);

// 3. All Meal Opt-Out Requests (3 User Team Avatars + Green Checkmark Badge)
export const AllRequestsIllustration: React.FC<{ className?: string }> = ({ className = "" }) => (
  <svg
    viewBox="0 0 110 110"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={`shrink-0 select-none ${className}`}
    aria-label="All Requests Team Icon"
  >
    <defs>
      <filter id="teamShadow" x="-20%" y="-15%" width="140%" height="140%">
        <feDropShadow dx="0" dy="5" stdDeviation="5" floodColor="#0F385A" floodOpacity="0.18" />
      </filter>
      <filter id="checkGlow" x="-25%" y="-25%" width="150%" height="150%">
        <feDropShadow dx="0" dy="4" stdDeviation="4" floodColor="#16A34A" floodOpacity="0.35" />
      </filter>
      <linearGradient id="userPrimary" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#0284C7" />
        <stop offset="100%" stopColor="#0052CC" />
      </linearGradient>
      <linearGradient id="userSecondary" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#7DD3FC" />
        <stop offset="100%" stopColor="#38BDF8" />
      </linearGradient>
    </defs>

    {/* Ambient Soft Glow Behind Illustration */}
    <circle cx="55" cy="55" r="46" fill="#BAE6FD" fillOpacity="0.3" />

    {/* 3 Users Group */}
    <g filter="url(#teamShadow)">
      {/* Left User (Secondary Sky-Blue) */}
      <circle cx="32" cy="40" r="9.5" fill="url(#userSecondary)" />
      <path
        d="M17 65C17 56.5 24 53 32 53C40 53 47 56.5 47 65V67H17V65Z"
        fill="url(#userSecondary)"
      />

      {/* Right User (Secondary Sky-Blue) */}
      <circle cx="76" cy="40" r="9.5" fill="url(#userSecondary)" />
      <path
        d="M61 65C61 56.5 68 53 76 53C84 53 91 56.5 91 65V67H61V65Z"
        fill="url(#userSecondary)"
      />

      {/* Center Front User (Prominent Primary Royal Blue) */}
      <circle cx="54" cy="34" r="12" fill="url(#userPrimary)" stroke="#FFFFFF" strokeWidth="2.2" />
      <path
        d="M34 68C34 57.5 43 51 54 51C65 51 74 57.5 74 68V72H34V68Z"
        fill="url(#userPrimary)"
        stroke="#FFFFFF"
        strokeWidth="2.2"
      />
    </g>

    {/* Green Checkmark Badge (Approved Headcounts / Kitchen Verified) */}
    <g filter="url(#checkGlow)">
      <circle cx="80" cy="74" r="14.5" fill="#22C55E" stroke="#FFFFFF" strokeWidth="3" />
      <path
        d="M74 74L78 78L86 70"
        stroke="#FFFFFF"
        strokeWidth="2.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </g>
  </svg>
);

// 4. Reports & Dashboard (Rising 3D Bar Chart + Upward Trend Arrow)
export const ReportsIllustration: React.FC<{ className?: string }> = ({ className = "" }) => (
  <svg
    viewBox="0 0 110 110"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={`shrink-0 select-none ${className}`}
    aria-label="Reports & Dashboard Analytics Icon"
  >
    <defs>
      <filter id="barsShadow" x="-20%" y="-15%" width="140%" height="140%">
        <feDropShadow dx="0" dy="5" stdDeviation="5" floodColor="#0F385A" floodOpacity="0.16" />
      </filter>
      <filter id="arrowShadow" x="-20%" y="-20%" width="140%" height="140%">
        <feDropShadow dx="0" dy="3.5" stdDeviation="3.5" floodColor="#0052CC" floodOpacity="0.32" />
      </filter>
      <linearGradient id="barSky" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stopColor="#38BDF8" />
        <stop offset="100%" stopColor="#0284C7" />
      </linearGradient>
      <linearGradient id="barBlue" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stopColor="#2563EB" />
        <stop offset="100%" stopColor="#0052CC" />
      </linearGradient>
      <linearGradient id="barGreen" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stopColor="#4ADE80" />
        <stop offset="100%" stopColor="#16A34A" />
      </linearGradient>
      <linearGradient id="trendArrowGrad" x1="0" y1="1" x2="1" y2="0">
        <stop offset="0%" stopColor="#38BDF8" />
        <stop offset="100%" stopColor="#0052CC" />
      </linearGradient>
    </defs>

    {/* Ambient Soft Glow Behind Illustration */}
    <circle cx="55" cy="55" r="46" fill="#BAE6FD" fillOpacity="0.3" />

    {/* 3 Rising Bar Chart Columns */}
    <g filter="url(#barsShadow)">
      {/* Baseline Shadow */}
      <rect x="22" y="80" width="66" height="3.5" rx="1.75" fill="#E2E8F0" />

      {/* Bar 1 (Sky Blue) */}
      <rect x="28" y="56" width="16" height="24" rx="4.5" fill="url(#barSky)" />
      {/* Top highlight */}
      <rect x="30" y="57" width="12" height="3" rx="1.5" fill="#FFFFFF" fillOpacity="0.35" />

      {/* Bar 2 (Royal Blue) */}
      <rect x="48" y="42" width="16" height="38" rx="4.5" fill="url(#barBlue)" />
      {/* Top highlight */}
      <rect x="50" y="43" width="12" height="3" rx="1.5" fill="#FFFFFF" fillOpacity="0.35" />

      {/* Bar 3 (Emerald Green) */}
      <rect x="68" y="28" width="16" height="52" rx="4.5" fill="url(#barGreen)" />
      {/* Top highlight */}
      <rect x="70" y="29" width="12" height="3" rx="1.5" fill="#FFFFFF" fillOpacity="0.35" />
    </g>

    {/* Upward Growth Trend Arrow */}
    <g filter="url(#arrowShadow)">
      {/* Curved Line */}
      <path
        d="M26 48C40 46 54 30 78 20"
        stroke="url(#trendArrowGrad)"
        strokeWidth="4"
        strokeLinecap="round"
      />
      {/* Arrow Head */}
      <path
        d="M68 18L81 19L79 32"
        stroke="url(#trendArrowGrad)"
        strokeWidth="4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </g>
  </svg>
);

// 5. Admin Configuration (Two Interlocking Gear Cogs: Blue & Green)
export const AdminConfigIllustration: React.FC<{ className?: string }> = ({ className = "" }) => (
  <svg
    viewBox="0 0 110 110"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={`shrink-0 select-none ${className}`}
    aria-label="Admin Configuration Gears Icon"
  >
    <defs>
      <filter id="gearShadow" x="-20%" y="-15%" width="140%" height="140%">
        <feDropShadow dx="0" dy="5" stdDeviation="5" floodColor="#0F385A" floodOpacity="0.18" />
      </filter>
      <filter id="greenGearShadow" x="-25%" y="-25%" width="150%" height="150%">
        <feDropShadow dx="0" dy="3.5" stdDeviation="3.5" floodColor="#16A34A" floodOpacity="0.32" />
      </filter>
      <linearGradient id="blueGearGrad" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#0284C7" />
        <stop offset="100%" stopColor="#0052CC" />
      </linearGradient>
      <linearGradient id="greenGearGrad" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#4ADE80" />
        <stop offset="100%" stopColor="#16A34A" />
      </linearGradient>
    </defs>

    {/* Ambient Soft Glow Behind Illustration */}
    <circle cx="55" cy="55" r="46" fill="#BAE6FD" fillOpacity="0.3" />

    {/* Large Primary Gear Cog (Royal Blue) */}
    <g filter="url(#gearShadow)">
      {/* 8 Gear Teeth */}
      {/* 0 deg & 180 deg */}
      <rect x="42" y="16" width="14" height="54" rx="4" fill="url(#blueGearGrad)" />
      {/* 90 deg & 270 deg */}
      <rect x="22" y="36" width="54" height="14" rx="4" fill="url(#blueGearGrad)" />
      {/* 45 deg & 225 deg */}
      <rect
        x="42"
        y="16"
        width="14"
        height="54"
        rx="4"
        fill="url(#blueGearGrad)"
        transform="rotate(45 49 43)"
      />
      {/* 135 deg & 315 deg */}
      <rect
        x="42"
        y="16"
        width="14"
        height="54"
        rx="4"
        fill="url(#blueGearGrad)"
        transform="rotate(-45 49 43)"
      />

      {/* Main Center Wheel */}
      <circle cx="49" cy="43" r="22" fill="url(#blueGearGrad)" />
      {/* Inner Cutout Hole */}
      <circle cx="49" cy="43" r="9.5" fill="#FFFFFF" />
      {/* Inner Hub Shadow Ring */}
      <circle cx="49" cy="43" r="9.5" stroke="#BAE6FD" strokeWidth="1.5" />
    </g>

    {/* Secondary Meshing Gear Cog (Vibrant Green) */}
    <g filter="url(#greenGearShadow)">
      {/* Green Teeth */}
      <rect x="74" y="58" width="10" height="32" rx="3" fill="url(#greenGearGrad)" />
      <rect x="63" y="69" width="32" height="10" rx="3" fill="url(#greenGearGrad)" />
      <rect
        x="74"
        y="58"
        width="10"
        height="32"
        rx="3"
        fill="url(#greenGearGrad)"
        transform="rotate(45 79 74)"
      />
      <rect
        x="74"
        y="58"
        width="10"
        height="32"
        rx="3"
        fill="url(#greenGearGrad)"
        transform="rotate(-45 79 74)"
      />

      {/* Green Wheel Body */}
      <circle cx="79" cy="74" r="13" fill="url(#greenGearGrad)" stroke="#FFFFFF" strokeWidth="2.2" />
      {/* Inner Cutout Hole */}
      <circle cx="79" cy="74" r="5" fill="#FFFFFF" />
    </g>
  </svg>
);

// 6. User Manual (Open Reference Handbook)
export const UserManualIllustration: React.FC<{ className?: string }> = ({ className = "" }) => (
  <svg
    viewBox="0 0 110 110"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={`shrink-0 select-none ${className}`}
    aria-label="User Manual Book Icon"
  >
    <defs>
      <filter id="bookShadow" x="-20%" y="-15%" width="140%" height="140%">
        <feDropShadow dx="0" dy="5" stdDeviation="5" floodColor="#0F385A" floodOpacity="0.16" />
      </filter>
      <linearGradient id="bookCoverGrad" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#0284C7" />
        <stop offset="100%" stopColor="#0052CC" />
      </linearGradient>
    </defs>

    {/* Ambient Soft Glow Behind Illustration */}
    <circle cx="55" cy="55" r="46" fill="#BAE6FD" fillOpacity="0.3" />

    {/* Open Book */}
    <g filter="url(#bookShadow)">
      {/* Outer Blue Hardcover */}
      <path
        d="M20 74C30 71 45 72 55 76C65 72 80 71 90 74V34C80 31 65 32 55 36C45 32 30 31 20 34V74Z"
        fill="url(#bookCoverGrad)"
      />

      {/* Left White Page */}
      <path
        d="M23 70C32 67.5 45 68.5 54 72V33C45 29.5 32 28.5 23 31V70Z"
        fill="#FFFFFF"
        stroke="#E0F2FE"
        strokeWidth="1.2"
      />

      {/* Left Page Text Lines */}
      <line x1="28" y1="38" x2="48" y2="38" stroke="#38BDF8" strokeWidth="2.4" strokeLinecap="round" />
      <line x1="28" y1="46" x2="48" y2="46" stroke="#94A3B8" strokeWidth="2" strokeLinecap="round" />
      <line x1="28" y1="53" x2="44" y2="53" stroke="#CBD5E1" strokeWidth="2" strokeLinecap="round" />
      <line x1="28" y1="60" x2="46" y2="60" stroke="#CBD5E1" strokeWidth="2" strokeLinecap="round" />

      {/* Right White Page */}
      <path
        d="M87 70C78 67.5 65 68.5 56 72V33C65 29.5 78 28.5 87 31V70Z"
        fill="#F8FAFC"
        stroke="#E0F2FE"
        strokeWidth="1.2"
      />

      {/* Right Page Text Lines */}
      <line x1="62" y1="38" x2="82" y2="38" stroke="#38BDF8" strokeWidth="2.4" strokeLinecap="round" />
      <line x1="62" y1="46" x2="82" y2="46" stroke="#94A3B8" strokeWidth="2" strokeLinecap="round" />
      <line x1="62" y1="53" x2="76" y2="53" stroke="#CBD5E1" strokeWidth="2" strokeLinecap="round" />
      <line x1="62" y1="60" x2="80" y2="60" stroke="#CBD5E1" strokeWidth="2" strokeLinecap="round" />

      {/* Center Spine Crease */}
      <line x1="55" y1="32" x2="55" y2="74" stroke="#004B87" strokeWidth="2.4" />

      {/* Bookmark Ribbon */}
      <path
        d="M53 32V46L56 43L59 46V32H53Z"
        fill="#22C55E"
      />
    </g>
  </svg>
);
