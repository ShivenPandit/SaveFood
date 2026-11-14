# VOGL Save Food Management System

<div align="center">

![Vedanta Logo](https://img.shields.io/badge/VEDANTA-Transforming%20For%20Good-004B87?style=for-the-badge)
![Initiative](https://img.shields.io/badge/INITIATIVE-Zero%20Food%20Wastage-009639?style=for-the-badge)
![Platform](https://img.shields.io/badge/PLATFORM-SharePoint%20SPFx-0078D4?style=for-the-badge)
![React](https://img.shields.io/badge/REACT-18.3-61DAFB?style=for-the-badge&logo=react&logoColor=black)
![TailwindCSS](https://img.shields.io/badge/TAILWIND-3.x-38B2AC?style=for-the-badge&logo=tailwind-css&logoColor=white)

### **Inform Early. Prepare Smart. Waste Less.**

A centralized enterprise digital platform engineered for **Vedanta Limited (VOGL)** to manage advance employee meal opt-out requests, streamline kitchen caterer headcount planning, eliminate food wastage, and champion ESG sustainability across plant mess facilities.

</div>

---

## 🌐 Portal Environments & Official Links

| Environment           | Portal Link                                                                                              | Purpose                                                           | Access Level                                    |
| :-------------------- | :------------------------------------------------------------------------------------------------------- | :---------------------------------------------------------------- | :---------------------------------------------- |
| **UAT (Staging)**     | [**https://vedanta.sharepoint.com/sites/SaveFoodUAT**](https://vedanta.sharepoint.com/sites/SaveFoodUAT) | Pre-production validation, cut-off policy testing, catering pilot | Testers, Administrators & Plant POCs            |
| **Production (Live)** | [**https://vedanta.sharepoint.com/sites/SaveFood**](https://vedanta.sharepoint.com/sites/SaveFood)       | Enterprise live operations across all VOGL plant mess locations   | All VOGL Employees, Contractors & Kitchen Staff |

---

## 📌 Executive Overview & Objectives

In large industrial plant operations, unforeseen shifts, official travel, and employee leaves frequently cause food over-preparation and significant cafeteria wastage. The **VOGL Save Food Management System** solves this by establishing an intuitive, advance notice mechanism:

1. **Advance Headcount Forecasting**: Enables employees to easily opt out of upcoming meals (Breakfast, Lunch, Dinner, Midnight snacks) ahead of scheduled cut-off windows.
2. **Catering Quantity Optimization**: Provides kitchen managers and catering contractors with real-time, consolidated headcount rosters to prepare the exact quantity of food needed.
3. **ESG & Sustainability Metrics**: Captures and visualizes kilograms of food saved, greenhouse gas reduction, and waste prevention statistics.

---

## 🍱 Key Functional Modules

The home view features a modern, high-density 3x2 responsive Bento Grid designed with custom food-specific vector iconography and responsive hover elevations:

```
┌──────────────────────────────────────┬──────────────────────────────────────┬──────────────────────────────────────┐
│ 🍽️ Create Meal Opt-Out Request       │ 📋 My Meal Opt-Out Requests          │ 👥 All Meal Opt-Out Requests          │
│ Submit advance cancellations for     │ Track personal opt-out history,      │ Kitchen consolidated headcount       │
│ upcoming shifts, leave, or travel.   │ attendance, and notice deadlines.    │ roster and location-wise mess counts.│
├──────────────────────────────────────┼──────────────────────────────────────┼──────────────────────────────────────┤
│ 📊 Reports & Dashboard               │ ⚙️ Admin Configuration               │ 📖 User Manual                       │
│ Real-time food savings metrics,      │ Configure plant canteen units,       │ Operating procedures, cut-off        │
│ waste reduction trends, ESG reports. │ cut-off windows, and role matrix.    │ policy timelines, and FAQ guide.     │
└──────────────────────────────────────┴──────────────────────────────────────┴──────────────────────────────────────┘
```

1. **Create Meal Opt-Out Request** _(Primary Action)_:
   - Intuitive advance date picker and meal-type selector.
   - Bulk date selection for approved leaves and shift rotations.
2. **My Meal Opt-Out Requests**:
   - Digital employee dining ledger.
   - Status tracking with cancellation reversal before cut-off deadlines.
3. **All Meal Opt-Out Requests**:
   - Caterer-facing operational headcount console.
   - Filterable by Plant Location, Mess Facility, Meal Shift, and Date.
4. **Reports & Dashboard**:
   - Analytical reporting displaying total meals saved, waste reduction trends, and catering planning accuracy.
5. **Admin Configuration**:
   - Mess and dining hall master data management.
   - Cut-off window rules engine (e.g., Breakfast cut-off at 08:00 PM prior day, Lunch at 09:00 AM same day, Dinner at 03:00 PM same day).
   - Administrative role and user access matrix.
6. **User Manual**:
   - Standard operating guidelines, cut-off policy summaries, and helpline contacts.

---

## 🛠️ Architecture & Tech Stack

- **SPFx Engine**: SharePoint Framework v1.24.0 (Beta) targeting modern SharePoint Online.
- **Frontend Framework**: React 18.3.1 with React DOM Client (`createRoot`).
- **Styling Architecture**: Tailwind CSS v3 compiled into isolated CSS, combined with custom SharePoint Modern Canvas full-bleed breakout engine (`.savefood-canvas-breakout` + recursive parent DOM margin/padding neutralizer).
- **Icons & Graphics**: Pure inline SVG vector illustrations tailored specifically for enterprise dining, cloche domes, kitchen consoles, and sustainability analytics.
- **Build System**: Microsoft Rush Stack Heft with TypeScript 5.8+ and Webpack.

---

## 📁 Repository Directory Structure

```
SaveFood/
├── .gemini/                           # AI Agent configuration & context logs
├── config/                            # SPFx & Heft build configurations
│   ├── config.json                    # Component bundling configuration
│   ├── package-solution.json          # Solution package (.sppkg) definition
│   └── serve.json                     # Local dev server endpoints (UAT & Prod)
├── sharepoint/
│   └── solution/
│       └── SaveFood.sppkg             # Production SharePoint deployment package
├── src/
│   ├── tailwind.generated.css         # Compiled Tailwind utility classes
│   └── webparts/
│       └── savefood/
│           ├── assets/                # Base64 brand assets & fallbacks
│           ├── components/            # React UI components
│           │   ├── Common/            # Reusable form & picker components
│           │   ├── HomeIllustrations.tsx # Bespoke food & canteen vector icons
│           │   ├── ISaveFoodProps.ts  # Web part property contracts
│           │   └── SaveFood.tsx       # Main dashboard & responsive layout
│           ├── config/
│           │   └── portalConfig.ts    # Official UAT & Production URLs
│           ├── hooks/                 # SharePoint REST API data hooks
│           ├── tailwind.css           # Tailwind source definitions
│           ├── SaveFoodWebPart.manifest.json # Web part manifest
│           └── SaveFoodWebPart.ts     # Web part lifecycle & DOM controller
├── package.json                       # Project manifest and scripts
├── tsconfig.json                      # TypeScript configuration
└── README.md                          # Project documentation
```

---

## 🚀 Developer Setup & Running Locally

### 1. Prerequisites

- **Node.js**: `v22.x` (LTS recommended)
- **Package Manager**: `npm`
- **Permissions**: Access to Vedanta SharePoint UAT / Production tenant.

### 2. Installation

```powershell
git clone https://github.com/YourOrg/SaveFood.git
cd SaveFood
npm install
```

### 3. Running the Local Dev Server

Start the local SPFx development server (port 4321):

```powershell
npm run start
```

### 4. Testing in SharePoint with Debug Manifests

Open either environment in Google Chrome or Microsoft Edge with the debug query parameters:

- **Test on UAT Portal**:
  ```
  https://vedanta.sharepoint.com/sites/SaveFoodUAT/SitePages/Home.aspx?debugManifestsFile=https://localhost:4321/temp/build/manifests.js&debug=true&noredir=true
  ```
- **Test on Production Portal**:
  ```
  https://vedanta.sharepoint.com/sites/SaveFood/SitePages/Home.aspx?debugManifestsFile=https://localhost:4321/temp/build/manifests.js&debug=true&noredir=true
  ```

When prompted by SharePoint, click **"Load debug scripts"**.

---

## 📦 Building and Packaging for Deployment

When ready to generate a production package for SharePoint App Catalog deployment:

```powershell
# Compiles Tailwind, compiles TypeScript, bundles client assets, and generates the .sppkg
npm run build
```

The resulting package will be generated at:

```
sharepoint/solution/SaveFood.sppkg
```
