# VOGL SaveFood Management System

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

## 📌 Executive Overview & Objectives

In large industrial plant operations, unforeseen shifts, official travel, and employee leaves frequently cause food over-preparation and significant cafeteria wastage. The **VOGL SaveFood Management System** solves this by establishing an intuitive, advance notice mechanism:

1. **Advance Headcount Forecasting**: Enables employees to easily opt out of upcoming meals (Breakfast, Lunch, Dinner, Midnight snacks) ahead of scheduled cut-off windows.
2. **Catering Quantity Optimization**: Provides kitchen managers and catering contractors with real-time, consolidated headcount rosters to prepare the exact quantity of food needed.
3. **ESG & Sustainability Metrics**: Captures and visualizes kilograms of food saved, greenhouse gas reduction, and waste prevention statistics.

---

## 🛠️ Architecture & Tech Stack

- **SPFx Engine**: SharePoint Framework v1.24.0 (Beta) targeting modern SharePoint Online.
- **Frontend Framework**: React 18.3.1 with React DOM Client (`createRoot`).
- **Styling Architecture**: Tailwind CSS v3 compiled into isolated CSS, combined with custom SharePoint Modern Canvas full-bleed breakout engine (`.savefood-canvas-breakout` + recursive parent DOM margin/padding neutralizer).
- **Icons & Graphics**: Pure inline SVG vector illustrations tailored specifically for enterprise dining, cloche domes, kitchen consoles, and sustainability analytics.
- **Build System**: Microsoft Rush Stack Heft with TypeScript 5.8+ and Webpack.

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
