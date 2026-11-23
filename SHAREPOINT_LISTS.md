# SharePoint Lists & Columns Specification

This document provides the exact list names, internal column names, data types, and configurations for the **SaveFood Management System**.

---

## 1. `LocationMaster` (Canteen / Dining Facilities Master)

Master list of active dining locations and mess facilities across plant sites.

- **List Name / Title**: `LocationMaster`
- **Description**: Stores active plant canteen facilities for request dropdowns.

### Columns Table

| Display Name | Internal Name | Field Type | Required | Default / Configuration Notes |
| :--- | :--- | :--- | :---: | :--- |
| **Location** | `Title` | Single line of text | **Yes** | Unique location identifier (e.g., `LQ`, `MCR`). |
| **Is Active** | `IsActive` | Yes/No (Boolean) | **Yes** | Default: `Yes` (`true`). Only active facilities appear in opt-out dropdowns. |

### Seed Data

| ID | Title (Location) | IsActive |
| :---: | :--- | :---: |
| 1 | `LQ` | Yes |
| 2 | `MCR` | Yes |
| 3 | `Project Office` | Yes |
| 4 | `CCR` | Yes |
| 5 | `CPF` | Yes |
| 6 | `Aish` | Yes |

---

## 2. `RoleMatrix` (Role Access Configuration)

Defines authorized Administrator accounts. Maintained by Site Owners.

- **List Name / Title**: `RoleMatrix`
- **Description**: Identifies Admin users. If missing or user is not listed, the user is treated as an Employee.

### Columns Table

| Display Name | Internal Name | Field Type | Required | Configuration Notes |
| :--- | :--- | :--- | :---: | :--- |
| **Title** | `Title` | Single line of text | **Yes** | Fixed value: `Application Roles`. |
| **Admin** | `Admin` | Person or Group | No | Enable **Allow multiple selections** (`AllowMultipleValues = true`). Add authorized Admin users here. |

### Configuration Item

- **Item Title**: `Application Roles`
- **Admin**: Add administrator person accounts (e.g., `Shiven Pandit`).

---

## 3. `MealOptOutRequest` (Transactional Requests)

Stores individual meal cancellation requests submitted by employees.

- **List Name / Title**: `MealOptOutRequest`
- **Description**: Transactional table recording advance meal cancellations.

### Columns Table

| Display Name | Internal Name | Field Type | Required | REST Key / Notes |
| :--- | :--- | :--- | :---: | :--- |
| **Request Summary** | `Title` | Single line of text | No | Renamed from default `Title`. Formatted as `[Location] - [MealType] - [MealDate]`. |
| **Location** | `Location` | Lookup *(or Text)* | **Yes** | Lookup to `LocationMaster` (`Title`) or text location name. In REST: `LocationId` for Lookup or `Location` for Text. |
| **Meal Type** | `MealType` | Choice | **Yes** | Fixed choices: `Breakfast`, `Lunch`, `Dinner`. In REST: `MealType`. |
| **Meal Date** | `MealDate` | Date and Time | **Yes** | Date Only format. In REST: `MealDate` (ISO format: `YYYY-MM-DDT00:00:00Z`). |
| **Requester** | `Requester` | Person or Group | **Yes** | Single user. In REST: `RequesterId` (SharePoint User ID). |

---

## Code Integration Mapping

| List | Operation | REST Endpoint | Payload Fields |
| :--- | :--- | :--- | :--- |
| `LocationMaster` | Read Active Locations | `GET /_api/web/lists/getbytitle('LocationMaster')/items?$filter=IsActive eq 1` | `Id`, `Title`, `IsActive` |
| `RoleMatrix` | Verify Admin Role | `GET /_api/web/lists/getbytitle('RoleMatrix')/items?$filter=Title eq 'Application Roles'&$select=Id,Title,Admin/Id,Admin/Title,Admin/EMail,Admin/UserName&$expand=Admin` | `Admin` user records matched by normalized Email / UPN |
| `MealOptOutRequest` | Submit Opt-Out | `POST /_api/web/lists/getbytitle('MealOptOutRequest')/items` | `Title`, `LocationId` / `Location`, `MealType`, `MealDate`, `RequesterId` |
| `MealOptOutRequest` | My Requests | `GET /_api/web/lists/getbytitle('MealOptOutRequest')/items?$filter=Requester/Id eq {UserId}` | `Id`, `Title`, `MealType`, `MealDate`, `Location`, `Requester`, `Created` |
| `MealOptOutRequest` | All Requests (Admin) | `GET /_api/web/lists/getbytitle('MealOptOutRequest')/items` | Full list with date/location filters and pagination |
