# Power Automate Setup Guide: Automated Daily Meal Cut-Off Summary Emails

This guide provides end-to-end instructions for creating automated cloud flows in **Microsoft Power Automate** that deliver daily meal cut-off notifications to all Administrators in the Vedanta Save Food Management System.

---

## 1. Schedule & Shift Matrix (India Standard Time - IST)

The system enforces three daily operational meal cut-offs. The flow triggers automatically at each cut-off deadline:

| Trigger Time (IST) | Meal Shift | Target Meal Date | Description |
|---|---|---|---|
| **10:00 AM IST** | **Lunch** | **Today** | Final headcount reduction for Today's Lunch. |
| **06:00 PM IST** (18:00) | **Dinner** | **Today** | Final headcount reduction for Today's Dinner. |
| **08:00 PM IST** (20:00) | **Breakfast** | **Tomorrow** *(Next Day)* | Advance cut-off for Tomorrow's Breakfast (previous day rule). |

---

## 2. Minimal Email Body Design

Per Vedanta Save Food requirements, the email is kept **short, minimal, and executive-friendly**:
- **Header**: Shift name, icon, and formatted date.
- **Location-Wise Summary Table**: 2-column table showing each canteen location and its opt-out count.
- **Total Headcount**: Concluding summary row labeled **`Total Headcount`**.
- **No Individual Employee Roster**: Detailed personal data is omitted to avoid clutter.
- **Direct Link Button**: One-click button to open the live SharePoint Report Dashboard.

### Email Layout Preview:
```html
+-----------------------------------------------------------------+
|  VEDANTA | Save Food Management System                          |
|  🥗 Lunch Opt-Out Summary — Friday, 20 Nov 2026                 |
+-----------------------------------------------------------------+
|                                                                 |
|  LOCATION-WISE BREAKDOWN                                        |
|  +-------------------------------+----------------------------+ |
|  | Canteen Location              | Opt-Out Headcount          | |
|  +-------------------------------+----------------------------+ |
|  | LQ                            | 5                          | |
|  | MCR                           | 3                          | |
|  | Project Office                | 2                          | |
|  | Aish                          | 1                          | |
|  | CPF                           | 1                          | |
|  | CCR                           | 0                          | |
|  +-------------------------------+----------------------------+ |
|  | Total Headcount               | 12                         | |
|  +-------------------------------+----------------------------+ |
|                                                                 |
|  [ 📊 Open Live Report Dashboard → ]                           |
|                                                                 |
|  © Vedanta Resources Ltd. • Save Food Platform                  |
+-----------------------------------------------------------------+
```

---

## 3. SharePoint Lists Referenced

| List Name | Purpose | OData Filter / Columns |
|---|---|---|
| **`RoleMatrix`** | Source of Admin recipients | `Title eq 'Application Roles'`<br>Extract emails from multi-person column: `Admin` |
| **`LocationMaster`** | Active mess facilities list | `IsActive eq 1`<br>Columns: `Title` |
| **`MealOptOutRequest`** | Transactional opt-out records | `MealType eq '<Shift>' and MealDate ge '<Date>T00:00:00Z' and MealDate le '<Date>T23:59:59Z'` |

---

## 4. Implementation Options

You can implement this in Power Automate in either of two ways:
- **Option A (Recommended: 1 Unified Scheduled Flow)**: A single scheduled recurrence flow running at 10:00, 18:00, and 20:00 IST that uses a `Switch` block.
- **Option B (3 Individual Separate Flows)**: Three distinct cloud flows (`SaveFood - Lunch Cut-Off`, `SaveFood - Dinner Cut-Off`, `SaveFood - Breakfast Cut-Off`).

---

# Complete Step-by-Step Guide: Option A (Single Unified Master Flow)

### Step 1: Create a Scheduled Cloud Flow
1. Sign in to [Power Automate](https://make.powerautomate.com/).
2. Select your environment, click **+ Create** -> **Scheduled cloud flow**.
3. Configure the trigger:
   - **Flow name**: `SaveFood - Daily Meal Cut-Off Notifications`
   - **Starting**: Today's date, `10:00:00 AM`
   - **Repeat every**: `1 Day`
4. Click **Create**.
5. In the designer, select the **Recurrence** trigger and configure:
   - **Time zone**: `(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi`
   - **At these hours**: `10, 18, 20`
   - **At these minutes**: `0`

---

### Step 2: Initialize Flow Variables
Add the following **Initialize variable** actions:

1. **`varMealType`**
   - Type: `String`
   - Value: *(leave blank)*
2. **`varTargetDate`**
   - Type: `String`
   - Value: *(leave blank)*
3. **`varTargetDateFormatted`**
   - Type: `String`
   - Value: *(leave blank)*
4. **`varShiftTitle`**
   - Type: `String`
   - Value: *(leave blank)*
5. **`varTableRowsHtml`**
   - Type: `String`
   - Value: `""` *(empty string)*
6. **`varTotalCount`**
   - Type: `Integer`
   - Value: `0`
7. **`varAdminEmails`**
   - Type: `String`
   - Value: `""` *(empty string)*

---

### Step 3: Determine Shift by Current IST Hour (`Switch`)
Add an action: **Control** -> **Switch**.
- **On**:
  ```
  int(formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'HH'))
  ```

#### Case 1: Value equals `10` (Lunch Cut-Off at 10:00 AM IST)
Add **Set variable** actions inside Case 10:
- `varMealType`: `Lunch`
- `varTargetDate`:
  ```
  formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'yyyy-MM-dd')
  ```
- `varTargetDateFormatted`:
  ```
  formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'dddd, dd MMM yyyy')
  ```
- `varShiftTitle`: `🥗 Lunch Opt-Out Summary`

#### Case 2: Value equals `18` (Dinner Cut-Off at 06:00 PM IST)
Add **Set variable** actions inside Case 18:
- `varMealType`: `Dinner`
- `varTargetDate`:
  ```
  formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'yyyy-MM-dd')
  ```
- `varTargetDateFormatted`:
  ```
  formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'dddd, dd MMM yyyy')
  ```
- `varShiftTitle`: `🍲 Dinner Opt-Out Summary`

#### Case 3: Value equals `20` (Breakfast Cut-Off at 08:00 PM IST for Next Day)
Add **Set variable** actions inside Case 20:
- `varMealType`: `Breakfast`
- `varTargetDate`:
  ```
  formatDateTime(addDays(convertFromUtc(utcNow(), 'India Standard Time'), 1), 'yyyy-MM-dd')
  ```
- `varTargetDateFormatted`:
  ```
  formatDateTime(addDays(convertFromUtc(utcNow(), 'India Standard Time'), 1), 'dddd, dd MMM yyyy')
  ```
- `varShiftTitle`: `🍳 Tomorrow's Breakfast Opt-Out Summary`

#### Default Branch (For Manual Testing at Any Other Hour):
> [!IMPORTANT]
> When you test the flow manually outside of 10 AM, 6 PM, or 8 PM (e.g., at 3:00 PM), the flow executes the **Default** branch. Add these fallback **Set variable** actions inside **Default** so manual tests always succeed:
> - `varMealType`: `Lunch`
> - `varTargetDate`:
>   ```
>   formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'yyyy-MM-dd')
>   ```
> - `varTargetDateFormatted`:
>   ```
>   formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'dddd, dd MMM yyyy')
>   ```
> - `varShiftTitle`: `🥗 Lunch Opt-Out Summary (Manual Test Run)`

---

### Step 4: Get Recipient Admins from `RoleMatrix`
After the Switch block, add a SharePoint **Get items** action:
1. **Action**: `Get items`
   - **Site Address**: `https://vedanta.sharepoint.com/sites/SaveFoodUAT`
   - **List Name**: `RoleMatrix`
   - **Filter Query**: `Title eq 'Application Roles'`
   - **Top Count**: `1`

2. **Action**: **Data Operation** -> **Select**
   - **From**:
     ```
     first(body('Get_items_-_RoleMatrix')?['value'])?['Admin']
     ```
   - **Map** (switch to text mode using the icon on top right of Map):
     ```
     item()?['Email']
     ```

3. **Action**: **Data Operation** -> **Join**
   - **From**: Output of `Select`
   - **Join with**: `;`

4. **Action**: **Set variable**
   - **Name**: `varAdminEmails`
   - **Value**: Output of `Join`

*(Fallback: If no admins are found or testing on a dev site, set a default email such as your own).*

---

### Step 5: Get Active Canteen Locations from `LocationMaster`
Add a SharePoint **Get items** action:
- **Site Address**: `https://vedanta.sharepoint.com/sites/SaveFoodUAT`
- **List Name**: `LocationMaster`
- **Filter Query**: `IsActive eq 1`
- **Order By**: `Title asc`
- **Top Count**: `100`

---

### Step 6: Get Opt-Out Requests for Target Shift
Add a SharePoint **Get items** action:
- **Site Address**: `https://vedanta.sharepoint.com/sites/SaveFoodUAT`
- **List Name**: `MealOptOutRequest`
- **Filter Query**:
  ```
  MealType eq '@{variables('varMealType')}' and MealDate ge '@{variables('varTargetDate')}T00:00:00Z' and MealDate le '@{variables('varTargetDate')}T23:59:59Z'
  ```
- **Top Count**: `5000`

> [!NOTE]
> **Why `Top Count: 5000` is the Optimal Setting**:
> - `5000` is Microsoft's official maximum limit allowed for a single SharePoint query in Power Automate without errors.
> - Because our query filters strictly for **a single shift on a single date** (e.g. *Today's Lunch*), total cancellations will typically be between 10 and 500 records.
> - Setting `Top Count: 5000` gives you maximum headroom while executing in a single, fast request (~0.5s) with **Pagination left OFF**.

---

### Step 7: Build Location-Wise Summary Table
Add an **Apply to each** action:
- **Select an output from previous steps**: Output `value` of `Get items - LocationMaster`

> [!IMPORTANT]
> **Keep Concurrency Control OFF (Default Sequential Mode)**:
> In the Settings of this `Apply to each` action, keep **Concurrency Control OFF**.
> - Because this loop updates variables (`Append to string variable` and `Increment variable`), running in parallel causes **race conditions** (counts get dropped or rows get jumbled).
> - Since there are only 6–7 locations and all data operations (`Filter array`, `Compose`) run entirely in-memory, the sequential loop finishes in **less than 1 second** while guaranteeing 100% data accuracy and alphabetical ordering!

Inside the loop:
1. Add action: **Data Operation** -> **Filter array**
   - **From**: Output `value` of `Get items - MealOptOutRequest`
   - **Condition**:
     - Left: `item()?['Location']?['Value']` *(or select **Location Value** from Dynamic Content)*
     - Operator: `is equal to`
     - Right: `items('Apply_to_each')?['Title']` *(or select **Title** from Dynamic Content under LocationMaster)*

   > [!IMPORTANT]
   > In SharePoint Lookup fields, the text name is stored in `.Value`, **NOT** `.Title` (i.e. `item()?['Location']?['Value']`). If `.Title` is used, it evaluates to `null`, causing the comparison to fail and return 0.

2. Add action: **Data Operation** -> **Compose** (Name: `Compose_Location_Count`)
   - **Inputs**:
     ```
     length(body('Filter_array'))
     ```

3. Add action: **Increment variable**
   - **Name**: `varTotalCount`
   - **Value**: Output of `Compose_Location_Count`

4. Add action: **Append to string variable**
   - **Name**: `varTableRowsHtml`
   - **Value**:
     ```html
     <tr style="border-bottom: 1px solid #e2e8f0;">
       <td style="padding: 10px 16px; font-weight: 600; color: #1e293b; font-size: 13.5px;">@{items('Apply_to_each')?['Title']}</td>
       <td style="padding: 10px 16px; text-align: center; font-weight: 700; color: #004B87; font-size: 14.5px;">@{outputs('Compose_Location_Count')}</td>
     </tr>
     ```

---

### Step 8: Send the Minimal Summary Email
After the loop completes, add an action: **Office 365 Outlook** -> **Send an email (V2)**:

- **To**: `@{variables('varAdminEmails')}`
- **Subject**: `[Save Food] @{variables('varShiftTitle')} — @{variables('varTargetDateFormatted')}`
- **Importance**: `Normal`
- **Body** (Click the `</>` Code View button to paste HTML):

```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Save Food Notification</title>
</head>
<body style="margin: 0; padding: 24px 12px; background-color: #f8fafc; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b;">
  <div style="max-width: 580px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; border: 1px solid #cbd5e1; overflow: hidden; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);">
    
    <!-- Header Strip -->
    <div style="background-color: #004B87; padding: 20px 24px; text-align: left; border-bottom: 4px solid #00A859;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td>
            <span style="font-size: 11px; font-weight: 800; letter-spacing: 1.2px; text-transform: uppercase; color: #93c5fd;">VEDANTA SUSTAINABILITY PLATFORM</span>
            <h1 style="margin: 4px 0 0 0; color: #ffffff; font-size: 20px; font-weight: 700; line-height: 1.3;">Save Food Management System</h1>
          </td>
        </tr>
      </table>
    </div>

    <!-- Title & Date -->
    <div style="padding: 18px 24px 10px 24px; background-color: #ffffff;">
      <h2 style="margin: 0; color: #0f172a; font-size: 17px; font-weight: 700;">@{variables('varShiftTitle')}</h2>
      <p style="margin: 4px 0 0 0; color: #64748b; font-size: 13.5px; font-weight: 500;">@{variables('varTargetDateFormatted')}</p>
    </div>

    <!-- Location-Wise Headcount Summary Table -->
    <div style="padding: 10px 24px 20px 24px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse; border: 1px solid #cbd5e1; border-radius: 12px; overflow: hidden;">
        <thead>
          <tr style="background-color: #f1f5f9; border-bottom: 2px solid #cbd5e1;">
            <th align="left" style="padding: 10px 16px; font-size: 12px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.5px;">Canteen Location</th>
            <th align="center" style="padding: 10px 16px; font-size: 12px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.5px; width: 140px;">Opt-Out Headcount</th>
          </tr>
        </thead>
        <tbody>
          @{variables('varTableRowsHtml')}
          <!-- Total Headcount Row -->
          <tr style="background-color: #f8fafc; border-top: 2px solid #cbd5e1;">
            <td style="padding: 12px 16px; font-size: 14px; font-weight: 800; color: #0f172a; text-transform: uppercase;">Total Headcount</td>
            <td align="center" style="padding: 12px 16px; font-size: 16px; font-weight: 800; color: #004B87;">@{variables('varTotalCount')}</td>
          </tr>
        </tbody>
      </table>
    </div>

    <!-- Direct Report Page Link CTA -->
    <div style="padding: 0 24px 24px 24px; text-align: center;">
      <a href="https://vedanta.sharepoint.com/sites/SaveFoodUAT/SitePages/SaveFood.aspx#/reports" 
         target="_blank" 
         style="display: inline-block; background-color: #004B87; color: #ffffff; font-size: 13.5px; font-weight: 700; padding: 11px 24px; border-radius: 10px; text-decoration: none; box-shadow: 0 2px 4px rgba(0, 75, 135, 0.25);">
        📊 Open Live Report Dashboard →
      </a>
    </div>

    <!-- Minimal Corporate Footer -->
    <div style="background-color: #f8fafc; border-top: 1px solid #e2e8f0; padding: 12px 24px; text-align: center;">
      <p style="margin: 0; color: #94a3b8; font-size: 11.5px;">© Vedanta Resources Ltd. • Save Food Management Platform • Automated Cut-Off Dispatch</p>
    </div>

  </div>
</body>
</html>
```

---

# Alternative Option B: Three Dedicated Modular Flows

If you prefer separate independent flows for Lunch, Dinner, and Breakfast:

### 1. `SaveFood - Lunch Cut-Off Report`
- **Trigger**: Recurrence daily at `10:00 AM IST` (`Hours: 10, Minutes: 0`).
- **Meal Type**: Fixed to `Lunch`.
- **Target Date**: Today (`formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'yyyy-MM-dd')`).
- **Subject**: `[Save Food] 🥗 Lunch Opt-Out Summary — @{outputs('Target_Date')}`

### 2. `SaveFood - Dinner Cut-Off Report`
- **Trigger**: Recurrence daily at `06:00 PM IST` (`Hours: 18, Minutes: 0`).
- **Meal Type**: Fixed to `Dinner`.
- **Target Date**: Today (`formatDateTime(convertFromUtc(utcNow(), 'India Standard Time'), 'yyyy-MM-dd')`).
- **Subject**: `[Save Food] 🍲 Dinner Opt-Out Summary — @{outputs('Target_Date')}`

### 3. `SaveFood - Breakfast Cut-Off Report`
- **Trigger**: Recurrence daily at `08:00 PM IST` (`Hours: 20, Minutes: 0`).
- **Meal Type**: Fixed to `Breakfast`.
- **Target Date**: Tomorrow (`formatDateTime(addDays(convertFromUtc(utcNow(), 'India Standard Time'), 1), 'yyyy-MM-dd')`).
- **Subject**: `[Save Food] 🍳 Tomorrow's Breakfast Opt-Out Summary — @{outputs('Target_Date')}`

All other steps (getting admins from `RoleMatrix`, querying locations, building the table, and sending email) remain identical.

---

## 5. Verification & Testing Instructions

1. **Manual Test Run**:
   - In Power Automate, click **Test** -> **Manually** -> **Run Flow**.
   - Note: If testing outside of 10:00, 18:00, or 20:00, temporarily add your current hour to the `Switch` cases or test individual actions.
2. **Verify Admin Resolution**:
   - Inspect the output of the `Join` action to confirm it extracted all email addresses from `RoleMatrix.Admin`.
3. **Verify Email Rendering**:
   - Open Outlook (Desktop / Web / Mobile).
   - Check that the email displays:
     - Shift title and date in header.
     - Location-wise table with clean borders.
     - **Total Headcount** summary row at the bottom.
     - The **Open Live Report Dashboard** link directs to `#/reports`.

---

## 6. Troubleshooting Common Issues

| Issue / Error | Cause | Solution |
| :--- | :--- | :--- |
| **All location counts show 0** even though records exist | In `Filter array`, the left condition is checking `item()?['Location']?['Title']`. SharePoint lookup columns return `{ "Id": 2, "Value": "MCR" }`. The property is `Value`, not `Title`. | Change left side of `Filter array` to `item()?['Location']?['Value']` (or pick **Location Value** from Dynamic Content). |
| **`Send an email (V2)` fails: Input parameter 'emailMessage/To' is required to be of type 'String/email'** | In the `Select` action, the mapping was configured as a JSON object `{"Email": ...}` instead of raw text mode. | Click the "Switch map to text mode" icon on the right side of `Map` in `Select`, and set the expression to `item()?['Email']`. |
| **DateTime 400 error during manual testing** | Manual test executed outside 10:00, 18:00, 20:00, leaving `varMealType` and `varTargetDate` blank. | Set fallback values inside the `Default` branch of the `Switch` (`varMealType = Lunch`, `varTargetDate = Today`). |
