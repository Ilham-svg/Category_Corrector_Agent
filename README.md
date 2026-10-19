# CartUp Category Update Agent

Updates the **Category** (and **Main Materials* → Other** when it appears) of
products in the **Rejected** tab of a seller's shop, found by Seller Code,
from an Excel file.

## Excel format

| Product Id | Category Path |
|-----------:|---------------|
| 8452 | Tools, DIY Outdoor > Lawn Garden > Outdoor Gardening Power Tools > Pressure Washers |

- Any number of `>` levels is supported (2, 3, 4, ...).
- Header names like `id`, `productID`, `sku` / `category` also work.

## Run it

Double-click `run_agent.bat` (or run `python seller_update_agent.py`).
The script asks in the terminal for:

1. **Seller Code** — typed into the 2nd search bar on the Seller List
   (Shop Name is the 1st), searched on Pending → Approve → Reject tabs until found.
2. **Path to Excel file** — you can drag the file into the terminal to paste its path.
3. **OTP code** — the panel always shows the login form; the script types the
   stored ID/password automatically, then you type the OTP in the terminal and
   the agent enters it and clicks **Verify**.

The login credentials live in constants `USERNAME` / `PASSWORD` at the top of
`seller_update_agent.py` — edit them there if they change.

## What it does

1. Opens the sign-in page, clicks **Sign in** if a landing button is shown, then
   enters the stored **ID + password** and submits.
2. Waits for the OTP screen, enters the code you type in the terminal and
   clicks **Verify** (continues immediately if no OTP is asked). The seller
   panel then loads and the flow continues.
3. **Seller Information → Seller List**.
4. Types your Seller Code into the 2nd search bar, checking
   **Pending / Approve / Reject** tabs until the seller row appears.
5. Clicks **Switch to Seller → Yes**.
6. **Product → Products → Offline → Rejected**.
7. For every Product Id in the Excel: sets the page size to 50 rows
   (paginator dropdown), finds the row (searching page by page to the end),
   clicks **Edit**, then in each category column **searches with short
   queries** — first word first (e.g. `Outdoor`), then the first two words
   (`Outdoor Gardening`), then the full text — clicking the option as soon
   as it appears in the filtered dropdown. After the last level it clicks
   **Yes, Continue**. If a **Main Materials\*** field is shown, its text bar
   is filled with **Other**. Then it scrolls down and clicks **Update**.
8. If the update comes back with a **Main Material** or **Warranty
   Type/Period missing** notification, the agent fills the requested field
   (`Other` / `No Warranty` in Warranty Type) and clicks **Update** again,
   up to 3 rounds.
9. Repeats until the Excel is finished, then writes
   `update_report.csv` next to your Excel (OK / NOT FOUND / FAILED per id)
   and saves step screenshots into `screenshots/`.

## If something doesn't match the site

All labels live in the constants at the top of `seller_update_agent.py`
(`LBL_SELLER_INFO`, `TABS_SHOP`, `TABS_OFFLINE`, `TABS_REJECTED`,
`MAIN_MATERIALS`, `MATERIALS_CHOICE`, `WARRANTY_TYPE`, `WARRANTY_VALUE`,
...). Edit them and re-run — screenshots in `screenshots/`
show exactly where it stopped.
