CARLISLE EOM STOCK SYSTEM - V16 INSTALLATION & USAGE
--------------------------------------------------

1. ARCHITECTURE:
   This system decouples structural sheet protections from user access control.
   - Protections (Menu 1): Defines WHICH cells can be edited. Only Owners/MDs are editors of the sheets; others edit via unblocked ranges.
   - Access Control (Menu 2): Defines WHO can open the file and which sheets they are authorized to edit.

2. TRIGGER SETUP:
   To ensure the onEdit permission checks and automation work for all users, you MUST set up an Installable Trigger:
   - In Apps Script editor, go to Triggers (Clock icon on the left).
   - Click "Add Trigger".
   - Choose "onEdit" for the function to run.
   - Choose "From spreadsheet" for the event source.
   - Choose "On edit" for the event type.
   - Save. This runs the script with your authority, allowing identity checks.

3. SYSTEM_ACCESS:
   - Use Menu 2 > Rebuild SYSTEM_ACCESS to initialize the control sheet.
   - Enter staff emails and assign Roles.
   - In "Sheets Controlled", you can type sheet names (e.g., CS BAR) or groups (e.g., ALL CS SHEETS).
   - Use Menu 2 > Sync User Permissions to add staff as file editors.

4. PROTECTIONS:
   - Use Menu 1 to apply protections.
   - For large groups (CS Sheets, Admin Sheets), use the "Next Sheet" buttons to process sheets one by one and avoid Google's execution timeouts.
   - Formula cells are automatically LOCKED even within editable ranges.

5. WORKFLOW:
   - Staff enter movements in the STOCK MOVEMENT APPROVAL LOG.
   - Status defaults to PENDING.
   - Only MD/Owner can change status to APPROVED.
   - Upon approval, the system automatically:
     a) Updates the Master Price List cost (from Purchases).
     b) Updates the specific Department sheet quantities.

6. REPORTS:
   - Located in the "Carlisle Reports" menu.
   - Use "Stock Audit Summary" for a full cross-department view.
   - Use "Formula Error Check" to find broken calculations workbook-wide.
