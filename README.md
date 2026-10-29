# Prime One Clock

The store time clock app. Installed on each store iPad, and on Jacksonville phones.

- Live app: https://primeoneoperations-sudo.github.io/prime-one-clock/
- Backend: Google Apps Script "Prime One Clock" in andre@primeoneinc.com. Its code is kept here in `backend/` (Code.gs, Office.html, appsscript.json).
- Data: Google Sheet "SYSTEM · Time Clock" (punches, devices, visits, refused punches, logs). PINs: "SYSTEM · Clock PINs (restricted)", in Andre's Drive only.

No employee data, PINs or punches are stored here. A device only works after it is set up with a one-time iPad code from the Office page.

## Two links, one Google project
1. Kiosk API (iPads): Deploy as web app, Execute as Me, Who has access: Anyone. Its URL is `API` in index.html.
2. Office page (Juliana, Andre): a second web-app deployment of the same project, Who has access: Anyone within primeoneinc.com. Only the emails in `OFFICE_EMAILS` get in, checked on every action.

## Updating
1. Change the files here.
2. Raise `VERSION` in index.html and `APP_VERSION` in backend/Code.gs to the same number.
3. Paste backend/Code.gs and backend/Office.html into the Apps Script project, then Deploy › Manage deployments › edit each deployment › New version (so both links stay the same).
4. Every device reloads itself when its screen is idle, and every night at 3 AM.
