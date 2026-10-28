# Prime One Clock

The store time clock app (screens only). Installed on each store iPad, and on Jacksonville phones.

- Live app: https://primeoneoperations-sudo.github.io/prime-one-clock/
- Backend: Google Apps Script "Prime One Clock" in andre@primeoneinc.com
- Data: Google Sheet "SYSTEM · Time Clock (punches, fixes, devices)" in Drive › PRIME ONE SYSTEM › Data

No employee data, PINs or punches are stored here. A device only works after it is set up with a one-time code from the Office screen.

## Updating the app
1. Change the files here.
2. Raise `VERSION` in index.html and `APP_VERSION` in the backend to the same number, then deploy the backend
   (Deploy › Manage deployments › edit › New version, so the link stays the same).
3. Every device reloads itself within a minute when its screen is idle, and every night at 3 AM.
