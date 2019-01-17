**GT Edition OCA is Global TecHub ls not available for everybody, please don't share it outside. In case of any doubts, don't hesitate to contact your manager, Jiri Novak or Vojtech Kulhavy.**

**GT** stands for Global TecHub.

OCA is our main quoting tool. Altough it offers wide portfolio of features and deliverables, our team decided to expand its features. and provide some basic automation as well as offer wider GUI settings. This was done using a well known browser plugin that injects scripts and styles. 

# Main features
- 1-click export of OCA file, XML file and fully modded Excel file with discount adjustments {+ use shift key during clicking for sdd]
- Switchable theme for OCA that removes paddings and empty spaces so the screen is better utilized, can be rotated, and is friendlier to eyes. Optionally you can change theme to dark.
Button for creating quote summary
- Button for switching quotes using UCID

# Implementation & installation:
1. Inctall Tampermonkey plugin for Chrome custom scripts:  https://chrome.google.com/webstore/detail/dhdgffkkebhmkfjojejmpbldmpobfkfo
2. install GT edition script for Tampermonkey: https://github.hpe.com/GlobalTecHub/GTeditionOCA/raw/master/GTeditionOCA.user.js

# Functionality

- The script is daily automatically updated via plugin refresh set in Chrome.
- It adds currently 5 buttons for
- toggling the theme which
  - works for both Internal and External (Akamai) and test environments
  - removes unnecessary paddings and borders
  - allows using OCA in portait mode
  - highlights riser types in DL380 Gen10
- toggling dark color scheme
- opening the quote using ucid
- showing summary
- making quick export of our delivarables - oca, xml and xlsx with 1 click. The CLIC check is performed automatically during each export. If there is unbuildable error, user is informed.

## Excel export (limited access - contact vojtech.kulhavy@hpe.com or jiri.novak@hpe.com)
different appearance of excel compared to the one in OCA
- almost matches Global TecHub previous appearance
- includes discounting formulas
- addition of padding of SKU to show nesting
- added CLIC check results in separate tab (click heart item links to BoM, so you can easily know which PN causes trouble). A short summary is in first tab of quote

#Tips and tricks
-Pressing Shift while clicking on the E button will generate the SDD file as well (handy for power calc in OneShow)
-Pressing Shift+E in menu shows only quoted products but it is editable.
-Chrome should ask you if you want to allow download multiple files. Once you allow it, it will allow you download multiple files without confirmation
-xml is from some reason considered as potentially harmful by Chrome. To override the behavior​, try setting to opposite value "Protect you and your device from dangerous sites" in Chrome Advanced settings (click ..., settings, advanced)
-Tampermonkey will check for script updates “every day” by default, but it can be reduced up to “every 6 hours” in the settings:
