# University of Otago Outlook Sensitivity Label Add-in
Version 4.2

## Change in 4.2
When a new message is composed, the add-in checks its current Purview sensitivity label.
If the label is `Unlabelled Sensitivity`, it attempts to open the add-in task pane automatically
so the five label choices are visible without first clicking the ribbon button.

The send-time Smart Alerts SoftBlock remains as the safety net. If a message is still
`Unlabelled Sensitivity` when Send is clicked, Outlook stops the send and asks the user
to choose a label.

## Labels
- Public
- Internal Use
- Private Confidential
- Business Confidential
- Restricted

## Hosted path
https://turnbe.github.io/UO-Outlook-Sensitivity-Label-Addin/

Upload all files over the existing v4.1 files. Because runtime.js is hosted, the code update
will publish from GitHub Pages. The manifest version is also updated to 4.2.0.0.

Note: automatic task-pane opening is dependent on the Outlook client supporting
Office.addin.showAsTaskpane in the event-based runtime. If classic Outlook declines to
open it automatically, the existing ribbon button and send-time Smart Alert remain functional.
