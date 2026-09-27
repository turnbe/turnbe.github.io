# University of Otago Outlook Sensitivity Label Add-in
Version 4.1

## GitHub Pages location
Upload this package to:
`UO-Outlook-Sensitivity-Label-Addin/`

Hosted base URL:
`https://turnbe.github.io/UO-Outlook-Sensitivity-Label-Addin/`

## Behaviour
- Retains the tested five-label picker: Public, Internal Use, Private Confidential, Business Confidential, Restricted.
- If a message is already classified, the add-in leaves it unchanged.
- At Send, `Unlabelled Sensitivity` triggers the Smart Alerts send-time check.
- Current enforcement mode is `SoftBlock`.
- The existing `Choose label` ribbon control remains available.
- Uses the University of Otago artwork supplied for this build.
- No message content is transmitted to GitHub or another backend.

## Upgrade
1. Create/rename the GitHub Pages folder to `UO-Outlook-Sensitivity-Label-Addin`.
2. Upload all files from this package.
3. Confirm `https://turnbe.github.io/UO-Outlook-Sensitivity-Label-Addin/taskpane.html` loads.
4. Reinstall/update `manifest.xml`, because its hosted resource URLs have changed.
5. Restart classic Outlook before testing.

Test first with an Unlabelled Sensitivity draft, then with each of the five valid labels.
