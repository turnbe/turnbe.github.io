# Internal Use Label Helper

Upload the entire `outlook-internal-use` folder to the root of `turnbe/turnbe.github.io`.

Published base URL: `https://turnbe.github.io/outlook-internal-use/`

After GitHub Pages updates, verify `https://turnbe.github.io/outlook-internal-use/taskpane.html` opens in a browser.

Then sideload the local `manifest.xml` into Outlook using Add a custom add-in > Add from File.

## First-test behaviour
The add-in applies `Internal Use` only when Outlook reports that the new draft has no sensitivity label.
If the University's default label is already applied, this version deliberately leaves it unchanged.

Use the `Label status` button on a compose window and click `Check current label`. This reports the label
name and GUID. That lets the next version safely replace only the organisational default, without replacing
a deliberate Confidential, Restricted, or other selection.
