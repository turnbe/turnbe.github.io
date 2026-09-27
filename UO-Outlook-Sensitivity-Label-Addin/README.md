# University of Otago Outlook Sensitivity Label Add-in
Version 4.1

## GitHub Pages location
Upload this package to hosted folder:
`UO-Outlook-Sensitivity-Label-Addin/`

Hosted base URL:
`https://turnbe.github.io/UO-Outlook-Sensitivity-Label-Addin/`

## Purpose
The University of Otago uses Microsoft Purview sensitivity labels to support the appropriate classification and handling of information. The default 'Unlabelled Sensitivity' classification requires users to make a conscious decision about the appropriate sensitivity of information before it is shared.

This Outlook add-in provides a simple control at the point an email is sent. If a message remains classified as 'Unlabelled Sensitivity', the add-in prompts the user to select one of the University's approved sensitivity labels before sending the message. This provides the effect of mandatory labelling with the user being able to chose this experience if they wish.
The add-in does not determine the classification on behalf of the user and does not change messages that have already been classified. Its purpose is to ensure that classification is an intentional decision, while keeping the process simple and integrated into the normal Outlook workflow.

The add-in uses the University's existing Microsoft Purview sensitivity labels and does not transmit email content, recipients, attachments or other message information to an external service.

## Behaviour
- Retains the tested five-label picker: Public, Internal Use, Private Confidential, Business Confidential, Restricted.
- If a message is already classified, the add-in leaves it unchanged.
- At Send, `Unlabelled Sensitivity` triggers the Smart Alerts send-time check.
- Current enforcement mode is `SoftBlock`.
- The existing `Choose label` ribbon control remains available.
- Uses the University of Otago artwork supplied for this build.
- No message content is transmitted to GitHub or another backend.

## Prerequisites
- Microsoft Outlook with support for modern Outlook web add-ins, event based activation and the Microsoft sensitivity label APIs. The add-in has been developed and tested using classic Outlook for Windows.
- Microsoft Purview sensitivity labels must be enabled for the user's mailbox and published to the user.
  
- The University sensitivity labels expected by the add-in are:
- Public
- Internal Use
- Private Confidential
- Business Confidential
- Restricted
- Unlabelled Sensitivity

- The user's Microsoft 365 licensing must support the Outlook sensitivity label APIs.
- Outlook add-ins must be enabled within the Microsoft 365 tenant (AppsForOfficeEnabled = True).
- For individual installation or testing, the user must have the Exchange My Custom Apps role, allowing a custom Outlook add-in manifest to be installed for their mailbox. Production deployment may instead be centrally managed through Microsoft 365.
- The device must be able to access the HTTPS location hosting the add-in resources. The current development deployment is hosted using GitHub Pages.
- Internet connectivity is normally required to load the add-in resources. The add-in uses SoftBlock for its send time control so an inability to load the add-in does not prevent email from being sent indefinitely.
- No local administrator access, Outlook COM add-in, locally installed executable, API key or external AI service is required.

## Architecture
The add-in is designed to be lightweight and to make use of the University's existing Microsoft 365 and Microsoft Purview capabilities rather than introducing a separate classification system or service.

The Outlook add-in manifest registers the add-in with Outlook and defines the Outlook functionality it uses, including the ribbon command and the send time event.

The add-in's HTML and JavaScript files are hosted on an HTTPS web service. The current development implementation uses GitHub Pages. These files provide the user interface and the logic used to check and apply sensitivity labels.

When the add-in runs, it uses the Microsoft Office JavaScript APIs within Outlook to interact with the sensitivity labels already published to the user through Microsoft Purview. Label identifiers are obtained from the user's available sensitivity label catalogue rather than being hard coded into the add-in.

When a user selects a classification, the add-in instructs Outlook to apply the corresponding Microsoft Purview sensitivity label to the message. The classification therefore remains part of the University's existing Microsoft 365 information protection environment.

At send time, the add-in checks the message's current sensitivity label. If the message remains classified as Unlabelled Sensitivity, the send is interrupted and the user is prompted to select an appropriate classification. Messages that already have another sensitivity label are not changed.

The add-in uses SoftBlock for this control. This provides classification enforcement when the add-in is available while avoiding a dependency that could prevent email from being sent if the add-in or its hosted resources are temporarily unavailable.

## Data Flow and Privacy
The hosted component contains only the static files required to operate the add-in. Email content is not sent to or processed by the web host.

The add-in does not transmit message bodies, recipients, subject lines, attachments or message content to GitHub Pages or another external service. Interaction with the message and its sensitivity classification occurs through the Microsoft Office APIs within the user's Outlook and Microsoft 365 environment.

No separate database, external API, analytics service or backend processing service is required.

In simplified form, the architecture is:
User → Outlook → Outlook Add-in → Microsoft Office APIs → Microsoft Purview Sensitivity Labels

The web host provides the add-in code to Outlook but does not form part of the information flow for the content of the email.
