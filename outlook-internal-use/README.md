# Sensitivity Label Prompt v3

Behaviour:
- Reads the current Outlook/Purview sensitivity label.
- If it is `Unlabelled Sensitivity`, the Choose label task pane offers:
  Public, Internal Use, Private Confidential, Business Confidential, Restricted.
- It resolves labels by name from your published label catalogue, rather than hard-coding GUIDs.
- If another label is already selected, it leaves it unchanged.
- No message content is sent anywhere.

Important: this build prompts through the `Choose label` ribbon button. It does NOT yet block Send or automatically open the task pane. A true send-time mandatory prompt requires Smart Alerts / OnMessageSend and should be implemented separately rather than claiming the compose event can block Send.
