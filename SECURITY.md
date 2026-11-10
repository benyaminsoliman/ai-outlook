# Security notes

This version is intended for a single user testing with their own API keys. It makes browser-origin requests to AI provider APIs. This means a key entered into the task pane can be inspected by someone with access to that browser session. Do not embed a permanent API key in source files or in `manifest.xml`.

For confidential business email or deployment to multiple users, use a backend proxy with authentication, server-side secret storage, rate limits, request logging controls, and appropriate data-retention agreements. Do not expose the provider key to the client. The Outlook manifest and task pane can remain the same while the API adapter is changed to call the private backend.
