# AI Email Assistant for Outlook — Multi-model package

This package contains one Outlook add-in and one XML manifest. The add-in can call multiple AI providers; you do **not** need a separate Outlook manifest for every model.

## Included features

- Read the currently selected email or compose draft.
- Summarize an email or thread.
- Draft a reply using a selectable tone and language.
- Improve a draft while preserving its meaning and details.
- Insert generated text into an open compose draft, or open a reply window from a received message.
- Provider and model selector with editable model ID.
- Optional model discovery through each provider's model-list endpoint where available.
- API keys are kept in task-pane memory only; this app does not save them.
- The add-in does **not** send email. You review and send manually.

## Supported provider presets

| Provider | API style | Preset models / notes |
|---|---|---|
| OpenAI | Responses API | `gpt-6-astra`, `gpt-6.1-sol`, `gpt-6.1-terra`, `gpt-6-luna` |
| Google Gemini | `generateContent` | `gemini-3.8-flash`, `gemini-3.7-flash`, `gemini-3.6-flash`, `gemini-3.5-flash`, `gemini-3.5-flash-lite`, `gemini-3.1-pro-preview`, `gemini-2.5-pro`, `gemini-2.5-flash` |
| Alibaba Qwen / Model Studio | OpenAI-compatible Chat Completions | `qwen3.8-max`, `qwen3.7-plus`, `qwen3.8-flash`; set your Workspace ID and region in the endpoint |
| Anthropic Claude | Messages API | `claude-opus-5-5`, `claude-sonnet-5-5`, `claude-fable-5-2`, `claude-haiku-5-5` |
| DeepSeek | OpenAI-compatible Chat Completions | `deepseek-flash`, `deepseek-v4-pro` |
| OpenRouter | OpenAI-compatible Chat Completions | Load available models or enter any model slug enabled for your account |
| Groq | OpenAI-compatible Chat Completions | `openai/gpt-oss-120b`, `openai/gpt-oss-20b`, `llama-3.3-70b-versatile`, `llama-3.1-8b-instant`, `qwen/qwen3.8-27b` |
| Mistral AI | OpenAI-compatible Chat Completions | `mistral-large-4`, `mistral-small-2603` |
| Custom | OpenAI-compatible Chat Completions | Supply a full endpoint URL and model ID |

Model names and availability can change. Use **Load models** where the provider exposes a model catalogue, or enter the current model ID supplied by your provider. Model access depends on your account, region, and API plan.

## Important deployment requirement

Outlook does not run the HTML/JavaScript files directly from the XML manifest. It loads them from an HTTPS web address. Therefore, publish the contents of this folder to an HTTPS host before sideloading the manifest.

### Option A — GitHub Pages (simple static hosting)

1. Create a GitHub repository named `outlook-ai-email-assistant`.
2. Upload `index.html`, `styles.css`, `taskpane.js`, and the `assets` folder to the repository root. You can also upload this README.
3. In the repository settings, enable **Pages** and publish from the main branch / root folder.
4. Wait until GitHub Pages shows the published HTTPS address, for example `https://YOUR-GITHUB-USERNAME.github.io/outlook-ai-email-assistant/`.
5. Open `manifest.xml` in a text editor and replace every occurrence of `https://YOUR-HOST/outlook-ai-email-assistant` with your exact GitHub Pages base URL (without an extra trailing slash after the folder name). Save the XML.

You can instead host these files on your own HTTPS web server. Keep the relative file structure intact.

### Option B — Use your own web host

Copy `index.html`, `styles.css`, `taskpane.js`, `assets/`, and optionally `README.md` into a folder served over HTTPS. Replace the placeholder URL in `manifest.xml` with that folder's public HTTPS URL.

## Install in Outlook

1. Confirm that your Outlook client supports Outlook Mailbox requirement set 1.3 or later.
2. In Outlook, open **Get Add-ins** / **Manage Add-ins** (the exact menu varies by Outlook version and account type).
3. Open **My add-ins** or the custom add-ins section, choose **Add a custom add-in** / **Add from file**, and select the configured `manifest.xml`.
4. Open a received email and choose **AI Email Assistant** from the message ribbon. It is also available from the compose ribbon when writing a new message or reply.

For Microsoft 365 work/school accounts, an administrator may need to allow custom add-ins or deploy the manifest centrally.

## First use

1. Open the add-in from Outlook.
2. Select a provider.
3. Paste the API key created in that provider's developer console.
4. Select a model ID, or press **Load models** when supported. You can type a different ID directly.
5. Press **Load current email**, review the extracted text, and choose **Summarize email** or **Draft a reply**.
6. Review the output. In compose mode, select **Insert into current draft**. In read mode, select **Open as a reply**.
7. Review the draft in Outlook and send it yourself when ready.

## Qwen / Alibaba Model Studio endpoint

The default endpoint is a Singapore-region template:

`https://{WorkspaceId}.ap-southeast-1.maas.aliyuncs.com/compatible-mode/v1/chat/completions`

Replace `{WorkspaceId}` with your Model Studio workspace ID. Use the endpoint for the same region in which your API key/workspace was created. Alibaba region-specific API keys and endpoints are not interchangeable. If your workspace is in another region, change the endpoint to that region's documented URL.

## Security and privacy — read before using

- This is a **client-side** add-in: provider API requests are made directly from the Outlook task pane. The API key is therefore visible to the browser runtime and to anyone who can inspect that local session. Do not use a shared computer or a public hosted copy with a key you do not trust.
- For company-wide or confidential production use, the recommended design is a private HTTPS backend that stores API keys server-side and authenticates each user. This package does not include or host such a backend.
- Email text is sent to whichever provider you select when you run an AI action. Check the provider's privacy, data-retention, and organizational terms before processing confidential company emails.
- Attachments are not extracted or sent. Only the text of the selected email/draft is used.
- The add-in does not store your API key or email text in local storage, cookies, or a database.
- The add-in never sends email automatically. It only creates or fills a draft for you to review.
- Browser CORS restrictions may prevent direct API calls for some providers or account configurations. For Anthropic, browser requests require the `anthropic-dangerous-direct-browser-access` header and may still be blocked by organization-level settings. In that case use a backend proxy.

## Troubleshooting

- **The manifest does not load:** verify the XML is well-formed, every URL uses your exact HTTPS host, and the host returns the files successfully.
- **Blank task pane:** verify `index.html`, `styles.css`, `taskpane.js`, and `assets` were uploaded to the exact paths in the manifest. Open the HTML URL in a browser to check for 404 errors.
- **Invalid API key / model:** check that the key belongs to the selected provider and that the model is enabled for your account/region.
- **Qwen URL error:** replace `{WorkspaceId}` and check your Model Studio region.
- **Network/CORS error:** check whether the provider permits browser-origin calls. Use a private backend proxy when required.
- **Model not found:** use **Load models** or paste the exact current model ID from the provider's API documentation.
- **Insert button unavailable:** open a new message or reply in compose mode. To create a reply from a received message, use **Open as a reply**.

## Files

- `manifest.xml` — Outlook XML manifest template.
- `index.html` — add-in user interface.
- `styles.css` — task-pane styling.
- `taskpane.js` — Outlook integration and provider API adapters.
- `assets/` — Outlook ribbon icons.
