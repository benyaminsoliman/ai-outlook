/* AI Email Assistant for Outlook.
 * Runs inside the Outlook task pane. No API key or email is persisted by this app.
 */
(function () {
  "use strict";

  const PROVIDERS = {
    openai: {
      name: "OpenAI",
      style: "responses",
      endpoint: "https://api.openai.com/v1/responses",
      models: ["gpt-6-astra", "gpt-6.1-sol", "gpt-6.1-terra", "gpt-6-luna"]
    },
    gemini: {
      name: "Google Gemini",
      style: "gemini",
      endpoint: "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
      models: ["gemini-3.8-flash", "gemini-3.7-flash", "gemini-3.6-flash", "gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.1-pro-preview", "gemini-2.5-pro", "gemini-2.5-flash"]
    },
    qwen: {
      name: "Alibaba Qwen / Model Studio",
      style: "chat",
      endpoint: "https://{WorkspaceId}.ap-southeast-1.maas.aliyuncs.com/compatible-mode/v1/chat/completions",
      models: ["qwen3.8-max", "qwen3.7-plus", "qwen3.8-flash"],
      hint: "Replace {WorkspaceId} with your Alibaba Model Studio workspace ID. Use the URL for the region in which your API key was created."
    },
    anthropic: {
      name: "Anthropic Claude",
      style: "anthropic",
      endpoint: "https://api.anthropic.com/v1/messages",
      models: ["claude-opus-5-5", "claude-sonnet-5-5", "claude-fable-5-2", "claude-haiku-5-5"],
      hint: "Direct browser access is enabled for personal/testing use. Some Anthropic organizations prohibit browser-origin requests; use a private backend proxy in that case."
    },
    deepseek: {
      name: "DeepSeek",
      style: "chat",
      endpoint: "https://api.deepseek.com/chat/completions",
      models: ["deepseek-flash", "deepseek-v4-pro"]
    },
    openrouter: {
      name: "OpenRouter",
      style: "chat",
      endpoint: "https://openrouter.ai/api/v1/chat/completions",
      models: []
    },
    groq: {
      name: "Groq",
      style: "chat",
      endpoint: "https://api.groq.com/openai/v1/chat/completions",
      models: ["openai/gpt-oss-120b", "openai/gpt-oss-20b", "llama-3.3-70b-versatile", "llama-3.1-8b-instant", "qwen/qwen3.8-27b"]
    },
    mistral: {
      name: "Mistral AI",
      style: "chat",
      endpoint: "https://api.mistral.ai/v1/chat/completions",
      models: ["mistral-large-4", "mistral-small-2603"]
    },
    custom: {
      name: "Custom OpenAI-compatible API",
      style: "chat",
      endpoint: "https://your-provider.example/v1/chat/completions",
      models: [],
      hint: "Use the provider's full chat-completions URL. The API must accept browser CORS requests and OpenAI-compatible messages."
    }
  };

  const $ = (id) => document.getElementById(id);
  let officeReady = false;
  let currentItem = null;
  let currentMode = "unknown";
  let loadedEmail = null;
  let isBusy = false;

  function setStatus(message, kind) {
    const el = $("status");
    el.textContent = message;
    el.className = "status" + (kind ? " " + kind : "");
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (ch) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    })[ch]);
  }

  function htmlToText(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(String(html || ""), "text/html");
    doc.querySelectorAll("script,style,head").forEach((node) => node.remove());
    return (doc.body.textContent || "").replace(/\u00a0/g, " ").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  function officeAsync(operation) {
    return new Promise((resolve, reject) => {
      try {
        operation((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error(result.error && result.error.message ? result.error.message : "Outlook operation failed."));
        });
      } catch (error) { reject(error); }
    });
  }

  function getSubject(item) {
    if (item && item.subject && typeof item.subject.getAsync === "function") {
      return officeAsync((cb) => item.subject.getAsync(cb));
    }
    return Promise.resolve((item && item.subject) || "");
  }

  function getBodyText(item) {
    if (!item || !item.body || typeof item.body.getAsync !== "function") {
      return Promise.reject(new Error("Outlook did not expose the current message body."));
    }
    return officeAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
  }

  function getDisplayName(person) {
    if (!person) return "";
    if (typeof person === "string") return person;
    return [person.displayName, person.emailAddress ? `<${person.emailAddress}>` : ""].filter(Boolean).join(" ");
  }

  async function loadCurrentEmail() {
    if (!officeReady || !currentItem) throw new Error("Outlook is not ready or no message is selected.");
    const item = currentItem;
    const [subject, body] = await Promise.all([getSubject(item), getBodyText(item)]);
    const sender = getDisplayName(item.from || item.sender);
    const recipients = Array.isArray(item.to) ? item.to.map(getDisplayName).filter(Boolean).join("; ") : "";
    const context = [
      `Subject: ${subject || "(no subject)"}`,
      sender ? `From: ${sender}` : "",
      recipients ? `To: ${recipients}` : "",
      "",
      body || "(The message body is empty.)"
    ].filter((line) => line !== "").join("\n");
    loadedEmail = { subject: subject || "", sender, recipients, body: body || "", context };
    $("emailContext").value = context;
    $("emailMeta").textContent = `${currentMode === "compose" ? "Compose item" : "Selected message"} · ${subject || "No subject"}${sender ? " · From: " + sender : ""}`;
    setStatus("Email loaded", "success");
    return loadedEmail;
  }

  function updateProviderUI(resetEndpoint) {
    const id = $("provider").value;
    const provider = PROVIDERS[id];
    if (!provider) return;
    if (resetEndpoint || !$("endpoint").value) $("endpoint").value = provider.endpoint;
    const list = $("modelSuggestions");
    list.innerHTML = "";
    provider.models.forEach((model) => {
      const option = document.createElement("option");
      option.value = model;
      list.appendChild(option);
    });
    if (resetEndpoint || !$("model").value) $("model").value = provider.models[0] || "";
    $("endpointHint").textContent = provider.hint || "You can edit this URL if your provider account uses a different API region or gateway.";
    $("endpoint").disabled = ["openai", "gemini", "anthropic", "deepseek", "openrouter", "groq", "mistral"].includes(id);
    if (id === "qwen" || id === "custom") $("endpoint").disabled = false;
  }

  function currentSettings() {
    const providerId = $("provider").value;
    const key = $("apiKey").value.trim();
    const model = $("model").value.trim();
    const endpoint = $("endpoint").value.trim();
    const maxOutput = Math.max(128, Math.min(16000, Number.parseInt($("maxOutput").value, 10) || 1800));
    if (!key) throw new Error("Enter your API key first.");
    if (!model) throw new Error("Enter or select a model ID.");
    if (!endpoint) throw new Error("Enter the provider API endpoint.");
    if (providerId === "qwen" && endpoint.includes("{WorkspaceId}")) {
      throw new Error("Replace {WorkspaceId} in the Qwen endpoint with your Alibaba Model Studio workspace ID.");
    }
    return { providerId, provider: PROVIDERS[providerId], key, model, endpoint, maxOutput };
  }

  async function parseResponse(response) {
    const raw = await response.text();
    let data;
    try { data = raw ? JSON.parse(raw) : {}; }
    catch (_) { throw new Error(`The provider returned a non-JSON response (HTTP ${response.status}). Check the endpoint and CORS support.`); }
    if (!response.ok) {
      const message = data && data.error && (data.error.message || data.error.type)
        ? (data.error.message || data.error.type)
        : (data && data.message) || raw.slice(0, 500) || response.statusText;
      throw new Error(`${response.status} ${response.statusText}: ${message}`);
    }
    return data;
  }

  async function fetchWithTimeout(url, options) {
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), 120000);
    try {
      return await fetch(url, { ...options, signal: controller.signal });
    } catch (error) {
      if (error.name === "AbortError") throw new Error("The AI request timed out after 120 seconds.");
      if (error instanceof TypeError) throw new Error("Network/CORS error. Confirm the endpoint, API key, internet connection, and whether the provider allows browser-origin API requests.");
      throw error;
    } finally { window.clearTimeout(timeout); }
  }

  function extractOpenAICompatible(data) {
    const content = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
    if (typeof content === "string") return content.trim();
    if (Array.isArray(content)) return content.map((part) => part.text || "").join("").trim();
    throw new Error("The provider response did not contain choices[0].message.content.");
  }

  async function callAI(systemPrompt, userPrompt) {
    const cfg = currentSettings();
    const headers = { "Content-Type": "application/json" };
    let url = cfg.endpoint;
    let body;

    if (cfg.provider.style === "responses") {
      headers.Authorization = `Bearer ${cfg.key}`;
      body = {
        model: cfg.model,
        instructions: systemPrompt,
        input: userPrompt,
        max_output_tokens: cfg.maxOutput
      };
    } else if (cfg.provider.style === "gemini") {
      url = cfg.endpoint.replace("{model}", encodeURIComponent(cfg.model));
      headers["x-goog-api-key"] = cfg.key;
      body = {
        systemInstruction: { parts: [{ text: systemPrompt }] },
        contents: [{ role: "user", parts: [{ text: userPrompt }] }],
        generationConfig: { temperature: 0.35, maxOutputTokens: cfg.maxOutput }
      };
    } else if (cfg.provider.style === "anthropic") {
      headers["x-api-key"] = cfg.key;
      headers["anthropic-version"] = "2026-10-21";
      headers["anthropic-dangerous-direct-browser-access"] = "true";
      body = {
        model: cfg.model,
        max_tokens: cfg.maxOutput,
        system: systemPrompt,
        messages: [{ role: "user", content: userPrompt }]
      };
    } else {
      headers.Authorization = `Bearer ${cfg.key}`;
      if (cfg.providerId === "openrouter") {
        headers["HTTP-Referer"] = window.location.origin;
        headers["X-Title"] = "AI Email Assistant for Outlook";
      }
      body = {
        model: cfg.model,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userPrompt }
        ],
        temperature: 0.35,
        max_tokens: cfg.maxOutput
      };
    }

    const response = await fetchWithTimeout(url, { method: "POST", headers, body: JSON.stringify(body) });
    const data = await parseResponse(response);

    if (cfg.provider.style === "responses") {
      if (typeof data.output_text === "string" && data.output_text.trim()) return data.output_text.trim();
      const textParts = [];
      (data.output || []).forEach((item) => (item.content || []).forEach((part) => {
        if (part.type === "output_text" && part.text) textParts.push(part.text);
      }));
      if (textParts.length) return textParts.join("\n").trim();
      throw new Error("The OpenAI response did not contain output text.");
    }
    if (cfg.provider.style === "gemini") {
      const parts = data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts;
      const text = Array.isArray(parts) ? parts.map((part) => part.text || "").join("").trim() : "";
      if (!text) throw new Error("Gemini returned no text. Check model access, safety filtering, and API quota.");
      return text;
    }
    if (cfg.provider.style === "anthropic") {
      const text = (data.content || []).filter((part) => part.type === "text").map((part) => part.text || "").join("\n").trim();
      if (!text) throw new Error("Claude returned no text content.");
      return text;
    }
    return extractOpenAICompatible(data);
  }

  async function refreshModels() {
    const cfg = currentSettings();
    const providerId = cfg.providerId;
    setStatus("Loading model list…");
    let url;
    const headers = {};
    if (providerId === "gemini") {
      url = "https://generativelanguage.googleapis.com/v1beta/models";
      headers["x-goog-api-key"] = cfg.key;
    } else if (providerId === "anthropic") {
      url = "https://api.anthropic.com/v1/models";
      headers["x-api-key"] = cfg.key;
      headers["anthropic-version"] = "2026-10-21";
      headers["anthropic-dangerous-direct-browser-access"] = "true";
    } else {
      url = cfg.endpoint.replace(/\/chat\/completions\/?$/, "/models").replace(/\/responses\/?$/, "/models");
      headers.Authorization = `Bearer ${cfg.key}`;
    }
    const response = await fetchWithTimeout(url, { method: "GET", headers });
    const data = await parseResponse(response);
    let ids = [];
    if (providerId === "gemini") {
      ids = (data.models || []).filter((m) => !m.supportedGenerationMethods || m.supportedGenerationMethods.includes("generateContent"))
        .map((m) => String(m.name || "").replace(/^models\//, ""));
    } else if (providerId === "anthropic") {
      ids = (data.data || []).map((m) => m.id).filter(Boolean);
    } else {
      ids = (data.data || data.models || []).map((m) => typeof m === "string" ? m : m.id).filter(Boolean);
    }
    ids = [...new Set(ids)].sort();
    if (!ids.length) throw new Error("The provider did not return model IDs. Enter a model ID manually; not every provider exposes model discovery.");
    const list = $("modelSuggestions");
    list.innerHTML = "";
    ids.forEach((id) => {
      const option = document.createElement("option");
      option.value = id;
      list.appendChild(option);
    });
    if (!ids.includes($("model").value)) $("model").value = ids[0];
    setStatus(`${ids.length} models loaded`, "success");
  }

  function buildSystemPrompt() {
    return "You are a careful professional email assistant. Produce clear, accurate, useful business writing. Preserve facts, names, dates, amounts, technical details, contractual meaning, and commitments exactly as supplied. Never invent facts, approvals, attachments, promises, or deadlines. If information is missing, use a neutral placeholder or state that it needs confirmation. Do not send emails or claim that an email has been sent. Output only the requested result without prefatory commentary unless asked.";
  }

  function emailText() {
    const text = $("emailContext").value.trim();
    if (!text) throw new Error("Load the current email or paste the email text first.");
    return text;
  }

  function userPreferences() {
    return `Tone: ${$("tone").value}.\nLanguage: ${$("language").value}.\nAdditional instructions: ${$("extraInstructions").value.trim() || "None."}`;
  }

  async function runAction(action) {
    if (isBusy) return;
    isBusy = true;
    const buttons = ["loadModels", "loadEmail", "summarize", "draftReply", "improveDraft", "insertOutput", "openReply"];
    buttons.forEach((id) => $(id).disabled = true);
    try {
      let result = "";
      const preferences = userPreferences();
      if (action === "summarize") {
        const email = emailText();
        setStatus("Summarizing email…");
        result = await callAI(buildSystemPrompt(), `Summarize the following email or email thread. Provide: (1) a concise overview, (2) key points and decisions, (3) action items with owners and dates when explicitly stated, and (4) unanswered questions or risks. Do not infer missing details.\n\n${preferences}\n\nEMAIL:\n${email}`);
      } else if (action === "draftReply") {
        const email = emailText();
        setStatus("Drafting reply…");
        result = await callAI(buildSystemPrompt(), `Draft a reply to the email below. Address the sender appropriately when their name is available. Respond to the actual points raised, keep the message coherent and suitable to send after human review, and do not add unsupported commitments. Do not include a subject line unless necessary.\n\n${preferences}\n\nEMAIL:\n${email}`);
      } else if (action === "improveDraft") {
        const draft = $("emailContext").value.trim() || $("output").value.trim();
        if (!draft) throw new Error("Load or paste the draft you want improved.");
        setStatus("Improving draft…");
        result = await callAI(buildSystemPrompt(), `Improve the following email draft for clarity, grammar, structure, and professional tone. Preserve the original meaning, technical content, dates, amounts, and level of firmness. Do not introduce new facts or commitments. Return only the revised email.\n\n${preferences}\n\nDRAFT:\n${draft}`);
      }
      $("output").value = result;
      setStatus("Done — review the output", "success");
    } catch (error) {
      setStatus("Action failed", "error");
      window.alert(error.message || String(error));
    } finally {
      isBusy = false;
      buttons.forEach((id) => $(id).disabled = false);
    }
  }

  function insertIntoDraft() {
    const text = $("output").value.trim();
    if (!text) throw new Error("There is no AI output to insert.");
    if (!officeReady || !currentItem) throw new Error("Outlook is not ready.");
    if (currentMode !== "compose" || !currentItem.body || typeof currentItem.body.setSelectedDataAsync !== "function") {
      throw new Error("Open a new message or reply in compose mode, then use Insert into current draft.");
    }
    currentItem.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) setStatus("Inserted into draft — review before sending", "success");
      else setStatus(result.error ? result.error.message : "Could not insert text", "error");
    });
  }

  function openAsReply() {
    const text = $("output").value.trim();
    if (!text) throw new Error("There is no AI output to open as a reply.");
    if (!officeReady || !currentItem || typeof currentItem.displayReplyForm !== "function") {
      throw new Error("Open a received email in read mode to use Open as a reply.");
    }
    const html = text.split(/\r?\n/).map((line) => line ? `<p>${escapeHtml(line)}</p>` : "<p>&nbsp;</p>").join("");
    currentItem.displayReplyForm({ htmlBody: html });
    setStatus("Reply window opened — review before sending", "success");
  }

  function attachEvents() {
    $("provider").addEventListener("change", () => updateProviderUI(true));
    $("loadModels").addEventListener("click", async () => {
      try { await refreshModels(); } catch (error) { setStatus("Could not load models", "error"); window.alert(error.message || String(error)); }
    });
    $("loadEmail").addEventListener("click", async () => {
      try { await loadCurrentEmail(); } catch (error) { setStatus("Could not load email", "error"); window.alert(error.message || String(error)); }
    });
    $("summarize").addEventListener("click", () => runAction("summarize"));
    $("draftReply").addEventListener("click", () => runAction("draftReply"));
    $("improveDraft").addEventListener("click", () => runAction("improveDraft"));
    $("insertOutput").addEventListener("click", () => {
      try { insertIntoDraft(); } catch (error) { window.alert(error.message || String(error)); }
    });
    $("openReply").addEventListener("click", () => {
      try { openAsReply(); } catch (error) { window.alert(error.message || String(error)); }
    });
  }

  function initializeOutlook() {
    if (typeof Office === "undefined" || !Office.onReady) {
      setStatus("Open this page inside Outlook", "error");
      return;
    }
    Office.onReady((info) => {
      if (info.host !== Office.HostType.Outlook) {
        setStatus("Open this add-in from Outlook", "error");
        return;
      }
      officeReady = true;
      currentItem = Office.context.mailbox.item;
      currentMode = currentItem && currentItem.subject && typeof currentItem.subject.getAsync === "function" ? "compose" : "read";
      $("emailMeta").textContent = currentMode === "read"
        ? "Received message selected. Load it to review its content."
        : "Compose window detected. Load the current draft or paste an email thread.";
      $("insertOutput").style.display = currentMode === "compose" ? "block" : "none";
      $("openReply").style.display = currentMode === "read" ? "block" : "none";
      $("improveDraft").style.display = currentMode === "compose" ? "block" : "none";
      setStatus("Connected to Outlook", "success");
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    attachEvents();
    updateProviderUI(true);
    initializeOutlook();
  });
})();
