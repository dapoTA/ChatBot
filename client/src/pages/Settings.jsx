import { useState, useMemo } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import { queryClient } from "@/lib/queryClient";
import { apiRequest } from "@/lib/queryClient";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { useToast } from "@/hooks/use-toast";
import {
  CheckCircle,
  XCircle,
  RefreshCw,
  Wifi,
  Save,
  Clock,
  AlertTriangle,
  Bot,
} from "lucide-react";
import { extractStyledPhrases, extractGlobalResponseStyle } from "@/lib/styleParser";
import { insertSharepointConfigSchema, insertAppSettingsSchema } from "@shared/schema";

// ─── Appearance form ────────────────────────────────────────────────────────

const appearanceSchema = insertAppSettingsSchema.extend({
  assistantName: z.string().min(1, "Assistant name is required"),
  welcomeMessage: z.string().min(1, "Welcome message is required"),
  notFoundMessage: z.string().min(1, "Not-found message is required"),
  customInstructions: z.string().optional().nullable(),
});

// ─── AI Parameters form ───────────────────────────────────────────────────────

const aiParamsSchema = z.object({
  temperature: z.coerce.number().min(0).max(2),
  topP: z.coerce.number().min(0).max(1),
  maxTokens: z.coerce.number().int().min(100).max(4096),
  frequencyPenalty: z.coerce.number().min(0).max(2),
  presencePenalty: z.coerce.number().min(0).max(2),
});

// ─── SharePoint form ─────────────────────────────────────────────────────────

const sharepointSchema = z.object({
  mode: z.enum(["onprem", "online"]),
  siteUrl: z.string().url("Must be a valid URL"),
  siteUrlOnprem: z.string().optional().nullable(),
  siteUrlOnline: z.string().optional().nullable(),
  libraryName: z.string().min(1, "Library name is required"),
  domain: z.string().optional().nullable(),
  username: z.string().optional().nullable(),
  password: z.string().optional().nullable(),
  allowSelfSigned: z.boolean().default(true),
  tenantId: z.string().optional().nullable(),
  clientId: z.string().optional().nullable(),
  clientSecret: z.string().optional().nullable(),
}).superRefine((data, ctx) => {
  if (data.mode === "onprem") {
    if (!data.domain) ctx.addIssue({ code: "custom", message: "Domain is required", path: ["domain"] });
    if (!data.username) ctx.addIssue({ code: "custom", message: "Username is required", path: ["username"] });
    if (!data.password) ctx.addIssue({ code: "custom", message: "Password is required", path: ["password"] });
  }
  // Online credentials (tenantId, clientId, clientSecret) are not validated here —
  // they may be supplied via environment variables server-side. The server will
  // surface a clear error at test/sync time if credentials are missing.
});

export default function Settings() {
  const { toast } = useToast();
  const [testResult, setTestResult] = useState(null);

  // ─── Fetch current settings ────────────────────────────────────────────────

  const { data: appSettingsData } = useQuery({
    queryKey: ["/api/settings"],
  });

  const { data: config } = useQuery({
    queryKey: ["/api/sharepoint/config"],
  });

  // ─── Appearance form ───────────────────────────────────────────────────────

  const appearanceForm = useForm({
    resolver: zodResolver(appearanceSchema),
    defaultValues: {
      assistantName: "ON-PNT® Assistant",
      welcomeMessage: "Ask me anything about your SharePoint documents.",
      notFoundMessage: "I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.",
      customInstructions: "",
    },
    values: appSettingsData
      ? {
          assistantName: appSettingsData.assistantName,
          welcomeMessage: appSettingsData.welcomeMessage,
          notFoundMessage: appSettingsData.notFoundMessage,
          customInstructions: appSettingsData.customInstructions ?? "",
        }
      : undefined,
  });

  const saveAppearance = useMutation({
    mutationFn: (data) => apiRequest("POST", "/api/settings", data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/settings"] });
      toast({ title: "Appearance saved", description: "Widget name and welcome message updated." });
    },
    onError: () => {
      toast({ title: "Save failed", description: "Could not save appearance settings.", variant: "destructive" });
    },
  });

  // ─── AI Parameters form ────────────────────────────────────────────────────

  const aiParamsForm = useForm({
    resolver: zodResolver(aiParamsSchema),
    defaultValues: { temperature: 0, topP: 1, maxTokens: 1500, frequencyPenalty: 0, presencePenalty: 0 },
    values: appSettingsData
      ? {
          temperature: appSettingsData.temperature ?? 0,
          topP: appSettingsData.topP ?? 1,
          maxTokens: appSettingsData.maxTokens ?? 1500,
          frequencyPenalty: appSettingsData.frequencyPenalty ?? 0,
          presencePenalty: appSettingsData.presencePenalty ?? 0,
        }
      : undefined,
  });

  const saveAiParams = useMutation({
    mutationFn: (data) => apiRequest("POST", "/api/settings", data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/settings"] });
      toast({ title: "AI parameters saved", description: "Model settings updated successfully." });
    },
    onError: () => {
      toast({ title: "Save failed", description: "Could not save AI parameters.", variant: "destructive" });
    },
  });

  // ─── SharePoint form ───────────────────────────────────────────────────────

  const sharepointForm = useForm({
    resolver: zodResolver(sharepointSchema),
    defaultValues: {
      mode: "onprem",
      siteUrl: "",
      siteUrlOnprem: "",
      siteUrlOnline: "",
      libraryName: "Documents",
      domain: "",
      username: "",
      password: "",
      allowSelfSigned: true,
      tenantId: "",
      clientId: "",
      clientSecret: "",
    },
    values: config
      ? {
          mode: config.mode ?? "onprem",
          siteUrl: config.siteUrl,
          // Each tab remembers its own URL independently
          siteUrlOnprem: config.siteUrlOnprem ?? (config.mode === "onprem" ? config.siteUrl : ""),
          siteUrlOnline: config.siteUrlOnline ?? (config.mode === "online" ? config.siteUrl : ""),
          libraryName: config.libraryName,
          domain: config.domain ?? "",
          username: config.username ?? "",
          password: config.password ?? "",
          allowSelfSigned: config.allowSelfSigned,
          tenantId: config.tenantId ?? "",
          clientId: config.clientId ?? "",
          clientSecret: config.clientSecret ?? "",
        }
      : undefined,
  });

  const spMode = sharepointForm.watch("mode");

  const saveSharepoint = useMutation({
    mutationFn: (data) => apiRequest("POST", "/api/sharepoint/config", data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/sharepoint/config"] });
      toast({ title: "Settings saved", description: "SharePoint configuration has been saved." });
    },
    onError: () => {
      toast({ title: "Save failed", description: "Could not save configuration.", variant: "destructive" });
    },
  });

  const testConnection = useMutation({
    mutationFn: () => apiRequest("POST", "/api/sharepoint/test", sharepointForm.getValues()),
    onSuccess: async (res) => {
      const result = await res.json();
      setTestResult(result);
    },
    onError: () => {
      setTestResult({ success: false, message: "Request failed. Check server logs." });
    },
  });

  const syncDocuments = useMutation({
    mutationFn: () => apiRequest("POST", "/api/sharepoint/sync"),
    onSuccess: async (res) => {
      const result = await res.json();
      queryClient.invalidateQueries({ queryKey: ["/api/documents"] });
      queryClient.invalidateQueries({ queryKey: ["/api/sharepoint/config"] });
      toast({
        title: result.synced > 0 ? "Sync complete" : "Sync finished",
        description: result.message,
        variant: result.synced > 0 ? "default" : "destructive",
      });
    },
    onError: () => {
      toast({
        title: "Sync failed",
        description: "Could not sync documents. Check your SharePoint configuration.",
        variant: "destructive",
      });
    },
  });

  // ─── Formatting preview (computed live from textarea) ─────────────────────
  const instructionsValue = appearanceForm.watch("customInstructions") ?? "";
  const previewPhrases = useMemo(() => extractStyledPhrases(instructionsValue), [instructionsValue]);
  const previewBodyStyle = useMemo(() => extractGlobalResponseStyle(instructionsValue), [instructionsValue]);
  const hasPreview = previewPhrases.length > 0 || previewBodyStyle !== null;

  return (
    <div className="min-h-screen bg-background">
      <div className="max-w-2xl mx-auto px-6 py-10 space-y-8">

        {/* Page Header */}
        <div>
          <h1 className="text-2xl font-semibold text-foreground" data-testid="text-settings-title">
            Administration
          </h1>
          <p className="text-sm text-muted-foreground mt-1">
            Configure the chat widget appearance and SharePoint document library connection.
          </p>
        </div>

        {/* ── Widget Appearance ───────────────────────────────────────────── */}
        <form onSubmit={appearanceForm.handleSubmit((d) => saveAppearance.mutate(d))}>
          <div className="rounded-xl border border-border bg-card p-6 space-y-5">
            <h2 className="text-sm font-semibold text-foreground uppercase tracking-wide">
              Widget Appearance
            </h2>

            <div className="space-y-1">
              <Label htmlFor="assistantName">Assistant Name</Label>
              <Input
                id="assistantName"
                placeholder="ON-PNT® Assistant"
                data-testid="input-assistant-name"
                {...appearanceForm.register("assistantName")}
              />
              {appearanceForm.formState.errors.assistantName && (
                <p className="text-xs text-destructive">{appearanceForm.formState.errors.assistantName.message}</p>
              )}
              <p className="text-xs text-muted-foreground">
                Displayed in the chat header and trigger button tooltip
              </p>
            </div>

            <div className="space-y-1">
              <Label htmlFor="welcomeMessage">Welcome Message</Label>
              <Textarea
                id="welcomeMessage"
                placeholder="Ask me anything about your SharePoint documents."
                rows={3}
                data-testid="input-welcome-message"
                className="resize-none"
                {...appearanceForm.register("welcomeMessage")}
              />
              {appearanceForm.formState.errors.welcomeMessage && (
                <p className="text-xs text-destructive">{appearanceForm.formState.errors.welcomeMessage.message}</p>
              )}
              <p className="text-xs text-muted-foreground">
                Shown when the chat is opened with no prior messages
              </p>
            </div>

            <div className="space-y-1">
              <Label htmlFor="notFoundMessage">Not Found Message</Label>
              <Textarea
                id="notFoundMessage"
                placeholder="I'm sorry, I couldn't find relevant information for your request..."
                rows={3}
                data-testid="input-not-found-message"
                className="resize-none"
                {...appearanceForm.register("notFoundMessage")}
              />
              {appearanceForm.formState.errors.notFoundMessage && (
                <p className="text-xs text-destructive">{appearanceForm.formState.errors.notFoundMessage.message}</p>
              )}
              <p className="text-xs text-muted-foreground">
                Returned by the assistant when no relevant documents are found for a query
              </p>
            </div>

            <div className="space-y-1">
              <Label htmlFor="customInstructions">Custom Instructions</Label>
              <Textarea
                id="customInstructions"
                placeholder={`Describe how the assistant should behave. For example:\n- Always respond in a formal, professional tone\n- When citing documents, include the section or page number if available\n- This assistant serves the Facilities Management department\n- Prioritise safety-related documents when relevant\n- If asked about pricing, always note that figures may have changed`}
                rows={8}
                data-testid="input-custom-instructions"
                className="resize-y font-mono text-xs"
                {...appearanceForm.register("customInstructions")}
              />
              <p className="text-xs text-muted-foreground">
                Shape the assistant's tone, focus, persona, and style. Applied to every conversation.
                Any not-found or fallback response you write here will be automatically removed before reaching the AI. Use the <strong>Not Found Message</strong> field above — it is the sole source of truth for what the assistant says when an answer isn't in the documents.
              </p>
            </div>

            {/* ── Formatting Preview ─────────────────────────────────────── */}
            <div className="space-y-2" data-testid="formatting-preview-section">
              <Label className="text-xs text-muted-foreground uppercase tracking-wide">
                Formatting Preview
              </Label>
              <div className="rounded-lg border border-border bg-muted/30 p-3">
                {hasPreview ? (
                  <div className="flex gap-2 items-start">
                    <div className="w-6 h-6 rounded-full bg-primary flex items-center justify-center flex-shrink-0 mt-0.5">
                      <Bot className="w-3 h-3 text-primary-foreground" />
                    </div>
                    <div
                      className="bg-white dark:bg-secondary border border-border rounded-lg rounded-bl-none p-3 text-sm leading-relaxed max-w-sm"
                      style={previewBodyStyle ?? {}}
                      data-testid="formatting-preview-bubble"
                    >
                      {previewPhrases.map((p, i) => (
                        <p
                          key={i}
                          className="mb-1"
                          dangerouslySetInnerHTML={{ __html: p.html }}
                        />
                      ))}
                      <p>
                        Here is the information you requested based on the available documents.
                      </p>
                    </div>
                  </div>
                ) : (
                  <p className="text-xs text-muted-foreground italic" data-testid="formatting-preview-empty">
                    No special formatting detected — add a phrase like{" "}
                    <code className="bg-muted px-1 rounded">"Hello!" in bold red text</code>{" "}
                    or{" "}
                    <code className="bg-muted px-1 rounded">Always respond in bold green</code>{" "}
                    to see a live preview here.
                  </p>
                )}
              </div>
              <div className="mt-2 space-y-1.5">
                <p className="text-xs font-medium text-muted-foreground">How formatting instructions work:</p>
                <div className="space-y-1 text-xs text-muted-foreground">
                  <div className="flex gap-2">
                    <span className="text-primary font-bold shrink-0">①</span>
                    <span>
                      <strong>Styled prefix</strong> — wrap a phrase in quotes followed by a style:{" "}
                      <code className="bg-muted px-1 py-0.5 rounded text-[11px]">Always begin with "Thank you for your question!" in bold blue text</code>
                    </span>
                  </div>
                  <div className="flex gap-2">
                    <span className="text-primary font-bold shrink-0">②</span>
                    <span>
                      <strong>Global body style</strong> — apply a colour or weight to all responses:{" "}
                      <code className="bg-muted px-1 py-0.5 rounded text-[11px]">Always respond in bold green</code>
                    </span>
                  </div>
                </div>
              </div>
            </div>

            <Button type="submit" disabled={saveAppearance.isPending} data-testid="button-save-appearance">
              <Save className="w-4 h-4 mr-2" />
              {saveAppearance.isPending ? "Saving…" : "Save Appearance"}
            </Button>
          </div>
        </form>

        {/* ── AI Model Parameters ─────────────────────────────────────────── */}
        <form onSubmit={aiParamsForm.handleSubmit((d) => saveAiParams.mutate(d))}>
          <div className="rounded-xl border border-border bg-card p-6 space-y-5">
            <div>
              <h2 className="text-sm font-semibold text-foreground uppercase tracking-wide">
                AI Model Parameters
              </h2>
              <p className="text-xs text-muted-foreground mt-1">
                Control how the AI model generates responses. Changes apply to all new conversations.
              </p>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <Label htmlFor="temperature">Temperature</Label>
                <Input
                  id="temperature"
                  type="number"
                  min="0" max="2" step="0.1"
                  data-testid="input-temperature"
                  {...aiParamsForm.register("temperature")}
                />
                {aiParamsForm.formState.errors.temperature && (
                  <p className="text-xs text-destructive">{aiParamsForm.formState.errors.temperature.message}</p>
                )}
                <p className="text-xs text-muted-foreground">Range: 0–2 · 0 = focused &amp; deterministic, 2 = creative &amp; varied · Neutral (no impact): 1.0</p>
              </div>

              <div className="space-y-1">
                <Label htmlFor="topP">Top P</Label>
                <Input
                  id="topP"
                  type="number"
                  min="0" max="1" step="0.05"
                  data-testid="input-top-p"
                  {...aiParamsForm.register("topP")}
                />
                {aiParamsForm.formState.errors.topP && (
                  <p className="text-xs text-destructive">{aiParamsForm.formState.errors.topP.message}</p>
                )}
                <p className="text-xs text-muted-foreground">Range: 0–1 · 1 = all tokens considered, 0.1 = only most likely tokens · Neutral (no impact): 1.0</p>
              </div>
            </div>

            <div className="space-y-1">
              <Label htmlFor="maxTokens">Max Tokens</Label>
              <Input
                id="maxTokens"
                type="number"
                min="100" max="4096" step="1"
                data-testid="input-max-tokens"
                {...aiParamsForm.register("maxTokens")}
              />
              {aiParamsForm.formState.errors.maxTokens && (
                <p className="text-xs text-destructive">{aiParamsForm.formState.errors.maxTokens.message}</p>
              )}
              <p className="text-xs text-muted-foreground">Range: 100–4096 · Maximum length of each AI response (1 token ≈ 4 characters) · Neutral (no impact): 4096</p>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <Label htmlFor="frequencyPenalty">Frequency Penalty</Label>
                <Input
                  id="frequencyPenalty"
                  type="number"
                  min="0" max="2" step="0.1"
                  data-testid="input-frequency-penalty"
                  {...aiParamsForm.register("frequencyPenalty")}
                />
                {aiParamsForm.formState.errors.frequencyPenalty && (
                  <p className="text-xs text-destructive">{aiParamsForm.formState.errors.frequencyPenalty.message}</p>
                )}
                <p className="text-xs text-muted-foreground">Range: 0–2 · Higher = less word repetition in responses · Neutral (no impact): 0.0</p>
              </div>

              <div className="space-y-1">
                <Label htmlFor="presencePenalty">Presence Penalty</Label>
                <Input
                  id="presencePenalty"
                  type="number"
                  min="0" max="2" step="0.1"
                  data-testid="input-presence-penalty"
                  {...aiParamsForm.register("presencePenalty")}
                />
                {aiParamsForm.formState.errors.presencePenalty && (
                  <p className="text-xs text-destructive">{aiParamsForm.formState.errors.presencePenalty.message}</p>
                )}
                <p className="text-xs text-muted-foreground">Range: 0–2 · Higher = encourages covering new topics · Neutral (no impact): 0.0</p>
              </div>
            </div>

            <Button type="submit" disabled={saveAiParams.isPending} data-testid="button-save-ai-params">
              <Save className="w-4 h-4 mr-2" />
              {saveAiParams.isPending ? "Saving…" : "Save AI Parameters"}
            </Button>
          </div>
        </form>

        {/* ── SharePoint Connection ───────────────────────────────────────── */}
        <form onSubmit={sharepointForm.handleSubmit((d) => saveSharepoint.mutate(d))} className="space-y-4">
          <div className="rounded-xl border border-border bg-card p-6 space-y-5">
            <h2 className="text-sm font-semibold text-foreground uppercase tracking-wide">
              SharePoint Connection
            </h2>

            {/* ── Mode toggle ─────────────────────────────────────────────── */}
            <div className="space-y-2">
              <Label>Connection Type</Label>
              <div className="flex rounded-lg border border-border overflow-hidden" data-testid="toggle-sp-mode">
                <button
                  type="button"
                  onClick={() => {
                    if (spMode !== "onprem") {
                      // Save the current online URL, restore the last onprem URL
                      sharepointForm.setValue("siteUrlOnline", sharepointForm.getValues("siteUrl"));
                      sharepointForm.setValue("siteUrl", sharepointForm.getValues("siteUrlOnprem") ?? "");
                    }
                    sharepointForm.setValue("mode", "onprem");
                    setTestResult(null);
                  }}
                  data-testid="button-mode-onprem"
                  className={`flex-1 px-4 py-2 text-sm font-medium transition-colors ${
                    spMode === "onprem"
                      ? "bg-primary text-primary-foreground"
                      : "bg-background text-muted-foreground hover:bg-muted"
                  }`}
                >
                  On-Premises (NTLM)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    if (spMode !== "online") {
                      // Save the current onprem URL, restore the last online URL
                      sharepointForm.setValue("siteUrlOnprem", sharepointForm.getValues("siteUrl"));
                      sharepointForm.setValue("siteUrl", sharepointForm.getValues("siteUrlOnline") ?? "");
                    }
                    sharepointForm.setValue("mode", "online");
                    setTestResult(null);
                  }}
                  data-testid="button-mode-online"
                  className={`flex-1 px-4 py-2 text-sm font-medium transition-colors border-l border-border ${
                    spMode === "online"
                      ? "bg-primary text-primary-foreground"
                      : "bg-background text-muted-foreground hover:bg-muted"
                  }`}
                >
                  SharePoint Online (OAuth)
                </button>
              </div>
              <p className="text-xs text-muted-foreground">
                {spMode === "onprem"
                  ? "SharePoint 2019 / 2016 hosted on your own servers. Uses NTLM with Active Directory credentials."
                  : "Microsoft 365 cloud SharePoint. Uses an Azure AD app registration with client credentials."}
              </p>
            </div>

            {/* ── Shared fields ────────────────────────────────────────────── */}
            <div className="space-y-1">
              <Label htmlFor="siteUrl">SharePoint Site URL</Label>
              <Input
                id="siteUrl"
                placeholder={spMode === "online" ? "https://company.sharepoint.com/sites/mysite" : "http://sharepoint.company.com/sites/mysite"}
                data-testid="input-site-url"
                {...sharepointForm.register("siteUrl")}
              />
              {sharepointForm.formState.errors.siteUrl && (
                <p className="text-xs text-destructive">{sharepointForm.formState.errors.siteUrl.message}</p>
              )}
              <p className="text-xs text-muted-foreground">Full URL to your SharePoint site (not just the root)</p>
            </div>

            <div className="grid grid-cols-2 gap-4">
              {spMode === "onprem" && (
              <div className="space-y-1">
                <Label htmlFor="domain">Active Directory Domain</Label>
                <Input
                  id="domain"
                  placeholder="COMPANY"
                  data-testid="input-domain"
                  {...sharepointForm.register("domain")}
                />
                {sharepointForm.formState.errors.domain && (
                  <p className="text-xs text-destructive">{sharepointForm.formState.errors.domain.message}</p>
                )}
              </div>
              )}
              <div className="space-y-1">
                <Label htmlFor="libraryName">Document Library Name</Label>
                <Input
                  id="libraryName"
                  placeholder="Documents"
                  data-testid="input-library-name"
                  {...sharepointForm.register("libraryName")}
                />
                {sharepointForm.formState.errors.libraryName && (
                  <p className="text-xs text-destructive">{sharepointForm.formState.errors.libraryName.message}</p>
                )}
              </div>
            </div>

            {/* ── On-premises fields ──────────────────────────────────────── */}
            {spMode === "onprem" && (
              <>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <Label htmlFor="username">Username</Label>
                    <Input
                      id="username"
                      placeholder="svc_account"
                      autoComplete="username"
                      data-testid="input-username"
                      {...sharepointForm.register("username")}
                    />
                    {sharepointForm.formState.errors.username && (
                      <p className="text-xs text-destructive">{sharepointForm.formState.errors.username.message}</p>
                    )}
                  </div>
                  <div className="space-y-1">
                    <Label htmlFor="password">Password</Label>
                    <Input
                      id="password"
                      type="password"
                      autoComplete="current-password"
                      data-testid="input-password"
                      {...sharepointForm.register("password")}
                    />
                    {sharepointForm.formState.errors.password && (
                      <p className="text-xs text-destructive">{sharepointForm.formState.errors.password.message}</p>
                    )}
                  </div>
                </div>

                <div className="flex items-center justify-between rounded-lg border border-border bg-muted/30 px-4 py-3">
                  <div>
                    <p className="text-sm font-medium text-foreground">Allow self-signed certificates</p>
                    <p className="text-xs text-muted-foreground">
                      Enable if your SharePoint uses an internal or untrusted SSL certificate
                    </p>
                  </div>
                  <Switch
                    checked={sharepointForm.watch("allowSelfSigned")}
                    onCheckedChange={(v) => sharepointForm.setValue("allowSelfSigned", v)}
                    data-testid="switch-self-signed"
                  />
                </div>
              </>
            )}

            {/* ── SharePoint Online fields ─────────────────────────────────── */}
            {spMode === "online" && (
              <>
                {/* Env-var banner when all three are set via environment */}
                {config?.envControlled?.tenantId && config?.envControlled?.clientId && config?.envControlled?.clientSecret && (
                  <div className="flex items-center gap-2 rounded-lg border border-green-200 dark:border-green-800 bg-green-50 dark:bg-green-950 px-4 py-3 text-xs text-green-800 dark:text-green-200" data-testid="banner-env-controlled">
                    <CheckCircle className="w-3.5 h-3.5 flex-shrink-0" />
                    <span>All OAuth credentials are set via environment variables and secured server-side.</span>
                  </div>
                )}

                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <Label htmlFor="tenantId">Tenant ID</Label>
                    {config?.envControlled?.tenantId && (
                      <span className="text-[10px] font-medium uppercase tracking-wide px-1.5 py-0.5 rounded bg-green-100 dark:bg-green-900 text-green-700 dark:text-green-300" data-testid="badge-env-tenant-id">env var</span>
                    )}
                  </div>
                  {config?.envControlled?.tenantId ? (
                    <div className="flex items-center h-9 px-3 rounded-md border border-border bg-muted text-sm text-muted-foreground" data-testid="readonly-tenant-id">
                      Set via <code className="mx-1 text-xs bg-background px-1 rounded">SHAREPOINT_TENANT_ID</code>
                    </div>
                  ) : (
                    <Input
                      id="tenantId"
                      placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                      data-testid="input-tenant-id"
                      {...sharepointForm.register("tenantId")}
                    />
                  )}
                  {sharepointForm.formState.errors.tenantId && (
                    <p className="text-xs text-destructive">{sharepointForm.formState.errors.tenantId.message}</p>
                  )}
                  {!config?.envControlled?.tenantId && (
                    <p className="text-xs text-muted-foreground">Found in Azure Portal → Azure Active Directory → Overview</p>
                  )}
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <div className="flex items-center gap-2">
                      <Label htmlFor="clientId">Client ID</Label>
                      {config?.envControlled?.clientId && (
                        <span className="text-[10px] font-medium uppercase tracking-wide px-1.5 py-0.5 rounded bg-green-100 dark:bg-green-900 text-green-700 dark:text-green-300" data-testid="badge-env-client-id">env var</span>
                      )}
                    </div>
                    {config?.envControlled?.clientId ? (
                      <div className="flex items-center h-9 px-3 rounded-md border border-border bg-muted text-sm text-muted-foreground" data-testid="readonly-client-id">
                        Set via <code className="mx-1 text-xs bg-background px-1 rounded">SHAREPOINT_CLIENT_ID</code>
                      </div>
                    ) : (
                      <Input
                        id="clientId"
                        placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                        data-testid="input-client-id"
                        {...sharepointForm.register("clientId")}
                      />
                    )}
                    {sharepointForm.formState.errors.clientId && (
                      <p className="text-xs text-destructive">{sharepointForm.formState.errors.clientId.message}</p>
                    )}
                    {!config?.envControlled?.clientId && (
                      <p className="text-xs text-muted-foreground">App registration Application (client) ID</p>
                    )}
                  </div>
                  <div className="space-y-1">
                    <div className="flex items-center gap-2">
                      <Label htmlFor="clientSecret">Client Secret</Label>
                      {config?.envControlled?.clientSecret && (
                        <span className="text-[10px] font-medium uppercase tracking-wide px-1.5 py-0.5 rounded bg-green-100 dark:bg-green-900 text-green-700 dark:text-green-300" data-testid="badge-env-client-secret">env var</span>
                      )}
                    </div>
                    {config?.envControlled?.clientSecret ? (
                      <div className="flex items-center h-9 px-3 rounded-md border border-border bg-muted text-sm text-muted-foreground" data-testid="readonly-client-secret">
                        Set via <code className="mx-1 text-xs bg-background px-1 rounded">SHAREPOINT_CLIENT_SECRET</code>
                      </div>
                    ) : (
                      <Input
                        id="clientSecret"
                        type="password"
                        autoComplete="off"
                        placeholder="Client secret value"
                        data-testid="input-client-secret"
                        {...sharepointForm.register("clientSecret")}
                      />
                    )}
                    {sharepointForm.formState.errors.clientSecret && (
                      <p className="text-xs text-destructive">{sharepointForm.formState.errors.clientSecret.message}</p>
                    )}
                    {!config?.envControlled?.clientSecret && (
                      <p className="text-xs text-muted-foreground">App registration → Certificates &amp; secrets</p>
                    )}
                  </div>
                </div>

                <div className="rounded-lg bg-blue-50 dark:bg-blue-950 border border-blue-200 dark:border-blue-800 px-4 py-3 text-xs text-blue-800 dark:text-blue-200 space-y-1">
                  <p className="font-medium">Azure AD app registration requirements:</p>
                  <ul className="list-disc list-inside space-y-0.5 text-blue-700 dark:text-blue-300">
                    <li>API permissions: SharePoint → Sites.Read.All (application)</li>
                    <li>Admin consent granted for your tenant</li>
                    <li>Authentication: no redirect URI needed (client credentials flow)</li>
                  </ul>
                </div>
              </>
            )}
          </div>

          <div className="flex items-center gap-3">
            <Button type="submit" disabled={saveSharepoint.isPending} data-testid="button-save-config">
              <Save className="w-4 h-4 mr-2" />
              {saveSharepoint.isPending ? "Saving…" : "Save Connection"}
            </Button>
            <Button
              type="button"
              variant="outline"
              disabled={testConnection.isPending || !config}
              onClick={() => { setTestResult(null); testConnection.mutate(); }}
              data-testid="button-test-connection"
            >
              <Wifi className="w-4 h-4 mr-2" />
              {testConnection.isPending ? "Testing…" : "Test Connection"}
            </Button>
          </div>

          {testResult && (
            <div
              className={`flex items-start gap-3 rounded-lg px-4 py-3 text-sm border ${
                testResult.success
                  ? "bg-green-50 dark:bg-green-950 border-green-200 dark:border-green-800 text-green-800 dark:text-green-200"
                  : "bg-red-50 dark:bg-red-950 border-red-200 dark:border-red-800 text-red-800 dark:text-red-200"
              }`}
              data-testid="text-test-result"
            >
              {testResult.success ? (
                <CheckCircle className="w-4 h-4 mt-0.5 flex-shrink-0" />
              ) : (
                <XCircle className="w-4 h-4 mt-0.5 flex-shrink-0" />
              )}
              {testResult.message}
            </div>
          )}
        </form>

        {/* ── Document Sync ────────────────────────────────────────────────── */}
        <div className="rounded-xl border border-border bg-card p-6">
          <div className="flex items-start justify-between gap-4">
            <div>
              <h2 className="text-sm font-semibold text-foreground">Sync Document Library</h2>
              <p className="text-xs text-muted-foreground mt-1">
                Fetches all files from your configured library and loads them into the chatbot.
                Supports Word (.docx), PDF, and plain text files. Previous synced documents will be replaced.
              </p>
              {config?.lastSyncedAt && (
                <div className="flex items-center gap-1.5 mt-2 text-xs text-muted-foreground">
                  <Clock className="w-3 h-3" />
                  Last synced: {new Date(config.lastSyncedAt).toLocaleString()}
                </div>
              )}
            </div>
            <Button
              onClick={() => syncDocuments.mutate()}
              disabled={syncDocuments.isPending || !config}
              className="flex-shrink-0"
              data-testid="button-sync"
            >
              <RefreshCw className={`w-4 h-4 mr-2 ${syncDocuments.isPending ? "animate-spin" : ""}`} />
              {syncDocuments.isPending ? "Syncing…" : "Sync Now"}
            </Button>
          </div>
          {!config && (
            <div className="flex items-center gap-2 mt-4 text-xs text-amber-600 dark:text-amber-400">
              <AlertTriangle className="w-3.5 h-3.5" />
              Save your SharePoint configuration before syncing.
            </div>
          )}
        </div>

        {/* Deployment notice */}
        <div className="rounded-lg bg-muted/40 border border-border px-4 py-3 text-xs text-muted-foreground">
          <strong className="text-foreground">On-premises mode:</strong> The server must have direct network access to your SharePoint environment. Authentication uses NTLM with Active Directory credentials.{" "}
          <strong className="text-foreground">SharePoint Online mode:</strong> Requires an Azure AD app registration with Sites.Read.All permission and admin consent.
        </div>

      </div>
    </div>
  );
}
