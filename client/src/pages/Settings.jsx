import { useState } from "react";
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
} from "lucide-react";
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

const sharepointSchema = insertSharepointConfigSchema.extend({
  siteUrl: z.string().url("Must be a valid URL (e.g. http://sharepoint.company.com/sites/mysite)"),
  domain: z.string().min(1, "Domain is required"),
  username: z.string().min(1, "Username is required"),
  password: z.string().min(1, "Password is required"),
  libraryName: z.string().min(1, "Library name is required"),
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
      siteUrl: "",
      domain: "",
      username: "",
      password: "",
      libraryName: "Documents",
      allowSelfSigned: true,
    },
    values: config
      ? {
          siteUrl: config.siteUrl,
          domain: config.domain,
          username: config.username,
          password: config.password,
          libraryName: config.libraryName,
          allowSelfSigned: config.allowSelfSigned,
        }
      : undefined,
  });

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
    mutationFn: () => apiRequest("POST", "/api/sharepoint/test"),
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
                Do not include a not-found or fallback response here — use the <strong>Not Found Message</strong> field above, which always takes priority.
              </p>
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

            <div className="space-y-1">
              <Label htmlFor="siteUrl">SharePoint Site URL</Label>
              <Input
                id="siteUrl"
                placeholder="http://sharepoint.company.com/sites/mysite"
                data-testid="input-site-url"
                {...sharepointForm.register("siteUrl")}
              />
              {sharepointForm.formState.errors.siteUrl && (
                <p className="text-xs text-destructive">{sharepointForm.formState.errors.siteUrl.message}</p>
              )}
              <p className="text-xs text-muted-foreground">Full URL to your SharePoint site (not just the root)</p>
            </div>

            <div className="grid grid-cols-2 gap-4">
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

        {/* On-premises notice */}
        <div className="rounded-lg bg-muted/40 border border-border px-4 py-3 text-xs text-muted-foreground">
          <strong className="text-foreground">On-premises deployment:</strong> This application must be hosted on a server
          that has direct network access to your SharePoint 2019 environment. Authentication uses NTLM with
          your Active Directory credentials.
        </div>

      </div>
    </div>
  );
}
