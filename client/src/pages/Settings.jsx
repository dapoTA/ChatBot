import { useState, useMemo, useEffect, useRef } from "react";
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
  Download,
  Plus,
  Pencil,
  Trash2,
  ChevronDown,
  ChevronUp,
  ShieldCheck,
  LayoutDashboard,
  Sparkles,
  FolderOpen,
  SlidersHorizontal,
  MessageSquare,
  Activity,
  LockKeyhole,
  BarChart3,
  ChevronRight,
} from "lucide-react";
import { extractStyledPhrases, extractGlobalResponseStyle } from "@/lib/styleParser";
import { insertSharepointConfigSchema, insertAppSettingsSchema } from "@shared/schema";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";

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

const EMPTY_SOURCE = {
  name: "",
  libraryName: "",
  description: "",
  instructions: "",
  smeTeam: "",
  contactMethod: "",
  contactDetails: "",
  escalationMessage: "",
  enabled: true,
  isPortalWide: false,
};

const ASSISTANT_ICON_OPTIONS = [
  { value: "message-circle", label: "Chat", Icon: MessageSquare },
  { value: "sparkles", label: "Sparkles", Icon: Sparkles },
  { value: "shield-check", label: "Assurance", Icon: ShieldCheck },
];

const THEME_OPTIONS = {
  teal: { label: "Assurance Teal", primary: "#188b6a", surface: "#edf8f4", text: "#0f513f" },
  ocean: { label: "Portal Blue", primary: "#2563eb", surface: "#eff6ff", text: "#1e3a8a" },
  slate: { label: "Executive Slate", primary: "#475569", surface: "#f1f5f9", text: "#1e293b" },
  amber: { label: "Warm Amber", primary: "#b45309", surface: "#fffbeb", text: "#78350f" },
};

export default function Settings() {
  const { toast } = useToast();
  const [testResult, setTestResult] = useState(null);
  const [enableChatLog, setEnableChatLog] = useState(false);
  const [logFrom, setLogFrom] = useState("");
  const [logTo, setLogTo] = useState("");
  const [activeTab, setActiveTab] = useState("general");

  // ─── Fetch current settings ────────────────────────────────────────────────

  const { data: appSettingsData } = useQuery({
    queryKey: ["/api/settings"],
  });

  const { data: config } = useQuery({
    queryKey: ["/api/sharepoint/config"],
  });

  const { data: knowledgeSources = [] } = useQuery({
    queryKey: ["/api/knowledge-sources"],
  });

  const [editingSourceId, setEditingSourceId] = useState(null);
  const [sourceDraft, setSourceDraft] = useState(EMPTY_SOURCE);
  const [sourceFormError, setSourceFormError] = useState("");

  const openSourceEditor = (source) => {
    setEditingSourceId(source.id);
    setSourceDraft({
      ...EMPTY_SOURCE,
      ...source,
      libraryName: source.libraryName ?? "",
      description: source.description ?? "",
      instructions: source.instructions ?? "",
      smeTeam: source.smeTeam ?? "",
      contactMethod: source.contactMethod ?? "",
      contactDetails: source.contactDetails ?? "",
      escalationMessage: source.escalationMessage ?? "",
    });
    setSourceFormError("");
  };

  const openNewSourceEditor = () => {
    setEditingSourceId("new");
    setSourceDraft({ ...EMPTY_SOURCE });
    setSourceFormError("");
  };

  const closeSourceEditor = () => {
    setEditingSourceId(null);
    setSourceDraft({ ...EMPTY_SOURCE });
    setSourceFormError("");
  };

  const updateSourceDraft = (field, value) => {
    setSourceDraft((current) => ({ ...current, [field]: value }));
  };

  const saveKnowledgeSource = useMutation({
    mutationFn: async ({ id, data }) => {
      const path = id === "new"
        ? "/api/knowledge-sources"
        : `/api/knowledge-sources/${id}`;
      const response = await apiRequest(id === "new" ? "POST" : "PATCH", path, data);
      return response.status === 204 ? null : response.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/knowledge-sources"] });
      toast({
        title: editingSourceId === "new" ? "Source added" : "Source saved",
        description: "Knowledge source settings have been updated.",
      });
      closeSourceEditor();
    },
    onError: (error) => {
      setSourceFormError(error?.message || "Could not save this knowledge source.");
    },
  });

  const deleteKnowledgeSource = useMutation({
    mutationFn: (id) => apiRequest("DELETE", `/api/knowledge-sources/${id}`),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/knowledge-sources"] });
      toast({ title: "Source removed" });
      if (editingSourceId !== null) closeSourceEditor();
    },
    onError: (error) => {
      toast({
        title: "Remove failed",
        description: error?.message || "Could not remove this knowledge source.",
        variant: "destructive",
      });
    },
  });

  const toggleKnowledgeSource = useMutation({
    mutationFn: ({ id, enabled }) =>
      apiRequest("PATCH", `/api/knowledge-sources/${id}`, { enabled }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/knowledge-sources"] });
    },
    onError: (error) => {
      toast({
        title: "Status update failed",
        description: error?.message || "Could not update this knowledge source.",
        variant: "destructive",
      });
    },
  });

  const handleSaveKnowledgeSource = (event) => {
    event.preventDefault();
    const name = sourceDraft.name.trim();
    const libraryName = sourceDraft.libraryName.trim();

    if (!name) {
      setSourceFormError("Source name is required.");
      return;
    }
    if (!sourceDraft.isPortalWide && !libraryName) {
      setSourceFormError("SharePoint library name is required for a named source.");
      return;
    }

    saveKnowledgeSource.mutate({
      id: editingSourceId,
      data: {
        name,
        libraryName: sourceDraft.isPortalWide ? null : libraryName,
        description: sourceDraft.description.trim(),
        instructions: sourceDraft.instructions.trim(),
        smeTeam: sourceDraft.smeTeam.trim(),
        contactMethod: sourceDraft.contactMethod.trim(),
        contactDetails: sourceDraft.contactDetails.trim(),
        escalationMessage: sourceDraft.escalationMessage.trim(),
        enabled: sourceDraft.isPortalWide ? true : Boolean(sourceDraft.enabled),
        isPortalWide: Boolean(sourceDraft.isPortalWide),
      },
    });
  };

  const handleToggleKnowledgeSource = (source, enabled) => {
    toggleKnowledgeSource.mutate({ id: source.id, enabled });
  };

  const handleDeleteKnowledgeSource = (source) => {
    if (window.confirm(`Remove the "${source.name}" knowledge source?`)) {
      deleteKnowledgeSource.mutate(source.id);
    }
  };

  // ─── Chat logging toggle ───────────────────────────────────────────────────

  const logToggleLoaded = useRef(false);
  useEffect(() => {
    if (appSettingsData && !logToggleLoaded.current) {
      logToggleLoaded.current = true;
      setEnableChatLog(appSettingsData.enableChatLog ?? false);
    }
  }, [appSettingsData]);

  const saveLogToggle = useMutation({
    mutationFn: (enabled) => apiRequest("POST", "/api/settings", { enableChatLog: enabled }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/settings"] });
      toast({ title: enableChatLog ? "Chat logging enabled" : "Chat logging disabled" });
    },
    onError: () => {
      toast({ title: "Save failed", description: "Could not update logging setting.", variant: "destructive" });
    },
  });

  const handleLogToggle = (v) => {
    setEnableChatLog(v);
    saveLogToggle.mutate(v);
  };

  const handleExportLogs = () => {
    const params = new URLSearchParams();
    if (logFrom) params.set("from", logFrom);
    if (logTo)   params.set("to",   logTo);
    const url = `/api/chat-logs${params.toString() ? `?${params}` : ""}`;
    // Use a temporary anchor to trigger a file download
    const a = document.createElement("a");
    a.href = url;
    a.download = "";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // ─── Appearance form ───────────────────────────────────────────────────────

  const appearanceForm = useForm({
    resolver: zodResolver(appearanceSchema),
    defaultValues: {
      assistantName: "ON-PNT® Assistant",
      welcomeMessage: "Ask me anything about your SharePoint documents.",
      notFoundMessage: "I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.",
      customInstructions: "",
      assistantIcon: "message-circle",
      theme: "teal",
      launcherLabel: "Ask inSite",
      launcherPosition: "bottom-right",
      launcherStyle: "bubble",
    },
  });

  // Load server data once on mount — using values prop would re-sync on every
  // refetch and overwrite in-progress user edits in the textarea.
  const appearanceLoaded = useRef(false);
  useEffect(() => {
    if (appSettingsData && !appearanceLoaded.current) {
      appearanceLoaded.current = true;
      appearanceForm.reset({
        assistantName: appSettingsData.assistantName,
        welcomeMessage: appSettingsData.welcomeMessage,
        notFoundMessage: appSettingsData.notFoundMessage,
        customInstructions: appSettingsData.customInstructions ?? "",
        assistantIcon: appSettingsData.assistantIcon ?? "message-circle",
        theme: appSettingsData.theme ?? "teal",
        launcherLabel: appSettingsData.launcherLabel ?? "Ask inSite",
        launcherPosition: appSettingsData.launcherPosition ?? "bottom-right",
        launcherStyle: appSettingsData.launcherStyle ?? "bubble",
      });
    }
  }, [appSettingsData]);

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
  });

  const aiParamsLoaded = useRef(false);
  useEffect(() => {
    if (appSettingsData && !aiParamsLoaded.current) {
      aiParamsLoaded.current = true;
      aiParamsForm.reset({
        temperature: appSettingsData.temperature ?? 0,
        topP: appSettingsData.topP ?? 1,
        maxTokens: appSettingsData.maxTokens ?? 1500,
        frequencyPenalty: appSettingsData.frequencyPenalty ?? 0,
        presencePenalty: appSettingsData.presencePenalty ?? 0,
      });
    }
  }, [appSettingsData]);

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
  const selectedTheme = THEME_OPTIONS[appearanceForm.watch("theme")] ?? THEME_OPTIONS.teal;
  const selectedIcon = ASSISTANT_ICON_OPTIONS.find(
    (option) => option.value === appearanceForm.watch("assistantIcon"),
  ) ?? ASSISTANT_ICON_OPTIONS[0];
  const PreviewIcon = selectedIcon.Icon;
  const launcherLabel = appearanceForm.watch("launcherLabel") || "Ask inSite";
  const launcherStyle = appearanceForm.watch("launcherStyle") || "bubble";
  const launcherPosition = appearanceForm.watch("launcherPosition") || "bottom-right";

  const handleSaveActiveTab = () => {
    if (activeTab === "general") {
      appearanceForm.handleSubmit((data) => saveAppearance.mutate(data))();
      return;
    }
    if (activeTab === "chat-behavior") {
      aiParamsForm.handleSubmit((data) => saveAiParams.mutate(data))();
      return;
    }
    if (activeTab === "knowledge-sources") {
      if (editingSourceId !== null) {
        handleSaveKnowledgeSource({ preventDefault: () => {} });
      } else {
        sharepointForm.handleSubmit((data) => saveSharepoint.mutate(data))();
      }
    }
  };

  const handleDiscardActiveTab = () => {
    if (activeTab === "general" && appSettingsData) {
      appearanceForm.reset({
        assistantName: appSettingsData.assistantName,
        welcomeMessage: appSettingsData.welcomeMessage,
        notFoundMessage: appSettingsData.notFoundMessage,
        customInstructions: appSettingsData.customInstructions ?? "",
        assistantIcon: appSettingsData.assistantIcon ?? "message-circle",
        theme: appSettingsData.theme ?? "teal",
        launcherLabel: appSettingsData.launcherLabel ?? "Ask inSite",
        launcherPosition: appSettingsData.launcherPosition ?? "bottom-right",
        launcherStyle: appSettingsData.launcherStyle ?? "bubble",
      });
    }
    if (activeTab === "chat-behavior" && appSettingsData) {
      aiParamsForm.reset({
        temperature: appSettingsData.temperature ?? 0,
        topP: appSettingsData.topP ?? 1,
        maxTokens: appSettingsData.maxTokens ?? 1500,
        frequencyPenalty: appSettingsData.frequencyPenalty ?? 0,
        presencePenalty: appSettingsData.presencePenalty ?? 0,
      });
    }
    if (activeTab === "knowledge-sources") {
      closeSourceEditor();
      setTestResult(null);
      if (config) {
        sharepointForm.reset({
          mode: config.mode ?? "onprem",
          siteUrl: config.siteUrl,
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
        });
      }
    }
  };

  const activeSavePending =
    (activeTab === "general" && saveAppearance.isPending)
    || (activeTab === "chat-behavior" && saveAiParams.isPending)
    || (activeTab === "knowledge-sources" && (saveKnowledgeSource.isPending || saveSharepoint.isPending));
  const activeHasUnsavedChanges =
    (activeTab === "general" && appearanceForm.formState.isDirty)
    || (activeTab === "chat-behavior" && aiParamsForm.formState.isDirty)
    || (activeTab === "knowledge-sources" && (sharepointForm.formState.isDirty || editingSourceId !== null));

  return (
    <div className="min-h-screen bg-[#eef4f8] text-foreground">
      <header className="h-16 bg-[#0b2942] text-white flex items-center justify-between px-5 md:px-8 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#1e9b75] shadow-inner">
            <ShieldCheck className="h-5 w-5" />
          </div>
          <div className="leading-tight">
            <p className="text-[9px] font-semibold tracking-[0.24em] text-[#8fc9b4] uppercase">Technical Assurance</p>
            <p className="text-sm font-semibold tracking-tight">inSite Assistant</p>
          </div>
        </div>
        <div className="flex items-center gap-2 text-xs text-[#c7d9e5]">
          <span className="h-1.5 w-1.5 rounded-full bg-[#61d39c]" />
          Configuration
        </div>
      </header>

      <div className="flex min-h-[calc(100vh-4rem)]">
        <aside className="hidden w-60 shrink-0 border-r border-[#dce7ed] bg-[#f7fafc] px-4 py-7 lg:block">
          <div className="px-3">
            <p className="text-[10px] font-semibold tracking-[0.18em] text-muted-foreground uppercase">Administration</p>
            <p className="mt-3 text-xs leading-relaxed text-muted-foreground">
              Keep the assistant aligned with your portal and assurance standards.
            </p>
          </div>
          <nav className="mt-8 space-y-1" aria-label="Administration">
            <a href="#" className="flex items-center gap-3 rounded-md px-3 py-2.5 text-xs text-muted-foreground hover:bg-white hover:text-foreground">
              <LayoutDashboard className="h-4 w-4" />
              Overview
            </a>
            <a href="#assistant-settings" className="flex items-center gap-3 rounded-md bg-white px-3 py-2.5 text-xs font-semibold text-[#188b6a] shadow-sm ring-1 ring-[#dfe9e8]">
              <Sparkles className="h-4 w-4" />
              Assistant
            </a>
            <a href="#" className="flex items-center gap-3 rounded-md px-3 py-2.5 text-xs text-muted-foreground hover:bg-white hover:text-foreground">
              <SlidersHorizontal className="h-4 w-4" />
              Experience
            </a>
            <a href="#" className="flex items-center gap-3 rounded-md px-3 py-2.5 text-xs text-muted-foreground hover:bg-white hover:text-foreground">
              <Activity className="h-4 w-4" />
              Usage
            </a>
          </nav>
          <div className="my-7 border-t border-[#dce7ed]" />
          <div className="px-3">
            <p className="text-[10px] font-semibold tracking-[0.18em] text-muted-foreground uppercase">Environment</p>
            <div className="mt-4 space-y-3 text-xs text-muted-foreground">
              <div className="flex items-center gap-3 text-[#188b6a]">
                <FolderOpen className="h-4 w-4" />
                Production portal
              </div>
              <div className="flex items-center gap-3">
                <LockKeyhole className="h-4 w-4" />
                Admin access only
              </div>
            </div>
          </div>
        </aside>

        <main className="min-w-0 flex-1 px-4 py-8 md:px-8 lg:px-12">
          <div className="mx-auto max-w-5xl space-y-6">
            <div id="assistant-settings" className="flex items-end justify-between gap-4">
              <div>
                <div className="mb-2 flex items-center gap-2 text-[11px] text-muted-foreground">
                  <span>Assistant</span>
                  <ChevronRight className="h-3 w-3" />
                  <span className="font-medium text-[#188b6a]">Settings</span>
                </div>
                <h1 className="text-2xl font-semibold tracking-tight text-[#102a43]" data-testid="text-settings-title">
                  Assistant settings
                </h1>
                <p className="mt-1 text-sm text-muted-foreground">
                  Control how inSite Assistant is presented, grounded, and maintained.
                </p>
              </div>
              <div className="hidden items-center gap-2 pb-1 text-xs text-muted-foreground sm:flex">
                <CheckCircle className={`h-3.5 w-3.5 ${activeHasUnsavedChanges ? "text-amber-500" : "text-[#1e9b75]"}`} />
                {activeHasUnsavedChanges ? "Unsaved changes" : "All changes saved"}
              </div>
            </div>

            <Tabs value={activeTab} onValueChange={setActiveTab} className="overflow-hidden rounded-xl border border-[#dce7ed] bg-white shadow-sm">
              <TabsList className="grid h-auto w-full grid-cols-2 gap-0 rounded-none border-b border-[#e2ebef] bg-white p-0 sm:grid-cols-4">
                <TabsTrigger value="general" className="h-16 rounded-none border-b-2 border-transparent px-3 text-left text-xs text-muted-foreground data-[state=active]:border-[#1e9b75] data-[state=active]:bg-[#fbfefd] data-[state=active]:text-[#167b60] data-[state=active]:shadow-none">
                  <span className="flex items-center gap-2.5">
                    <SlidersHorizontal className="h-4 w-4" />
                    <span><strong className="block font-semibold">General</strong><span className="hidden text-[10px] font-normal opacity-70 md:block">Branding and presentation</span></span>
                  </span>
                </TabsTrigger>
                <TabsTrigger value="knowledge-sources" className="h-16 rounded-none border-b-2 border-transparent px-3 text-left text-xs text-muted-foreground data-[state=active]:border-[#1e9b75] data-[state=active]:bg-[#fbfefd] data-[state=active]:text-[#167b60] data-[state=active]:shadow-none">
                  <span className="flex items-center gap-2.5">
                    <FolderOpen className="h-4 w-4" />
                    <span><strong className="block font-semibold">Knowledge Sources</strong><span className="hidden text-[10px] font-normal opacity-70 md:block">Where answers come from</span></span>
                  </span>
                </TabsTrigger>
                <TabsTrigger value="chat-behavior" className="h-16 rounded-none border-b-2 border-transparent px-3 text-left text-xs text-muted-foreground data-[state=active]:border-[#1e9b75] data-[state=active]:bg-[#fbfefd] data-[state=active]:text-[#167b60] data-[state=active]:shadow-none">
                  <span className="flex items-center gap-2.5">
                    <MessageSquare className="h-4 w-4" />
                    <span><strong className="block font-semibold">Chat Behavior</strong><span className="hidden text-[10px] font-normal opacity-70 md:block">Response guardrails</span></span>
                  </span>
                </TabsTrigger>
                <TabsTrigger value="logging-exports" className="h-16 rounded-none border-b-2 border-transparent px-3 text-left text-xs text-muted-foreground data-[state=active]:border-[#1e9b75] data-[state=active]:bg-[#fbfefd] data-[state=active]:text-[#167b60] data-[state=active]:shadow-none">
                  <span className="flex items-center gap-2.5">
                    <BarChart3 className="h-4 w-4" />
                    <span><strong className="block font-semibold">Logging &amp; Exports</strong><span className="hidden text-[10px] font-normal opacity-70 md:block">Records and retention</span></span>
                  </span>
                </TabsTrigger>
              </TabsList>
              <div className="p-4 sm:p-6">

        {/* ── Widget Appearance ───────────────────────────────────────────── */}
        <div
          className={activeTab === "general" ? "block" : "hidden"}
          role="tabpanel"
          aria-label="General settings"
          data-testid="settings-panel-general"
        >
        <form onSubmit={appearanceForm.handleSubmit((d) => saveAppearance.mutate(d))}>
          <div className="rounded-xl border border-border bg-card p-6 space-y-5">
            <div>
              <p className="text-[10px] font-semibold uppercase tracking-[0.16em] text-[#188b6a]">General presentation</p>
              <h2 className="mt-1 text-lg font-semibold text-[#102a43]">Introduce the assistant clearly</h2>
              <p className="mt-1 text-xs text-muted-foreground">
                Set the language administrators and portal users see before and during a conversation.
              </p>
            </div>

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

          </div>

          <div className="mt-6 rounded-xl border border-border bg-card p-6 space-y-5" data-testid="theme-launcher-section">
            <div>
              <p className="text-[10px] font-semibold uppercase tracking-[0.16em] text-[#188b6a]">Theme &amp; launcher</p>
              <h2 className="mt-1 text-lg font-semibold text-[#102a43]">Match the portal experience</h2>
              <p className="mt-1 text-xs text-muted-foreground">
                Choose the assistant mark, accent theme, launcher wording, placement, and shape.
              </p>
            </div>

            <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
              <div className="space-y-2">
                <Label>Assistant icon</Label>
                <div className="grid grid-cols-3 gap-2">
                  {ASSISTANT_ICON_OPTIONS.map(({ value, label, Icon }) => {
                    const isSelected = appearanceForm.watch("assistantIcon") === value;
                    return (
                      <button
                        key={value}
                        type="button"
                        onClick={() => appearanceForm.setValue("assistantIcon", value, { shouldDirty: true })}
                        className={`flex min-h-20 flex-col items-center justify-center gap-2 rounded-lg border px-3 py-3 text-xs font-medium transition-colors ${
                          isSelected
                            ? "border-[#188b6a] bg-[#edf8f4] text-[#126b52] ring-1 ring-[#188b6a]/20"
                            : "border-border bg-white text-muted-foreground hover:border-[#a9c9bd] hover:bg-[#fbfefd]"
                        }`}
                        aria-pressed={isSelected}
                        data-testid={`button-assistant-icon-${value}`}
                      >
                        <Icon className="h-5 w-5" />
                        {label}
                      </button>
                    );
                  })}
                </div>
              </div>

              <div className="space-y-2">
                <Label>Accent theme</Label>
                <div className="grid grid-cols-2 gap-2">
                  {Object.entries(THEME_OPTIONS).map(([value, theme]) => {
                    const isSelected = appearanceForm.watch("theme") === value;
                    return (
                      <button
                        key={value}
                        type="button"
                        onClick={() => appearanceForm.setValue("theme", value, { shouldDirty: true })}
                        className={`flex items-center gap-3 rounded-lg border px-3 py-3 text-left text-xs font-medium transition-colors ${
                          isSelected ? "border-[#188b6a] ring-1 ring-[#188b6a]/20" : "border-border hover:bg-muted/30"
                        }`}
                        aria-pressed={isSelected}
                        data-testid={`button-theme-${value}`}
                      >
                        <span className="h-7 w-7 rounded-full border-4 border-white shadow-sm" style={{ backgroundColor: theme.primary }} />
                        <span>{theme.label}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-4 md:grid-cols-3">
              <div className="space-y-1">
                <Label htmlFor="launcherLabel">Launcher label</Label>
                <Input
                  id="launcherLabel"
                  maxLength={40}
                  placeholder="Ask inSite"
                  {...appearanceForm.register("launcherLabel")}
                  data-testid="input-launcher-label"
                />
                <p className="text-xs text-muted-foreground">Shown when the pill launcher is selected.</p>
              </div>
              <div className="space-y-1">
                <Label htmlFor="launcherStyle">Launcher style</Label>
                <select
                  id="launcherStyle"
                  {...appearanceForm.register("launcherStyle")}
                  className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-sm shadow-sm outline-none focus:ring-1 focus:ring-ring"
                  data-testid="select-launcher-style"
                >
                  <option value="bubble">Icon bubble</option>
                  <option value="pill">Icon with label</option>
                </select>
              </div>
              <div className="space-y-1">
                <Label htmlFor="launcherPosition">Screen position</Label>
                <select
                  id="launcherPosition"
                  {...appearanceForm.register("launcherPosition")}
                  className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-sm shadow-sm outline-none focus:ring-1 focus:ring-ring"
                  data-testid="select-launcher-position"
                >
                  <option value="bottom-right">Bottom right</option>
                  <option value="bottom-left">Bottom left</option>
                </select>
              </div>
            </div>

            <div className="overflow-hidden rounded-xl border border-[#dbe7eb] bg-[#edf3f6]" data-testid="launcher-preview">
              <div className="flex items-center justify-between border-b border-[#dbe7eb] bg-white px-4 py-3">
                <div>
                  <p className="text-xs font-semibold text-[#102a43]">Live launcher preview</p>
                  <p className="text-[10px] text-muted-foreground">Updates before you save</p>
                </div>
                <span className="text-[10px] font-medium uppercase tracking-wide text-muted-foreground">
                  {launcherPosition === "bottom-left" ? "Bottom left" : "Bottom right"}
                </span>
              </div>
              <div className={`flex min-h-44 items-end p-5 ${launcherPosition === "bottom-left" ? "justify-start" : "justify-end"}`}>
                <div className="space-y-3">
                  <div className="w-64 overflow-hidden rounded-xl bg-white shadow-lg ring-1 ring-black/5">
                    <div className="flex items-center gap-2 px-3 py-2.5 text-white" style={{ backgroundColor: selectedTheme.primary }}>
                      <span className="flex h-7 w-7 items-center justify-center rounded-full bg-white/15">
                        <PreviewIcon className="h-4 w-4" />
                      </span>
                      <div>
                        <p className="text-xs font-semibold">{appearanceForm.watch("assistantName") || "inSite Assistant"}</p>
                        <p className="text-[9px] text-white/75">Ready to help</p>
                      </div>
                    </div>
                    <div className="p-3 text-[10px]" style={{ backgroundColor: selectedTheme.surface, color: selectedTheme.text }}>
                      {appearanceForm.watch("welcomeMessage") || "Ask me anything about your portal documents."}
                    </div>
                  </div>
                  <div className={`flex ${launcherPosition === "bottom-left" ? "justify-start" : "justify-end"}`}>
                    <div
                      className={`flex items-center justify-center gap-2 text-white shadow-lg ${
                        launcherStyle === "pill" ? "min-h-11 rounded-full px-4" : "h-12 w-12 rounded-full"
                      }`}
                      style={{ backgroundColor: selectedTheme.primary }}
                    >
                      <PreviewIcon className="h-5 w-5" />
                      {launcherStyle === "pill" && <span className="text-xs font-semibold">{launcherLabel}</span>}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </form>
        </div>

        {/* ── Knowledge Sources ───────────────────────────────────────────── */}
        <div
          className={activeTab === "knowledge-sources" ? "block space-y-6" : "hidden"}
          role="tabpanel"
          aria-label="Knowledge Sources settings"
          data-testid="settings-panel-knowledge-sources"
        >
        <section id="knowledge-sources" className="rounded-xl border border-border bg-card p-6 space-y-5" data-testid="knowledge-sources-section">
          <div className="flex items-start justify-between gap-4">
            <div>
              <h2 className="text-sm font-semibold text-foreground uppercase tracking-wide">
                Knowledge Sources
              </h2>
              <p className="text-xs text-muted-foreground mt-1">
                Choose which SharePoint knowledge area the assistant should use. Global Custom Instructions still apply to every source.
              </p>
            </div>
            <Button type="button" variant="outline" onClick={openNewSourceEditor} data-testid="button-add-source">
              <Plus className="w-4 h-4 mr-2" />
              Add Source
            </Button>
          </div>

          <div className="space-y-2">
            {knowledgeSources.length === 0 ? (
              <div className="rounded-lg border border-dashed border-border px-4 py-5 text-sm text-muted-foreground" data-testid="knowledge-sources-empty">
                No knowledge sources have been configured yet. Add a named SharePoint library to get started.
              </div>
            ) : (
              knowledgeSources.map((source) => {
                const isEditing = editingSourceId === source.id;
                return (
                  <div
                    key={source.id}
                    className={`rounded-lg border transition-colors ${isEditing ? "border-primary/60 bg-primary/5" : "border-border bg-muted/20"}`}
                    data-testid={`knowledge-source-${source.id}`}
                  >
                    <div className="flex items-start gap-3 p-4">
                      <button
                        type="button"
                        className="flex-1 min-w-0 text-left"
                        onClick={() => openSourceEditor(source)}
                        aria-expanded={isEditing}
                        data-testid={`button-expand-source-${source.id}`}
                      >
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="text-sm font-semibold text-foreground">{source.name}</span>
                          <span className={`text-[10px] font-medium uppercase tracking-wide px-1.5 py-0.5 rounded ${
                            source.enabled
                              ? "bg-green-100 dark:bg-green-900 text-green-700 dark:text-green-300"
                              : "bg-muted text-muted-foreground"
                          }`}>
                            {source.enabled ? "Enabled" : "Disabled"}
                          </span>
                          {source.isPortalWide && (
                            <span className="text-[10px] font-medium uppercase tracking-wide px-1.5 py-0.5 rounded bg-blue-100 dark:bg-blue-900 text-blue-700 dark:text-blue-300">
                              Portal-wide
                            </span>
                          )}
                        </div>
                        <p className="text-xs text-muted-foreground mt-1 truncate">
                          {source.isPortalWide
                            ? "Searches the full configured SharePoint site collection."
                            : `Library: ${source.libraryName || "Not configured"}`}
                        </p>
                        {source.description && (
                          <p className="text-xs text-muted-foreground mt-1 line-clamp-2">{source.description}</p>
                        )}
                      </button>
                      <div className="flex items-center gap-1 shrink-0">
                        {!source.isPortalWide && (
                          <Switch
                            checked={Boolean(source.enabled)}
                            onCheckedChange={(enabled) => handleToggleKnowledgeSource(source, enabled)}
                            disabled={toggleKnowledgeSource.isPending}
                            aria-label={`${source.enabled ? "Disable" : "Enable"} ${source.name}`}
                            data-testid={`switch-source-${source.id}`}
                          />
                        )}
                        <Button
                          type="button"
                          variant="ghost"
                          size="icon"
                          onClick={() => isEditing ? closeSourceEditor() : openSourceEditor(source)}
                          aria-label={`Edit ${source.name}`}
                          data-testid={`button-edit-source-${source.id}`}
                        >
                          <Pencil className="w-4 h-4" />
                        </Button>
                        <Button
                          type="button"
                          variant="ghost"
                          size="icon"
                          onClick={() => openSourceEditor(source)}
                          aria-label={`${isEditing ? "Collapse" : "Expand"} ${source.name}`}
                          data-testid={`button-toggle-source-${source.id}`}
                        >
                          {isEditing ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                        </Button>
                        {!source.isPortalWide && (
                          <Button
                            type="button"
                            variant="ghost"
                            size="icon"
                            onClick={() => handleDeleteKnowledgeSource(source)}
                            disabled={deleteKnowledgeSource.isPending}
                            aria-label={`Remove ${source.name}`}
                            data-testid={`button-delete-source-${source.id}`}
                          >
                            <Trash2 className="w-4 h-4 text-destructive" />
                          </Button>
                        )}
                      </div>
                    </div>

                    {isEditing && (
                      <form onSubmit={handleSaveKnowledgeSource} className="border-t border-border p-4 space-y-4" data-testid={`source-editor-${source.id}`}>
                        <div className="flex items-center justify-between gap-3">
                          <div>
                            <h3 className="text-sm font-semibold text-foreground">Edit Source</h3>
                            <p className="text-xs text-muted-foreground mt-0.5">
                              Keep advanced response and escalation guidance with the source it serves.
                            </p>
                          </div>
                          <Button type="button" variant="ghost" size="sm" onClick={closeSourceEditor}>
                            Close
                          </Button>
                        </div>

                        {source.isPortalWide ? (
                          <div className="rounded-lg border border-blue-200 dark:border-blue-800 bg-blue-50 dark:bg-blue-950 px-3 py-2 text-xs text-blue-800 dark:text-blue-200">
                            <strong>All Portal Sources</strong> searches the full configured SharePoint site collection. Its guidance is separate from every individual library.
                          </div>
                        ) : (
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div className="space-y-1">
                              <Label htmlFor={`source-name-${source.id}`}>Source Name</Label>
                              <Input
                                id={`source-name-${source.id}`}
                                value={sourceDraft.name}
                                onChange={(event) => updateSourceDraft("name", event.target.value)}
                                placeholder="HR"
                                data-testid={`input-source-name-${source.id}`}
                              />
                            </div>
                            <div className="space-y-1">
                              <Label htmlFor={`source-library-${source.id}`}>SharePoint Library</Label>
                              <Input
                                id={`source-library-${source.id}`}
                                value={sourceDraft.libraryName}
                                onChange={(event) => updateSourceDraft("libraryName", event.target.value)}
                                placeholder="HR"
                                data-testid={`input-source-library-${source.id}`}
                              />
                            </div>
                          </div>
                        )}

                        {!source.isPortalWide && (
                          <div className="flex items-center justify-between rounded-lg border border-border bg-muted/30 px-3 py-2">
                            <div>
                              <p className="text-sm font-medium text-foreground">Source enabled</p>
                              <p className="text-xs text-muted-foreground">Disabled sources are hidden from the chat selector.</p>
                            </div>
                            <Switch
                              checked={Boolean(sourceDraft.enabled)}
                              onCheckedChange={(enabled) => updateSourceDraft("enabled", enabled)}
                              data-testid={`switch-source-editor-${source.id}`}
                            />
                          </div>
                        )}

                        <div className="space-y-1">
                          <Label htmlFor={`source-description-${source.id}`}>Description</Label>
                          <Textarea
                            id={`source-description-${source.id}`}
                            value={sourceDraft.description}
                            onChange={(event) => updateSourceDraft("description", event.target.value)}
                            placeholder="Short plain-language explanation shown to administrators and users."
                            rows={2}
                            className="resize-y"
                            data-testid={`input-source-description-${source.id}`}
                          />
                          <p className="text-xs text-muted-foreground">
                            Helper text only; it is not interpreted as an AI instruction.
                          </p>
                        </div>

                        <div className="space-y-1">
                          <Label htmlFor={`source-instructions-${source.id}`}>
                            {source.isPortalWide ? "Portal-wide Instructions" : "Instructions for This Source"}
                          </Label>
                          <Textarea
                            id={`source-instructions-${source.id}`}
                            value={sourceDraft.instructions}
                            onChange={(event) => updateSourceDraft("instructions", event.target.value)}
                            placeholder={source.isPortalWide
                              ? "Add guidance for questions that search the full portal."
                              : "Add guidance that applies only when this source is selected."}
                            rows={5}
                            className="resize-y font-mono text-xs"
                            data-testid={`input-source-instructions-${source.id}`}
                          />
                          <p className="text-xs text-muted-foreground">
                            Global Custom Instructions are added automatically. Only this source's guidance is added here.
                          </p>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          <div className="space-y-1">
                            <Label htmlFor={`source-sme-${source.id}`}>SME / Team</Label>
                            <Input
                              id={`source-sme-${source.id}`}
                              value={sourceDraft.smeTeam}
                              onChange={(event) => updateSourceDraft("smeTeam", event.target.value)}
                              placeholder="Human Resources"
                              data-testid={`input-source-sme-${source.id}`}
                            />
                          </div>
                          <div className="space-y-1">
                            <Label htmlFor={`source-contact-method-${source.id}`}>Contact Method</Label>
                            <Input
                              id={`source-contact-method-${source.id}`}
                              value={sourceDraft.contactMethod}
                              onChange={(event) => updateSourceDraft("contactMethod", event.target.value)}
                              placeholder="Email, Teams, phone"
                              data-testid={`input-source-contact-method-${source.id}`}
                            />
                          </div>
                        </div>

                        <div className="space-y-1">
                          <Label htmlFor={`source-contact-details-${source.id}`}>Contact Details</Label>
                          <Input
                            id={`source-contact-details-${source.id}`}
                            value={sourceDraft.contactDetails}
                            onChange={(event) => updateSourceDraft("contactDetails", event.target.value)}
                            placeholder="Mailbox, extension, Teams channel, or approved route"
                            data-testid={`input-source-contact-details-${source.id}`}
                          />
                        </div>

                        <div className="space-y-1">
                          <Label htmlFor={`source-escalation-${source.id}`}>
                            {source.isPortalWide ? "Portal-wide Escalation Message" : "Escalation Message"}
                          </Label>
                          <Textarea
                            id={`source-escalation-${source.id}`}
                            value={sourceDraft.escalationMessage}
                            onChange={(event) => updateSourceDraft("escalationMessage", event.target.value)}
                            placeholder="Tell the user what to do when this source cannot answer."
                            rows={3}
                            className="resize-y"
                            data-testid={`input-source-escalation-${source.id}`}
                          />
                          <p className="text-xs text-muted-foreground">
                            Used with the contact routing above when the selected source cannot answer.
                          </p>
                        </div>

                        {sourceFormError && (
                          <p className="text-xs text-destructive" role="alert" data-testid="source-form-error">
                            {sourceFormError}
                          </p>
                        )}

                        <div className="flex items-center gap-2">
                          <Button type="submit" disabled={saveKnowledgeSource.isPending} data-testid={`button-save-source-${source.id}`}>
                            <Save className="w-4 h-4 mr-2" />
                            {saveKnowledgeSource.isPending ? "Saving…" : "Save Source"}
                          </Button>
                          <Button type="button" variant="outline" onClick={closeSourceEditor}>
                            Cancel
                          </Button>
                        </div>
                      </form>
                    )}
                  </div>
                );
              })
            )}
          </div>

          {editingSourceId === "new" && (
            <form onSubmit={handleSaveKnowledgeSource} className="rounded-lg border border-primary/60 bg-primary/5 p-4 space-y-4" data-testid="source-editor-new">
              <div className="flex items-center justify-between gap-3">
                <div>
                  <h3 className="text-sm font-semibold text-foreground">Add Knowledge Source</h3>
                  <p className="text-xs text-muted-foreground mt-0.5">
                    Add a named source and connect it to the matching SharePoint library.
                  </p>
                </div>
                <Button type="button" variant="ghost" size="sm" onClick={closeSourceEditor}>Close</Button>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <Label htmlFor="new-source-name">Source Name</Label>
                  <Input
                    id="new-source-name"
                    value={sourceDraft.name}
                    onChange={(event) => updateSourceDraft("name", event.target.value)}
                    placeholder="Finance"
                    data-testid="input-new-source-name"
                  />
                </div>
                <div className="space-y-1">
                  <Label htmlFor="new-source-library">SharePoint Library</Label>
                  <Input
                    id="new-source-library"
                    value={sourceDraft.libraryName}
                    onChange={(event) => updateSourceDraft("libraryName", event.target.value)}
                    placeholder="Finance"
                    data-testid="input-new-source-library"
                  />
                </div>
              </div>

              <div className="space-y-1">
                <Label htmlFor="new-source-description">Description</Label>
                <Textarea
                  id="new-source-description"
                  value={sourceDraft.description}
                  onChange={(event) => updateSourceDraft("description", event.target.value)}
                  placeholder="Short plain-language explanation shown to administrators and users."
                  rows={2}
                  className="resize-y"
                  data-testid="input-new-source-description"
                />
                <p className="text-xs text-muted-foreground">Helper text only; it is not interpreted as an AI instruction.</p>
              </div>

              <div className="space-y-1">
                <Label htmlFor="new-source-instructions">Instructions for This Source</Label>
                <Textarea
                  id="new-source-instructions"
                  value={sourceDraft.instructions}
                  onChange={(event) => updateSourceDraft("instructions", event.target.value)}
                  placeholder="Add guidance that applies only when this source is selected."
                  rows={5}
                  className="resize-y font-mono text-xs"
                  data-testid="input-new-source-instructions"
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <Label htmlFor="new-source-sme">SME / Team</Label>
                  <Input
                    id="new-source-sme"
                    value={sourceDraft.smeTeam}
                    onChange={(event) => updateSourceDraft("smeTeam", event.target.value)}
                    placeholder="Finance team"
                    data-testid="input-new-source-sme"
                  />
                </div>
                <div className="space-y-1">
                  <Label htmlFor="new-source-contact-method">Contact Method</Label>
                  <Input
                    id="new-source-contact-method"
                    value={sourceDraft.contactMethod}
                    onChange={(event) => updateSourceDraft("contactMethod", event.target.value)}
                    placeholder="Email, Teams, phone"
                    data-testid="input-new-source-contact-method"
                  />
                </div>
              </div>

              <div className="space-y-1">
                <Label htmlFor="new-source-contact-details">Contact Details</Label>
                <Input
                  id="new-source-contact-details"
                  value={sourceDraft.contactDetails}
                  onChange={(event) => updateSourceDraft("contactDetails", event.target.value)}
                  placeholder="Mailbox, extension, Teams channel, or approved route"
                  data-testid="input-new-source-contact-details"
                />
              </div>

              <div className="space-y-1">
                <Label htmlFor="new-source-escalation">Escalation Message</Label>
                <Textarea
                  id="new-source-escalation"
                  value={sourceDraft.escalationMessage}
                  onChange={(event) => updateSourceDraft("escalationMessage", event.target.value)}
                  placeholder="Tell the user what to do when this source cannot answer."
                  rows={3}
                  className="resize-y"
                  data-testid="input-new-source-escalation"
                />
              </div>

              {sourceFormError && (
                <p className="text-xs text-destructive" role="alert" data-testid="source-form-error-new">
                  {sourceFormError}
                </p>
              )}

              <div className="flex items-center gap-2">
                <Button type="submit" disabled={saveKnowledgeSource.isPending} data-testid="button-save-new-source">
                  <Save className="w-4 h-4 mr-2" />
                  {saveKnowledgeSource.isPending ? "Saving…" : "Add Source"}
                </Button>
                <Button type="button" variant="outline" onClick={closeSourceEditor}>Cancel</Button>
              </div>
            </form>
          )}
        </section>
        </div>

        {/* ── AI Model Parameters ─────────────────────────────────────────── */}
        <div
          className={activeTab === "chat-behavior" ? "block" : "hidden"}
          role="tabpanel"
          aria-label="Chat Behavior settings"
          data-testid="settings-panel-chat-behavior"
        >
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

          </div>
        </form>
        </div>

        {/* ── Chat Logging ────────────────────────────────────────────────── */}
        <div
          className={activeTab === "logging-exports" ? "block" : "hidden"}
          role="tabpanel"
          aria-label="Logging and Exports settings"
          data-testid="settings-panel-logging-exports"
        >
        <div className="rounded-xl border border-border bg-card p-6 space-y-5">
          <div>
            <h2 className="text-sm font-semibold text-foreground uppercase tracking-wide">
              Chat Logging
            </h2>
            <p className="text-xs text-muted-foreground mt-1">
              When enabled, every conversation is recorded with the user's Windows/SharePoint login name.
              Intended for testing only — disable before final release.
            </p>
          </div>

          <div className="flex items-center justify-between rounded-lg border border-border bg-muted/30 px-4 py-3">
            <div>
              <p className="text-sm font-medium text-foreground">Enable chat logging</p>
              <p className="text-xs text-muted-foreground">
                Saves prompt, response, username, and session ID to the database
              </p>
            </div>
            <Switch
              checked={enableChatLog}
              onCheckedChange={handleLogToggle}
              disabled={saveLogToggle.isPending}
              data-testid="switch-chat-log"
            />
          </div>

          <div className="space-y-3">
            <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">
              Export logs as CSV
            </p>
            <div className="grid grid-cols-2 gap-3">
              <div className="space-y-1">
                <label className="text-xs text-muted-foreground">From date</label>
                <input
                  type="date"
                  value={logFrom}
                  onChange={(e) => setLogFrom(e.target.value)}
                  className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-sm shadow-sm transition-colors focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring"
                />
              </div>
              <div className="space-y-1">
                <label className="text-xs text-muted-foreground">To date</label>
                <input
                  type="date"
                  value={logTo}
                  onChange={(e) => setLogTo(e.target.value)}
                  className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-sm shadow-sm transition-colors focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring"
                />
              </div>
            </div>
            <p className="text-xs text-muted-foreground">
              Leave both dates blank to export all logs.
            </p>
            <Button variant="outline" onClick={handleExportLogs} data-testid="button-export-logs">
              <Download className="w-4 h-4 mr-2" />
              Download CSV
            </Button>
          </div>
        </div>
        </div>

        {/* ── SharePoint Connection ───────────────────────────────────────── */}
        <div
          className={activeTab === "knowledge-sources" ? "block space-y-6" : "hidden"}
          aria-label="SharePoint connection settings"
          data-testid="settings-panel-sharepoint"
        >
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
                Fetches files from every enabled named knowledge source and loads them into the chatbot.
                Supports Word (.docx), PDF, and plain text files. Each source replaces only its own previously synced documents.
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
              <div className="flex flex-col gap-3 border-t border-[#e2ebef] bg-[#fbfcfd] px-4 py-4 sm:flex-row sm:items-center sm:justify-between sm:px-6">
                <div className="flex items-center gap-2 text-xs text-muted-foreground">
                  <CheckCircle className={`h-3.5 w-3.5 ${activeHasUnsavedChanges ? "text-amber-500" : "text-[#1e9b75]"}`} />
                  {activeHasUnsavedChanges ? "Review changes before leaving this tab" : "Last updated just now"}
                </div>
                <div className="flex items-center gap-2">
                  <Button type="button" variant="outline" onClick={handleDiscardActiveTab} data-testid="button-discard-settings">
                    Discard
                  </Button>
                  <Button type="button" onClick={handleSaveActiveTab} disabled={activeSavePending} className="bg-[#188b6a] text-white hover:bg-[#147457]" data-testid="button-save-settings">
                    <Save className="mr-2 h-4 w-4" />
                    {activeSavePending ? "Saving…" : "Save changes"}
                  </Button>
                </div>
              </div>
            </Tabs>
          </div>
        </main>
      </div>
    </div>
  );
}
