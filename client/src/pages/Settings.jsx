import { useState, useEffect, useRef } from "react";
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

const appearanceSchema = insertAppSettingsSchema.omit({
  welcomeMessage: true,
  notFoundMessage: true,
  customInstructions: true,
}).extend({
  assistantName: z.string().min(1, "Assistant name is required"),
}).superRefine((data, ctx) => {
  if (data.theme === "custom" && !/^#[0-9a-fA-F]{6}$/.test(data.customThemeColor ?? "")) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      path: ["customThemeColor"],
      message: "Choose a valid 6-digit hexadecimal color for the custom theme.",
    });
  }
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
  sharepointMode: "online",
  description: "",
  welcomeMessage: "",
  notFoundMessage: "",
  instructions: "",
  smeTeam: "",
  contactMethod: "",
  contactDetails: "",
  escalationMessage: "",
  enabled: true,
  isPortalWide: false,
};

function SourceInstructionPreview({ instructions }) {
  const value = instructions ?? "";
  const previewPhrases = extractStyledPhrases(value);
  const previewBodyStyle = extractGlobalResponseStyle(value);
  const hasPreview = previewPhrases.length > 0 || previewBodyStyle !== null;

  return (
    <div className="rounded-lg border border-[#d9e2e8] bg-white px-4 py-3" data-testid="source-formatting-preview">
      <p className="mb-2 text-[10px] font-bold uppercase tracking-[0.13em] text-[#617589]">
        Response formatting preview
      </p>
      {hasPreview ? (
        <div style={previewBodyStyle ?? {}}>
          {previewPhrases.map((phrase, index) => (
            <p
              key={index}
              className="mb-1 text-[12px] leading-5"
              dangerouslySetInnerHTML={{ __html: phrase.html }}
            />
          ))}
          <p className="text-[12px] leading-5">
            Here is the information requested from this knowledge source.
          </p>
        </div>
      ) : (
        <p className="text-[11px] italic leading-5 text-[#718198]">
          Add formatting guidance to this source's instructions to preview it here.
        </p>
      )}
    </div>
  );
}

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

const HEX_COLOR_PATTERN = /^#[0-9a-fA-F]{6}$/;

function isValidHexColor(value) {
  return HEX_COLOR_PATTERN.test(String(value ?? "").trim());
}

function hexToRgb(hex) {
  const value = String(hex).replace("#", "");
  return {
    red: Number.parseInt(value.slice(0, 2), 16),
    green: Number.parseInt(value.slice(2, 4), 16),
    blue: Number.parseInt(value.slice(4, 6), 16),
  };
}

function hexToHsl(hex) {
  const { red, green, blue } = hexToRgb(hex);
  const values = [red, green, blue].map((channel) => channel / 255);
  const max = Math.max(...values);
  const min = Math.min(...values);
  const lightness = (max + min) / 2;

  if (max === min) return `0 0% ${Math.round(lightness * 100)}%`;

  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue;
  if (max === values[0]) hue = (values[1] - values[2]) / delta + (values[1] < values[2] ? 6 : 0);
  else if (max === values[1]) hue = (values[2] - values[0]) / delta + 2;
  else hue = (values[0] - values[1]) / delta + 4;

  return `${Math.round(hue * 60)} ${Math.round(saturation * 100)}% ${Math.round(lightness * 100)}%`;
}

function getContrastColor(hex) {
  const { red, green, blue } = hexToRgb(hex);
  const luminance = [red, green, blue]
    .map((channel) => channel / 255)
    .map((channel) => (channel <= 0.03928 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4))
    .reduce((sum, channel, index) => sum + channel * [0.2126, 0.7152, 0.0722][index], 0);
  const whiteContrast = 1.05 / (luminance + 0.05);
  const blackContrast = (luminance + 0.05) / 0.05;
  return whiteContrast >= blackContrast ? "#ffffff" : "#000000";
}

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
      welcomeMessage: source.welcomeMessage ?? "",
      notFoundMessage: source.notFoundMessage ?? "",
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
        sharepointMode: sourceDraft.sharepointMode,
        description: sourceDraft.description.trim(),
        welcomeMessage: sourceDraft.welcomeMessage.trim(),
        notFoundMessage: sourceDraft.notFoundMessage.trim(),
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
      assistantIcon: "message-circle",
      theme: "teal",
      customThemeColor: null,
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
        assistantIcon: appSettingsData.assistantIcon ?? "message-circle",
        theme: appSettingsData.theme ?? "teal",
        customThemeColor: appSettingsData.customThemeColor ?? null,
        launcherLabel: appSettingsData.launcherLabel ?? "Ask inSite",
        launcherPosition: appSettingsData.launcherPosition ?? "bottom-right",
        launcherStyle: appSettingsData.launcherStyle ?? "bubble",
      });
    }
  }, [appSettingsData]);

  const saveAppearance = useMutation({
    mutationFn: async (data) => {
      const response = await apiRequest("POST", "/api/settings", data);
      return response.json();
    },
    onSuccess: (saved) => {
      appearanceForm.reset({
        assistantName: saved.assistantName,
        assistantIcon: saved.assistantIcon ?? "message-circle",
        theme: saved.theme ?? "teal",
        customThemeColor: saved.customThemeColor ?? null,
        launcherLabel: saved.launcherLabel ?? "Ask inSite",
        launcherPosition: saved.launcherPosition ?? "bottom-right",
        launcherStyle: saved.launcherStyle ?? "bubble",
      });
      queryClient.setQueryData(["/api/settings"], (current) => ({ ...current, ...saved }));
      queryClient.invalidateQueries({ queryKey: ["/api/settings"] });
      toast({ title: "Appearance saved", description: "Widget branding and launcher settings updated." });
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

  const selectedThemeName = appearanceForm.watch("theme");
  const customThemeColor = appearanceForm.watch("customThemeColor") ?? "";
  const selectedThemeBase = THEME_OPTIONS[selectedThemeName] ?? THEME_OPTIONS.teal;
  const selectedThemeColor = selectedThemeName === "custom" && isValidHexColor(customThemeColor)
    ? customThemeColor.trim()
    : selectedThemeBase.primary;
  const selectedTheme = {
    ...selectedThemeBase,
    primary: selectedThemeColor,
    foreground: getContrastColor(selectedThemeColor),
  };
  const selectedIcon = ASSISTANT_ICON_OPTIONS.find(
    (option) => option.value === appearanceForm.watch("assistantIcon"),
  ) ?? ASSISTANT_ICON_OPTIONS[0];
  const PreviewIcon = selectedIcon.Icon;
  const updateCustomThemeColor = (value) => {
    appearanceForm.setValue("customThemeColor", value, { shouldDirty: true, shouldValidate: true });
    if (appearanceForm.getValues("theme") !== "custom") {
      appearanceForm.setValue("theme", "custom", { shouldDirty: true });
    }
  };

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
        assistantIcon: appSettingsData.assistantIcon ?? "message-circle",
        theme: appSettingsData.theme ?? "teal",
        customThemeColor: appSettingsData.customThemeColor ?? null,
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
    <div className="min-h-[100dvh] bg-[#f3f6f8] font-sans text-[#25364c]">
      <header className="flex h-[68px] items-center justify-between border-b border-[#d8e0e7] bg-[#10263f] px-5 text-white sm:px-8">
        <div className="flex items-center gap-3">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#1e765c] shadow-[0_4px_14px_rgba(15,74,56,.3)]">
            <ShieldCheck className="h-5 w-5" />
          </div>
          <div className="border-l border-white/20 pl-3 leading-tight">
            <p className="text-[10px] font-bold uppercase tracking-[0.18em] text-[#a7c8bb]">Technical Assurance</p>
            <p className="mt-0.5 text-[15px] font-semibold tracking-[-0.01em]">inSite Assistant</p>
          </div>
        </div>
        <div className="hidden items-center gap-2 text-[11px] text-[#c6d2df] sm:flex">
          <span className="h-2 w-2 rounded-full bg-[#64c392]" />
          Configuration workspace
        </div>
      </header>

      <div className="mx-auto flex min-h-[calc(100vh-68px)] max-w-[1440px]">
        <aside className="hidden w-[252px] shrink-0 border-r border-[#dbe3e9] bg-[#edf2f5] px-4 py-7 md:block">
          <div className="mb-8 px-3">
            <p className="text-[10px] font-bold uppercase tracking-[0.16em] text-[#8190a2]">Administration</p>
            <p className="mt-2 text-[12px] leading-5 text-[#617288]">
              Keep the assistant aligned with your portal and assurance standards.
            </p>
          </div>
          <nav className="space-y-1" aria-label="Administration">
            <a href="#" className="flex items-center gap-3 rounded-lg px-3 py-2.5 text-[13px] text-[#65768a] hover:bg-white/70">
              <LayoutDashboard className="h-4 w-4" />
              Overview
            </a>
            <a href="#assistant-settings" className="flex items-center gap-3 rounded-lg bg-white px-3 py-2.5 text-[13px] font-semibold text-[#1e765c] shadow-sm">
              <Sparkles className="h-4 w-4" />
              Assistant
            </a>
            <a href="#" className="flex items-center gap-3 rounded-lg px-3 py-2.5 text-[13px] text-[#65768a] hover:bg-white/70">
              <SlidersHorizontal className="h-4 w-4" />
              Experience
            </a>
            <a href="#" className="flex items-center gap-3 rounded-lg px-3 py-2.5 text-[13px] text-[#65768a] hover:bg-white/70">
              <Activity className="h-4 w-4" />
              Usage
            </a>
          </nav>
          <div className="mx-3 mt-10 border-t border-[#d7e0e7] pt-5">
            <p className="text-[10px] font-bold uppercase tracking-[0.16em] text-[#8190a2]">Environment</p>
            <div className="mt-3 space-y-2 text-[11px] text-[#8795a4]">
              <div className="flex items-center gap-2 text-[12px] text-[#51657a]">
                <FolderOpen className="h-4 w-4" />
                Production portal
              </div>
              <div className="flex items-center gap-2">
                <LockKeyhole className="h-[13px] w-[13px]" />
                Admin access only
              </div>
            </div>
          </div>
        </aside>

        <main className="min-w-0 flex-1 px-5 py-7 sm:px-8 lg:px-12">
          <div className="mx-auto max-w-[1120px] space-y-7">
            <div id="assistant-settings" className="flex items-end justify-between gap-4">
              <div>
                <div className="mb-2 flex items-center gap-2 text-[11px] font-medium text-[#7c8b9c]">
                  <span>Assistant</span>
                  <ChevronRight className="h-3 w-3" />
                  <span className="text-[#1e765c]">Settings</span>
                </div>
                <h1 className="text-[27px] font-semibold tracking-[-0.035em] text-[#172d46]" data-testid="text-settings-title">
                  Assistant settings
                </h1>
                <p className="mt-1.5 text-[13px] text-[#718198]">
                  Control how inSite Assistant is presented, grounded, and maintained.
                </p>
              </div>
              <div className={`hidden items-center gap-2 pb-1 text-[11px] sm:flex ${activeHasUnsavedChanges ? "text-[#a16a33]" : "text-[#477b68]"}`}>
                <span className={`h-2 w-2 rounded-full ${activeHasUnsavedChanges ? "bg-[#d18a3d]" : "bg-[#43a274]"}`} />
                {activeHasUnsavedChanges ? "Unsaved changes" : "All changes saved"}
              </div>
            </div>

            <Tabs value={activeTab} onValueChange={setActiveTab} className="overflow-hidden rounded-xl border border-[#d8e1e8] bg-white shadow-[0_8px_30px_rgba(35,57,77,.06)]">
              <TabsList className="grid h-auto w-full grid-cols-2 gap-0 rounded-none border-b border-[#dce4e9] bg-[#fbfcfd] px-2 py-0 sm:grid-cols-4 sm:px-4">
                <TabsTrigger id="settings-tab-general" aria-controls="settings-panel-general" value="general" className="h-16 rounded-none border-b-2 border-transparent bg-transparent px-3 text-left text-[12px] text-[#718198] shadow-none data-[state=active]:border-[#1e765c] data-[state=active]:bg-transparent data-[state=active]:text-[#1b7057] data-[state=active]:shadow-none">
                  <span className="flex items-center gap-2.5">
                    <SlidersHorizontal className="h-4 w-4" />
                    <span><strong className="block font-semibold">General</strong><span className="hidden text-[10px] font-normal opacity-70 md:block">Branding and presentation</span></span>
                  </span>
                </TabsTrigger>
                <TabsTrigger id="settings-tab-knowledge-sources" aria-controls="settings-panel-knowledge-sources settings-panel-sharepoint" value="knowledge-sources" className="h-16 rounded-none border-b-2 border-transparent bg-transparent px-3 text-left text-[12px] text-[#718198] shadow-none data-[state=active]:border-[#1e765c] data-[state=active]:bg-transparent data-[state=active]:text-[#1b7057] data-[state=active]:shadow-none">
                  <span className="flex items-center gap-2.5">
                    <FolderOpen className="h-4 w-4" />
                    <span><strong className="block font-semibold">Knowledge Sources</strong><span className="hidden text-[10px] font-normal opacity-70 md:block">Where answers come from</span></span>
                  </span>
                </TabsTrigger>
                <TabsTrigger id="settings-tab-chat-behavior" aria-controls="settings-panel-chat-behavior" value="chat-behavior" className="h-16 rounded-none border-b-2 border-transparent bg-transparent px-3 text-left text-[12px] text-[#718198] shadow-none data-[state=active]:border-[#1e765c] data-[state=active]:bg-transparent data-[state=active]:text-[#1b7057] data-[state=active]:shadow-none">
                  <span className="flex items-center gap-2.5">
                    <MessageSquare className="h-4 w-4" />
                    <span><strong className="block font-semibold">Chat Behavior</strong><span className="hidden text-[10px] font-normal opacity-70 md:block">Response guardrails</span></span>
                  </span>
                </TabsTrigger>
                <TabsTrigger id="settings-tab-logging-exports" aria-controls="settings-panel-logging-exports" value="logging-exports" className="h-16 rounded-none border-b-2 border-transparent bg-transparent px-3 text-left text-[12px] text-[#718198] shadow-none data-[state=active]:border-[#1e765c] data-[state=active]:bg-transparent data-[state=active]:text-[#1b7057] data-[state=active]:shadow-none">
                  <span className="flex items-center gap-2.5">
                    <BarChart3 className="h-4 w-4" />
                    <span><strong className="block font-semibold">Logging &amp; Exports</strong><span className="hidden text-[10px] font-normal opacity-70 md:block">Records and retention</span></span>
                  </span>
                </TabsTrigger>
              </TabsList>
              <div className="flex flex-col p-6 sm:p-8 lg:p-10">

        {/* ── Widget Appearance ───────────────────────────────────────────── */}
        <div
          id="settings-panel-general"
          className={activeTab === "general" ? "block" : "hidden"}
          role="tabpanel"
          aria-labelledby="settings-tab-general"
          aria-label="General settings"
          data-testid="settings-panel-general"
        >
        <form onSubmit={appearanceForm.handleSubmit((d) => saveAppearance.mutate(d))} className="space-y-10">
          <section aria-labelledby="presentation-heading">
            <div className="mb-6">
              <p className="mb-1 text-[10px] font-bold uppercase tracking-[0.16em] text-[#1e765c]">Presentation</p>
              <h2 id="presentation-heading" className="text-[19px] font-semibold tracking-[-0.02em] text-[#1c3048]">Make the assistant recognisable</h2>
              <p className="mt-1.5 text-[12px] leading-5 text-[#718198]">These values appear in the chat launcher, header, and response footer across the inSite portal.</p>
            </div>
            <div className="space-y-6">
                <div className="grid gap-5 sm:grid-cols-2">
                  <div className="space-y-2">
                    <Label>Organisation label</Label>
                    <div className="flex h-11 items-center rounded-lg border border-[#ced8e2] bg-[#fbfcfd] px-3.5 text-[13px] text-[#596d82]">TECHNICAL ASSURANCE</div>
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="assistantName">Assistant name</Label>
                    <Input id="assistantName" placeholder="ON-PNT® Assistant" data-testid="input-assistant-name" {...appearanceForm.register("assistantName")} className="h-11 bg-[#fbfcfd]" />
                    {appearanceForm.formState.errors.assistantName && <p className="text-xs text-destructive">{appearanceForm.formState.errors.assistantName.message}</p>}
                  </div>
                </div>
                <div className="space-y-2">
                  <Label>Description <span className="float-right font-normal text-[#8090a3]">Shown beneath the assistant name</span></Label>
                  <div className="min-h-[78px] rounded-lg border border-[#ced8e2] bg-[#fbfcfd] px-3.5 py-3 text-[13px] leading-5 text-[#596d82]">Find answers from your trusted portal documents, without leaving inSite.</div>
                </div>
                <div className="grid gap-5 sm:grid-cols-2">
                  <div className="space-y-2">
                    <Label>Footer attribution</Label>
                    <div className="flex h-11 items-center rounded-lg border border-[#ced8e2] bg-[#fbfcfd] px-3.5 text-[13px] text-[#596d82]">Powered by inSite knowledge</div>
                  </div>
                  <div className="space-y-2">
                    <Label>Grounding statement</Label>
                    <div className="flex h-11 items-center rounded-lg border border-[#ced8e2] bg-[#fbfcfd] px-3.5 text-[13px] text-[#596d82]">Answers are Technical Assurance specific</div>
                  </div>
                </div>
            </div>
          </section>

          <section id="theme-launcher" className="border-t border-[#e3e9ed] pt-9" data-testid="theme-launcher-section">
            <div className="mb-6">
              <p className="mb-1 text-[10px] font-bold uppercase tracking-[0.16em] text-[#1e765c]">Theme &amp; launcher</p>
              <h2 className="text-[19px] font-semibold tracking-[-0.02em] text-[#1c3048]">Keep every assistant state on-brand</h2>
              <p className="mt-1.5 text-[12px] leading-5 text-[#718198]">Choose the visual theme and launcher icon once. The selection is used in the expanded header, minimized long bar, and closed floating bubble.</p>
            </div>
            <div className="grid gap-6 lg:grid-cols-[1fr_1.2fr]">
              <div className="rounded-xl border border-[#dbe3e9] p-5">
                <div className="text-[13px] font-semibold text-[#263c54]">Assistant icon</div>
                <p className="mt-1 text-[11px] leading-5 text-[#8290a0]">The selected mark appears in the header, minimized bar, and floating launcher.</p>
                <div className="mt-4 grid grid-cols-3 gap-2">
                  {ASSISTANT_ICON_OPTIONS.map(({ value, label, Icon }) => {
                    const isSelected = appearanceForm.watch("assistantIcon") === value;
                    return (
                      <button key={value} type="button" onClick={() => appearanceForm.setValue("assistantIcon", value, { shouldDirty: true })} className={`flex min-h-[72px] flex-col items-center justify-center gap-1.5 rounded-lg border px-2 py-2 text-[11px] font-medium transition ${isSelected ? "border-[#1e765c] bg-[#edf8f4] text-[#126b52] ring-1 ring-[#1e765c]/20" : "border-[#dbe3e9] bg-[#fbfcfd] text-[#718198] hover:border-[#a9c2b6]"}`} aria-label={`Assistant icon: ${label}`} aria-pressed={isSelected} data-testid={`button-assistant-icon-${value}`}>
                        <Icon className="h-5 w-5" />
                        {label}
                      </button>
                    );
                  })}
                </div>
                <div className="mt-7 border-t border-[#e3e9ed] pt-5">
                  <div className="text-[13px] font-semibold text-[#263c54]">Accent theme</div>
                  <p className="mt-1 text-[11px] leading-5 text-[#8290a0]">Theme colors are applied to the assistant header, actions, focus states, and launcher accents.</p>
                  <div className="mt-4 grid gap-3 sm:grid-cols-2">
                    {Object.entries(THEME_OPTIONS).map(([value, theme]) => {
                      const isSelected = appearanceForm.watch("theme") === value;
                      return (
                        <button key={value} type="button" onClick={() => appearanceForm.setValue("theme", value, { shouldDirty: true })} className={`flex items-center gap-3 rounded-lg border p-3 text-left transition ${isSelected ? "border-[#1e765c] bg-[#f4faf7] ring-2 ring-[#1e765c]/10" : "border-[#dbe3e9] bg-[#fbfcfd] hover:border-[#a9c2b6]"}`} aria-pressed={isSelected} data-testid={`button-theme-${value}`}>
                          <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg" style={{ backgroundColor: value === "teal" ? "#10263f" : theme.primary }}><span className="h-3.5 w-3.5 rounded-md" style={{ backgroundColor: theme.primary }} /></span>
                          <span className="min-w-0 flex-1"><span className="block text-[12px] font-semibold text-[#344a60]">{theme.label}</span><span className="mt-0.5 block text-[10px] text-[#8290a0]">{value === "teal" ? "Current portal direction" : "Assistant accent theme"}</span></span>
                          {isSelected && <span className="text-[#1e765c]">✓</span>}
                        </button>
                      );
                    })}
                    <button type="button" onClick={() => appearanceForm.setValue("theme", "custom", { shouldDirty: true })} className={`flex items-center gap-3 rounded-lg border p-3 text-left transition ${selectedThemeName === "custom" ? "border-[#1e765c] bg-[#f4faf7] ring-2 ring-[#1e765c]/10" : "border-[#dbe3e9] bg-[#fbfcfd] hover:border-[#a9c2b6]"}`} aria-pressed={selectedThemeName === "custom"} data-testid="button-theme-custom">
                      <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg" style={{ background: `linear-gradient(135deg, ${isValidHexColor(customThemeColor) ? customThemeColor : "#188b6a"} 0 50%, #10263f 50% 100%)` }}><span className="h-3.5 w-3.5 rounded-md border border-white/80 bg-white/30" /></span>
                      <span className="min-w-0 flex-1"><span className="block text-[12px] font-semibold text-[#344a60]">Custom color</span><span className="mt-0.5 block text-[10px] text-[#8290a0]">Use your portal accent</span></span>
                      {selectedThemeName === "custom" && <span className="text-[#1e765c]">✓</span>}
                    </button>
                  </div>
                  {selectedThemeName === "custom" && (
                    <div className="mt-4 rounded-lg border border-[#dbe3e9] bg-[#fbfcfd] p-3">
                      <div className="flex items-center gap-3">
                        <input
                          type="color"
                          aria-label="Custom theme color picker"
                          value={isValidHexColor(customThemeColor) ? customThemeColor : "#188b6a"}
                          onChange={(event) => updateCustomThemeColor(event.target.value)}
                          className="h-10 w-12 cursor-pointer rounded-md border border-[#ced8e2] bg-white p-1"
                          data-testid="input-custom-theme-color-picker"
                        />
                        <div className="min-w-0 flex-1 space-y-1">
                          <Label htmlFor="customThemeColor">Hexadecimal color</Label>
                          <Input id="customThemeColor" value={customThemeColor} onChange={(event) => updateCustomThemeColor(event.target.value)} placeholder="#188B6A" maxLength={7} aria-invalid={Boolean(appearanceForm.formState.errors.customThemeColor)} data-testid="input-custom-theme-color" className="h-10 bg-white font-mono text-[12px] uppercase" />
                        </div>
                      </div>
                      {appearanceForm.formState.errors.customThemeColor ? <p className="mt-2 text-[11px] text-destructive">{appearanceForm.formState.errors.customThemeColor.message}</p> : <p className="mt-2 text-[11px] text-[#718198]">Enter a 6-digit hexadecimal value, for example #188B6A.</p>}
                    </div>
                  )}
                </div>
              </div>
              <div className="flex min-h-[338px] flex-col overflow-hidden rounded-xl border border-[#dbe3e9] bg-[#eef5f8]">
                <div className="flex items-start justify-between border-b border-[#dbe3e9] bg-white px-4 py-3">
                  <div><div className="text-[12px] font-semibold text-[#344a60]">Live launcher preview</div><div className="mt-0.5 text-[10px] text-[#8290a0]">Updates before you save</div></div>
                  <span className="pt-0.5 text-[10px] font-semibold uppercase tracking-[0.1em] text-[#8290a0]">{appearanceForm.watch("launcherPosition") === "bottom-left" ? "Bottom left" : "Bottom right"}</span>
                </div>
                <div className={`relative min-h-[290px] flex-1 p-5 ${appearanceForm.watch("launcherPosition") === "bottom-left" ? "text-left" : "text-right"}`}>
                  <div className={`absolute bottom-12 flex max-w-[calc(100%-2.5rem)] items-center gap-2 rounded-xl border bg-white px-3 py-2.5 text-left shadow-[0_8px_20px_rgba(44,67,86,.12)] ${appearanceForm.watch("launcherPosition") === "bottom-left" ? "left-5" : "right-5"}`} style={{ borderColor: `${selectedTheme.primary}44` }}>
                    <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-[9px]" style={{ backgroundColor: selectedTheme.primary, color: selectedTheme.foreground }}><PreviewIcon className="h-[17px] w-[17px]" /></span>
                    <span className="min-w-0"><span className="block truncate text-[12px] font-semibold text-[#25364c]">{appearanceForm.watch("launcherStyle") === "pill" ? appearanceForm.watch("launcherLabel") || "Ask inSite" : appearanceForm.watch("assistantName") || "inSite Assistant"}</span><span className="block text-[10px] text-[#8290a0]">Ready to help!</span></span>
                    <span className="ml-2 text-[12px] text-[#52667b]">↗</span><span className="text-[15px] text-[#263443]">×</span>
                  </div>
                  <div className={`absolute bottom-1 ${appearanceForm.watch("launcherPosition") === "bottom-left" ? "left-5" : "right-5"}`}>
                    <span className="flex h-11 w-11 items-center justify-center rounded-full shadow-[0_5px_14px_rgba(38,74,62,.22)]" style={{ backgroundColor: selectedTheme.primary, color: selectedTheme.foreground }}><PreviewIcon className="h-5 w-5" /></span>
                  </div>
                </div>
              </div>
            </div>
            <div className="mt-5 grid gap-4 md:grid-cols-3">
              <div className="space-y-2"><Label htmlFor="launcherLabel">Launcher label</Label><Input id="launcherLabel" maxLength={40} placeholder="Ask inSite" {...appearanceForm.register("launcherLabel")} data-testid="input-launcher-label" className="h-11 bg-[#fbfcfd]" /><p className="text-[11px] text-[#718198]">Shown when the pill launcher is selected.</p></div>
              <div className="space-y-2"><Label htmlFor="launcherStyle">Launcher style</Label><select id="launcherStyle" {...appearanceForm.register("launcherStyle")} className="flex h-11 w-full rounded-lg border border-[#ced8e2] bg-[#fbfcfd] px-3.5 text-[13px] text-[#27394f] outline-none" data-testid="select-launcher-style"><option value="bubble">Icon bubble</option><option value="pill">Icon with label</option></select></div>
              <div className="space-y-2"><Label htmlFor="launcherPosition">Screen position</Label><select id="launcherPosition" {...appearanceForm.register("launcherPosition")} className="flex h-11 w-full rounded-lg border border-[#ced8e2] bg-[#fbfcfd] px-3.5 text-[13px] text-[#27394f] outline-none" data-testid="select-launcher-position"><option value="bottom-right">Bottom right</option><option value="bottom-left">Bottom left</option></select></div>
            </div>
          </section>

          <section className="border-t border-[#e3e9ed] pt-9" aria-labelledby="widget-appearance-heading">
            <div className="mb-6">
              <p className="mb-1 text-[10px] font-bold uppercase tracking-[0.16em] text-[#1e765c]">Conversation guidance</p>
              <h2 id="widget-appearance-heading" className="text-[19px] font-semibold tracking-[-0.02em] text-[#1c3048]">Managed by knowledge source</h2>
              <p className="mt-1.5 text-[12px] leading-5 text-[#718198]">
                Welcome messages, not-found responses, and assistant instructions now belong to each knowledge source so the selected source controls the conversation.
              </p>
            </div>
            <div className="rounded-xl border border-[#dbe8e1] bg-[#f5faf7] px-5 py-4 text-[12px] leading-5 text-[#557568]">
              <p>
                Open <strong className="font-semibold text-[#315f50]">Knowledge Sources</strong>, expand PTO, HR, Company Policies, or All Portal Sources, and configure that source's messages and instructions.
              </p>
              <Button
                type="button"
                variant="outline"
                className="mt-3 border-[#b9d3c7] bg-white text-[#1e765c] hover:bg-[#edf8f4]"
                onClick={() => setActiveTab("knowledge-sources")}
                data-testid="button-open-source-guidance"
              >
                Configure knowledge sources
              </Button>
            </div>
            <Button type="submit" className="mt-6 bg-[#1e765c] text-[12px] font-semibold text-white hover:bg-[#176049]" data-testid="button-save-appearance"><Save className="mr-2 h-4 w-4" /> Save appearance</Button>
            <div className="mt-5 flex items-start gap-3 rounded-lg border border-[#dbe8e1] bg-[#f5faf7] px-4 py-3.5 text-[11px] leading-5 text-[#557568]"><span className="mt-0.5 text-[#1e765c]">ⓘ</span><span><strong className="font-semibold text-[#315f50]">Preview note:</strong> Changes are applied to the launcher preview after saving. Connection, response, and logging controls remain in their own tabs.</span></div>
          </section>
        </form>
        </div>

        {/* ── Knowledge Sources ───────────────────────────────────────────── */}
        <div
          id="settings-panel-knowledge-sources"
          className={activeTab === "knowledge-sources" ? "contents" : "hidden"}
          role="tabpanel"
          aria-labelledby="settings-tab-knowledge-sources"
          aria-label="Knowledge Sources settings"
          data-testid="settings-panel-knowledge-sources"
        >
        <section id="knowledge-sources" className="order-2 mt-7 space-y-5" data-testid="knowledge-sources-section">
          <div className="flex items-start justify-between gap-4">
            <div>
              <p className="mb-1 text-[10px] font-bold uppercase tracking-[0.16em] text-[#1e765c]">Named sources</p>
              <h2 className="text-[19px] font-semibold tracking-[-0.02em] text-[#1c3048]">
                Expose the libraries employees recognise
              </h2>
              <p className="mt-1.5 max-w-2xl text-[12px] leading-5 text-[#718198]">
                Add as many sources as your portal needs. Each source keeps its own response guidance and escalation route.
              </p>
            </div>
            <Button type="button" variant="outline" onClick={openNewSourceEditor} data-testid="button-add-source">
              <Plus className="w-4 h-4 mr-2" />
              Add Source
            </Button>
          </div>

          <div className="overflow-hidden rounded-xl border border-[#dbe3e9]">
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
                    className={`border-b border-[#edf1f3] transition-colors ${isEditing ? "bg-[#f4faf7]" : "bg-white"}`}
                    data-testid={`knowledge-source-${source.id}`}
                  >
                    <div className="flex items-center gap-3 px-4 py-3.5">
                      <button
                        type="button"
                        className="flex-1 min-w-0 text-left"
                        onClick={() => openSourceEditor(source)}
                        aria-expanded={isEditing}
                        data-testid={`button-expand-source-${source.id}`}
                      >
                        <div className="flex flex-wrap items-center gap-2">
                          <FolderOpen className="h-4 w-4 text-[#1e765c]" />
                          <span className="text-[12px] font-semibold text-[#344a60]">{source.name}</span>
                          {!source.isPortalWide && (
                            <span className="text-[10px] font-medium uppercase tracking-wide px-1.5 py-0.5 rounded bg-slate-100 text-slate-600">
                              {source.sharepointMode === "inherit"
                                ? "Active connection"
                                : source.sharepointMode === "onprem"
                                  ? "SharePoint 2019"
                                  : "SharePoint Online"}
                            </span>
                          )}
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
                        <p className="mt-1 truncate text-[11px] text-[#596d82]">
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
                      <form onSubmit={handleSaveKnowledgeSource} className="space-y-4 border-t border-[#dce7e1] bg-[#f4faf7] p-4" data-testid={`source-editor-${source.id}`}>
                        <div className="flex items-center justify-between gap-3">
                          <div>
                            <h3 className="text-[13px] font-semibold text-[#315f50]">Edit knowledge source</h3>
                            <p className="mt-1 text-[11px] leading-5 text-[#6f897f]">
                              These instructions and escalation details apply only when this source is selected.
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
                           <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
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
                            <div className="space-y-1">
                              <Label htmlFor={`source-mode-${source.id}`}>SharePoint environment</Label>
                              <select
                                id={`source-mode-${source.id}`}
                                value={sourceDraft.sharepointMode}
                                onChange={(event) => updateSourceDraft("sharepointMode", event.target.value)}
                                className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-sm shadow-sm focus:outline-none focus:ring-1 focus:ring-ring"
                                data-testid={`select-source-mode-${source.id}`}
                              >
                                <option value="inherit">Use active connection</option>
                                <option value="online">SharePoint Online</option>
                                <option value="onprem">SharePoint 2019 on-premises</option>
                              </select>
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

                        <div className="grid gap-4 lg:grid-cols-2">
                          <div className="space-y-1">
                            <Label htmlFor={`source-welcome-${source.id}`}>Welcome Message</Label>
                            <Textarea
                              id={`source-welcome-${source.id}`}
                              value={sourceDraft.welcomeMessage}
                              onChange={(event) => updateSourceDraft("welcomeMessage", event.target.value)}
                              placeholder={source.isPortalWide
                                ? "Ask me anything about the full portal."
                                : `Ask me anything ${source.name} related.`}
                              rows={3}
                              className="resize-y"
                              data-testid={`input-source-welcome-${source.id}`}
                            />
                            <p className="text-xs text-muted-foreground">Shown when this source is selected and the conversation is empty.</p>
                          </div>
                          <div className="space-y-1">
                            <Label htmlFor={`source-not-found-${source.id}`}>Not Found Message</Label>
                            <Textarea
                              id={`source-not-found-${source.id}`}
                              value={sourceDraft.notFoundMessage}
                              onChange={(event) => updateSourceDraft("notFoundMessage", event.target.value)}
                              placeholder={`I couldn't find relevant information in ${source.name}.`}
                              rows={3}
                              className="resize-y"
                              data-testid={`input-source-not-found-${source.id}`}
                            />
                            <p className="text-xs text-muted-foreground">Used only when this selected source cannot answer.</p>
                          </div>
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
                            Only this source's guidance is added when it is selected.
                          </p>
                        </div>
                        <SourceInstructionPreview instructions={sourceDraft.instructions} />

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
                <div className="space-y-1">
                  <Label htmlFor="new-source-mode">SharePoint environment</Label>
                  <select
                    id="new-source-mode"
                    value={sourceDraft.sharepointMode}
                    onChange={(event) => updateSourceDraft("sharepointMode", event.target.value)}
                    className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-sm shadow-sm focus:outline-none focus:ring-1 focus:ring-ring"
                    data-testid="select-new-source-mode"
                  >
                    <option value="inherit">Use active connection</option>
                    <option value="online">SharePoint Online</option>
                    <option value="onprem">SharePoint 2019 on-premises</option>
                  </select>
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

              <div className="grid gap-4 lg:grid-cols-2">
                <div className="space-y-1">
                  <Label htmlFor="new-source-welcome">Welcome Message</Label>
                  <Textarea
                    id="new-source-welcome"
                    value={sourceDraft.welcomeMessage}
                    onChange={(event) => updateSourceDraft("welcomeMessage", event.target.value)}
                    placeholder={`Ask me anything ${sourceDraft.name || "about this source"} related.`}
                    rows={3}
                    className="resize-y"
                    data-testid="input-new-source-welcome"
                  />
                  <p className="text-xs text-muted-foreground">Shown when this source is selected and the conversation is empty.</p>
                </div>
                <div className="space-y-1">
                  <Label htmlFor="new-source-not-found">Not Found Message</Label>
                  <Textarea
                    id="new-source-not-found"
                    value={sourceDraft.notFoundMessage}
                    onChange={(event) => updateSourceDraft("notFoundMessage", event.target.value)}
                    placeholder="I couldn't find relevant information in this source."
                    rows={3}
                    className="resize-y"
                    data-testid="input-new-source-not-found"
                  />
                  <p className="text-xs text-muted-foreground">Used only when this selected source cannot answer.</p>
                </div>
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
                <p className="text-xs text-muted-foreground">Only this source's guidance is added when it is selected.</p>
              </div>
              <SourceInstructionPreview instructions={sourceDraft.instructions} />

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
          id="settings-panel-chat-behavior"
          className={activeTab === "chat-behavior" ? "block" : "hidden"}
          role="tabpanel"
          aria-labelledby="settings-tab-chat-behavior"
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
          id="settings-panel-logging-exports"
          className={activeTab === "logging-exports" ? "block" : "hidden"}
          role="tabpanel"
          aria-labelledby="settings-tab-logging-exports"
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
          id="settings-panel-sharepoint"
          className={activeTab === "knowledge-sources" ? "contents" : "hidden"}
          role="region"
          aria-labelledby="settings-tab-knowledge-sources"
          aria-label="SharePoint connection settings"
          data-testid="settings-panel-sharepoint"
        >
        <form onSubmit={sharepointForm.handleSubmit((d) => saveSharepoint.mutate(d))} className="order-1 space-y-4">
          <div className="mb-6">
            <p className="mb-1 text-[10px] font-bold uppercase tracking-[0.16em] text-[#1e765c]">SharePoint connection</p>
            <h2 className="text-[19px] font-semibold tracking-[-0.02em] text-[#1c3048]">Connect the knowledge source</h2>
            <p className="mt-1.5 max-w-2xl text-[12px] leading-5 text-[#718198]">Configure the active SharePoint environment, test access, then choose the named document libraries exposed to employees.</p>
          </div>
          <div className="space-y-5 rounded-xl border border-[#dbe3e9] bg-white p-5">

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
        <div className="order-3 mt-7 rounded-xl border border-[#dbe3e9] bg-[#fbfcfd] p-5">
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
        <div className="order-4 mt-5 rounded-lg border border-[#dce5ea] bg-[#f6f9fb] px-4 py-3 text-[11px] leading-5 text-[#65788c]">
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
