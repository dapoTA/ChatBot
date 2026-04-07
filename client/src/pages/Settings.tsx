import { useState } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import { queryClient } from "@/lib/queryClient";
import { apiRequest } from "@/lib/queryClient";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { useToast } from "@/hooks/use-toast";
import {
  CheckCircle,
  XCircle,
  RefreshCw,
  Wifi,
  Save,
  FileText,
  Clock,
  AlertTriangle,
} from "lucide-react";
import type { SharepointConfig } from "@shared/schema";
import { insertSharepointConfigSchema } from "@shared/schema";

const formSchema = insertSharepointConfigSchema.extend({
  siteUrl: z.string().url("Must be a valid URL (e.g. http://sharepoint.company.com/sites/mysite)"),
  domain: z.string().min(1, "Domain is required"),
  username: z.string().min(1, "Username is required"),
  password: z.string().min(1, "Password is required"),
  libraryName: z.string().min(1, "Library name is required"),
});

type FormValues = z.infer<typeof formSchema>;

export default function Settings() {
  const { toast } = useToast();
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);

  const { data: config, isLoading } = useQuery<SharepointConfig | null>({
    queryKey: ["/api/sharepoint/config"],
  });

  const form = useForm<FormValues>({
    resolver: zodResolver(formSchema),
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

  const saveConfig = useMutation({
    mutationFn: (data: FormValues) =>
      apiRequest("POST", "/api/sharepoint/config", data),
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
    onError: async (err: any) => {
      toast({
        title: "Sync failed",
        description: "Could not sync documents. Check your SharePoint configuration.",
        variant: "destructive",
      });
    },
  });

  const onSubmit = (data: FormValues) => {
    saveConfig.mutate(data);
  };

  return (
    <div className="min-h-screen bg-background">
      <div className="max-w-2xl mx-auto px-6 py-10">

        {/* Page Header */}
        <div className="mb-8">
          <h1 className="text-2xl font-semibold text-foreground" data-testid="text-settings-title">
            SharePoint Settings
          </h1>
          <p className="text-sm text-muted-foreground mt-1">
            Connect to your on-premises SharePoint 2019 document library.
            Documents are synced into the chatbot's knowledge base.
          </p>
        </div>

        {/* Config Form */}
        <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-6">

          <div className="rounded-xl border border-border bg-card p-6 space-y-5">
            <h2 className="text-sm font-semibold text-foreground uppercase tracking-wide">
              Connection Details
            </h2>

            <div className="space-y-1">
              <Label htmlFor="siteUrl">SharePoint Site URL</Label>
              <Input
                id="siteUrl"
                placeholder="http://sharepoint.company.com/sites/mysite"
                data-testid="input-site-url"
                {...form.register("siteUrl")}
              />
              {form.formState.errors.siteUrl && (
                <p className="text-xs text-destructive">{form.formState.errors.siteUrl.message}</p>
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
                  {...form.register("domain")}
                />
                {form.formState.errors.domain && (
                  <p className="text-xs text-destructive">{form.formState.errors.domain.message}</p>
                )}
              </div>
              <div className="space-y-1">
                <Label htmlFor="libraryName">Document Library Name</Label>
                <Input
                  id="libraryName"
                  placeholder="Documents"
                  data-testid="input-library-name"
                  {...form.register("libraryName")}
                />
                {form.formState.errors.libraryName && (
                  <p className="text-xs text-destructive">{form.formState.errors.libraryName.message}</p>
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
                  {...form.register("username")}
                />
                {form.formState.errors.username && (
                  <p className="text-xs text-destructive">{form.formState.errors.username.message}</p>
                )}
              </div>
              <div className="space-y-1">
                <Label htmlFor="password">Password</Label>
                <Input
                  id="password"
                  type="password"
                  autoComplete="current-password"
                  data-testid="input-password"
                  {...form.register("password")}
                />
                {form.formState.errors.password && (
                  <p className="text-xs text-destructive">{form.formState.errors.password.message}</p>
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
                checked={form.watch("allowSelfSigned")}
                onCheckedChange={(v) => form.setValue("allowSelfSigned", v)}
                data-testid="switch-self-signed"
              />
            </div>
          </div>

          {/* Actions */}
          <div className="flex items-center gap-3">
            <Button
              type="submit"
              disabled={saveConfig.isPending}
              data-testid="button-save-config"
            >
              <Save className="w-4 h-4 mr-2" />
              {saveConfig.isPending ? "Saving…" : "Save Settings"}
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

          {/* Test result */}
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

        {/* Sync Section */}
        <div className="mt-8 rounded-xl border border-border bg-card p-6">
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
        <div className="mt-6 rounded-lg bg-muted/40 border border-border px-4 py-3 text-xs text-muted-foreground">
          <strong className="text-foreground">On-premises deployment:</strong> This application must be hosted on a server
          that has direct network access to your SharePoint 2019 environment. Authentication uses NTLM with
          your Active Directory credentials.
        </div>
      </div>
    </div>
  );
}
