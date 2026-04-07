import { QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { queryClient } from "./lib/queryClient";
import { Route, Switch, Link, useLocation } from "wouter";
import { ChatWidget } from "@/components/ChatWidget";
import Settings from "@/pages/Settings";
import { Settings as SettingsIcon, MessageCircle } from "lucide-react";
import { cn } from "@/lib/utils";

function NavBar() {
  const [location] = useLocation();

  return (
    <nav className="fixed top-0 left-0 right-0 z-40 h-12 bg-background border-b border-border flex items-center justify-between px-5">
      <div className="flex items-center gap-2">
        <div className="w-5 h-5 rounded-full bg-primary flex items-center justify-center">
          <MessageCircle className="w-3 h-3 text-primary-foreground" />
        </div>
        <span className="text-sm font-semibold text-foreground">ON-PNT® Assistant</span>
      </div>
      <div className="flex items-center gap-1">
        <Link
          href="/"
          className={cn(
            "flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg transition-colors",
            location === "/"
              ? "bg-primary/10 text-primary font-medium"
              : "text-muted-foreground hover:text-foreground hover:bg-muted"
          )}
          data-testid="link-nav-home"
        >
          <MessageCircle className="w-3.5 h-3.5" />
          Chat
        </Link>
        <Link
          href="/settings"
          className={cn(
            "flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg transition-colors",
            location === "/settings"
              ? "bg-primary/10 text-primary font-medium"
              : "text-muted-foreground hover:text-foreground hover:bg-muted"
          )}
          data-testid="link-nav-settings"
        >
          <SettingsIcon className="w-3.5 h-3.5" />
          Settings
        </Link>
      </div>
    </nav>
  );
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <div className="w-screen min-h-screen bg-background">
          <NavBar />
          <div className="pt-12">
            <Switch>
              <Route path="/settings" component={Settings} />
              <Route path="/">
                {/* Main page: full-viewport background with floating chat widget */}
                <div className="w-full h-[calc(100vh-48px)] flex items-center justify-center">
                  <div className="text-center space-y-3 select-none">
                    <div className="w-14 h-14 rounded-full bg-primary/10 flex items-center justify-center mx-auto">
                      <MessageCircle className="w-7 h-7 text-primary" />
                    </div>
                    <h2 className="text-lg font-semibold text-foreground">ON-PNT® Assistant</h2>
                    <p className="text-sm text-muted-foreground max-w-xs">
                      Click the chat button in the bottom-right corner to ask questions about your SharePoint documents.
                    </p>
                  </div>
                </div>
                <ChatWidget />
              </Route>
            </Switch>
          </div>
        </div>
        <Toaster />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
