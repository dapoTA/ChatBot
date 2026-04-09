import { QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { queryClient } from "./lib/queryClient";
import { Route, Switch } from "wouter";
import { ChatWidget } from "@/components/ChatWidget";
import Settings from "@/pages/Settings";

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <Switch>
          {/* Admin-only settings page — navigate to /settings directly */}
          <Route path="/settings" component={Settings} />

          {/* Embeddable widget page — just the floating chat button, no chrome */}
          <Route>
            <div className="w-screen h-screen bg-transparent">
              <ChatWidget />
            </div>
          </Route>
        </Switch>
        <Toaster />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
