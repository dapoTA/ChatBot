import { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { motion, AnimatePresence } from "framer-motion";
import { MessageCircle, X, Minus, Maximize2 } from "lucide-react";
import { apiRequest } from "@/lib/queryClient";
import Chat from "@/pages/Chat";

const DEFAULTS = {
  assistantName: "ON-PNT® Assistant",
  welcomeMessage: "Ask me anything about your SharePoint documents.",
};

export function ChatWidget() {
  const [isOpen, setIsOpen] = useState(false);
  const [isMinimized, setIsMinimized] = useState(false);
  const [messages, setMessages] = useState([]);
  const [isPending, setIsPending] = useState(false);
  const [isError, setIsError] = useState(false);

  const { data: settings } = useQuery({
    queryKey: ["/api/settings"],
  });

  const assistantName = settings?.assistantName || DEFAULTS.assistantName;
  const welcomeMessage = settings?.welcomeMessage || DEFAULTS.welcomeMessage;
  const responseStyle = settings?.responseStyle || null;

  const notifyParent = (event) => {
    try {
      if (window.parent && window.parent !== window) {
        window.parent.postMessage(event, "*");
      }
    } catch {
      // Silently ignore cross-origin errors in non-iframe contexts
    }
  };

  const handleOpen = () => {
    setIsOpen(true);
    setIsMinimized(false);
    notifyParent("chatbot:open");
  };

  const handleClose = () => {
    setIsOpen(false);
    setIsMinimized(false);
    setMessages([]);
    setIsPending(false);
    setIsError(false);
    apiRequest("DELETE", "/api/chat").catch(() => {});
    notifyParent("chatbot:close");
  };

  const handleMinimize = () => {
    setIsMinimized(true);
    notifyParent("chatbot:close");
  };

  const handleMaximize = () => {
    setIsMinimized(false);
    notifyParent("chatbot:open");
  };

  const handleSend = async (message) => {
    const userMsg = {
      id: Date.now(),
      role: "user",
      content: message,
      createdAt: new Date(),
    };

    setMessages((prev) => [...prev, userMsg]);
    setIsPending(true);
    setIsError(false);

    try {
      const res = await apiRequest("POST", "/api/chat", { message });
      if (!res.ok) throw new Error("Failed");
      const assistantMsg = await res.json();
      setMessages((prev) => [...prev, assistantMsg]);
    } catch {
      setIsError(true);
    } finally {
      setIsPending(false);
    }
  };

  return (
    <>
      {/* Floating Trigger Button — hidden when dialog is open */}
      <AnimatePresence>
        {!isOpen && (
          <motion.button
            key="trigger"
            initial={{ scale: 0, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0, opacity: 0 }}
            onClick={handleOpen}
            style={{ position: "fixed", bottom: 24, right: 24, zIndex: 9999 }}
            className="p-4 rounded-full bg-primary text-primary-foreground shadow-lg hover:shadow-xl transition-shadow"
            data-testid="button-open-chat"
            aria-label="Open chat"
          >
            <MessageCircle className="w-6 h-6" />
          </motion.button>
        )}
      </AnimatePresence>

      {/* Chat Dialog */}
      <AnimatePresence>
        {isOpen && (
          <motion.div
            key="dialog"
            initial={{ opacity: 0, y: 60, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 60, scale: 0.95 }}
            transition={{ type: "spring", stiffness: 300, damping: 30 }}
            style={{
              position: "fixed",
              bottom: 24,
              right: 24,
              zIndex: 9999,
              width: 420,
              maxWidth: "calc(100vw - 48px)",
            }}
            data-testid="dialog-chat"
          >
            {isMinimized ? (
              /* Minimized pill */
              <div
                style={{
                  background: "hsl(var(--background))",
                  border: "1px solid hsl(var(--border))",
                  borderRadius: 16,
                  boxShadow: "0 8px 32px rgba(0,0,0,0.18)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  padding: "0 16px",
                  height: 56,
                }}
              >
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div
                    style={{
                      width: 28,
                      height: 28,
                      borderRadius: "50%",
                      background: "hsl(var(--primary))",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                    }}
                  >
                    <MessageCircle
                      style={{ width: 14, height: 14, color: "hsl(var(--primary-foreground))" }}
                    />
                  </div>
                  <span
                    style={{ fontSize: 14, fontWeight: 600, color: "hsl(var(--foreground))" }}
                  >
                    {assistantName}
                  </span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                  <IconButton onClick={handleMaximize} label="Maximize" testId="button-maximize">
                    <Maximize2 style={{ width: 14, height: 14 }} />
                  </IconButton>
                  <IconButton onClick={handleClose} label="Close" testId="button-close-minimized">
                    <X style={{ width: 14, height: 14 }} />
                  </IconButton>
                </div>
              </div>
            ) : (
              /* Expanded dialog */
              <div
                style={{
                  background: "hsl(var(--background))",
                  border: "1px solid hsl(var(--border))",
                  borderRadius: 16,
                  boxShadow: "0 8px 32px rgba(0,0,0,0.18)",
                  height: 580,
                  display: "flex",
                  flexDirection: "column",
                  overflow: "hidden",
                }}
              >
                {/* Header */}
                <div
                  style={{
                    height: 56,
                    borderBottom: "1px solid hsl(var(--border))",
                    background: "hsl(var(--background))",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    padding: "0 16px",
                    flexShrink: 0,
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <div
                      style={{
                        width: 28,
                        height: 28,
                        borderRadius: "50%",
                        background: "hsl(var(--primary))",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                      }}
                    >
                      <MessageCircle
                        style={{ width: 14, height: 14, color: "hsl(var(--primary-foreground))" }}
                      />
                    </div>
                    <span style={{ fontSize: 14, fontWeight: 600, color: "hsl(var(--foreground))" }}>
                      {assistantName}
                    </span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <IconButton onClick={handleMinimize} label="Minimize" testId="button-minimize">
                      <Minus style={{ width: 14, height: 14 }} />
                    </IconButton>
                    <IconButton onClick={handleClose} label="Close and clear chat" testId="button-close-chat">
                      <X style={{ width: 14, height: 14 }} />
                    </IconButton>
                  </div>
                </div>

                {/* Chat Body */}
                <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
                  <Chat
                    isWidget
                    messages={messages}
                    isPending={isPending}
                    isError={isError}
                    onSend={handleSend}
                    welcomeMessage={welcomeMessage}
                    responseStyle={responseStyle}
                  />
                </div>
              </div>
            )}
          </motion.div>
        )}
      </AnimatePresence>
    </>
  );
}

function IconButton({ onClick, label, testId, children }) {
  const [hovered, setHovered] = useState(false);

  return (
    <button
      onClick={onClick}
      data-testid={testId}
      aria-label={label}
      style={{
        width: 32,
        height: 32,
        borderRadius: 8,
        border: "none",
        background: hovered ? "hsl(var(--muted))" : "transparent",
        cursor: "pointer",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        color: "hsl(var(--foreground))",
        transition: "background 0.15s",
      }}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
    >
      {children}
    </button>
  );
}
