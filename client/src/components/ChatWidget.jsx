import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { motion, AnimatePresence } from "framer-motion";
import { MessageCircle, X, Minus, Maximize2, Sparkles, ShieldCheck } from "lucide-react";
import { apiRequest } from "@/lib/queryClient";
import Chat from "@/pages/Chat";

const DEFAULTS = {
  assistantName: "ON-PNT® Assistant",
  welcomeMessage: "Ask me anything about your SharePoint documents.",
  assistantIcon: "message-circle",
  theme: "teal",
  launcherLabel: "Ask inSite",
  launcherPosition: "bottom-right",
  launcherStyle: "bubble",
};

const THEME_PRESETS = {
  teal: { color: "#188b6a", hsl: "160 71% 32%" },
  ocean: { color: "#2563eb", hsl: "221 83% 53%" },
  slate: { color: "#475569", hsl: "215 25% 37%" },
  amber: { color: "#b45309", hsl: "32 95% 37%" },
};

const HEX_COLOR_PATTERN = /^#[0-9a-fA-F]{6}$/;

function isValidHexColor(value) {
  return HEX_COLOR_PATTERN.test(String(value ?? "").trim());
}

function hexToHsl(hex) {
  const value = String(hex).replace("#", "");
  const channels = [0, 2, 4].map((offset) => Number.parseInt(value.slice(offset, offset + 2), 16) / 255);
  const max = Math.max(...channels);
  const min = Math.min(...channels);
  const lightness = (max + min) / 2;
  if (max === min) return `0 0% ${Math.round(lightness * 100)}%`;

  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue;
  if (max === channels[0]) hue = (channels[1] - channels[2]) / delta + (channels[1] < channels[2] ? 6 : 0);
  else if (max === channels[1]) hue = (channels[2] - channels[0]) / delta + 2;
  else hue = (channels[0] - channels[1]) / delta + 4;
  return `${Math.round(hue * 60)} ${Math.round(saturation * 100)}% ${Math.round(lightness * 100)}%`;
}

function getContrastColor(hex) {
  const value = String(hex).replace("#", "");
  const channels = [0, 2, 4].map((offset) => Number.parseInt(value.slice(offset, offset + 2), 16) / 255);
  const luminance = channels
    .map((channel) => (channel <= 0.03928 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4))
    .reduce((sum, channel, index) => sum + channel * [0.2126, 0.7152, 0.0722][index], 0);
  const whiteContrast = 1.05 / (luminance + 0.05);
  const blackContrast = (luminance + 0.05) / 0.05;
  return whiteContrast >= blackContrast ? "#ffffff" : "#000000";
}

function resolveTheme(themeName, customThemeColor) {
  const preset = THEME_PRESETS[themeName] || THEME_PRESETS.teal;
  const color = themeName === "custom" && isValidHexColor(customThemeColor)
    ? customThemeColor.trim()
    : preset.color;
  return {
    color,
    hsl: themeName === "custom" && isValidHexColor(customThemeColor) ? hexToHsl(color) : preset.hsl,
    foreground: getContrastColor(color),
  };
}

const ICONS = {
  "message-circle": MessageCircle,
  sparkles: Sparkles,
  "shield-check": ShieldCheck,
};

export function ChatWidget() {
  const [isOpen, setIsOpen] = useState(false);
  const [isMinimized, setIsMinimized] = useState(false);
  const [messages, setMessages] = useState([]);
  const [isPending, setIsPending] = useState(false);
  const [isError, setIsError] = useState(false);
  const [selectedSourceId, setSelectedSourceId] = useState("");
  const [chatResetKey, setChatResetKey] = useState(0);

  // Read SharePoint user from URL param injected by the embed script / SPFx customizer.
  // Lazy initializer runs once on mount — no effect needed.
  const [spUser] = useState(() =>
    new URLSearchParams(window.location.search).get("spuser") ?? ""
  );
  // Session ID: generated on first open, cleared on close so a new session starts next time.
  const [sessionId, setSessionId] = useState("");

  const { data: settings } = useQuery({
    queryKey: ["/api/settings/public"],
  });

  const { data: knowledgeSources = [] } = useQuery({
    queryKey: ["/api/knowledge-sources/options"],
  });

  useEffect(() => {
    if (knowledgeSources.length === 0) {
      setSelectedSourceId("");
      return;
    }

    const currentIsAvailable = knowledgeSources.some(
      (source) => String(source.id) === String(selectedSourceId),
    );
    if (!currentIsAvailable) {
      const defaultSource = knowledgeSources.find((source) => source.isPortalWide)
        ?? knowledgeSources[0];
      setSelectedSourceId(defaultSource ? String(defaultSource.id) : "");
    }
  }, [knowledgeSources, selectedSourceId]);

  const assistantName = settings?.assistantName || DEFAULTS.assistantName;
  const selectedSource = knowledgeSources.find(
    (source) => String(source.id) === String(selectedSourceId),
  );
  const welcomeMessage = selectedSource?.welcomeMessage
    || settings?.welcomeMessage
    || DEFAULTS.welcomeMessage;
  const responseStyle = selectedSource?.responseStyle
    ?? (selectedSource?.isPortalWide ? settings?.responseStyle : null)
    ?? null;
  const assistantIcon = settings?.assistantIcon || DEFAULTS.assistantIcon;
  const themeName = settings?.theme || DEFAULTS.theme;
  const customThemeColor = settings?.customThemeColor || "";
  const launcherLabel = settings?.launcherLabel || DEFAULTS.launcherLabel;
  const launcherPosition = settings?.launcherPosition || DEFAULTS.launcherPosition;
  const launcherStyle = settings?.launcherStyle || DEFAULTS.launcherStyle;
  const theme = resolveTheme(themeName, customThemeColor);
  const AssistantIcon = ICONS[assistantIcon] || MessageCircle;
  const launcherSide = launcherPosition === "bottom-left" ? { left: 24 } : { right: 24 };
  const dialogSide = launcherPosition === "bottom-left" ? { left: 20 } : { right: 20 };

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
    // Assign a fresh session UUID the first time the widget opens (or after a close).
    setSessionId((prev) => prev || crypto.randomUUID());
    notifyParent("chatbot:open");
  };

  const handleClose = () => {
    setIsOpen(false);
    setIsMinimized(false);
    setMessages([]);
    setIsPending(false);
    setIsError(false);
    setSessionId("");
    apiRequest("DELETE", "/api/chat").catch(() => {});
    notifyParent("chatbot:close");
  };

  const handleMinimize = () => {
    setIsMinimized(true);
    notifyParent("chatbot:minimize");
  };

  const handleMaximize = () => {
    setIsMinimized(false);
    notifyParent("chatbot:open");
  };

  const handleSourceChange = (nextSourceId) => {
    if (String(nextSourceId) === String(selectedSourceId)) return;
    setSelectedSourceId(nextSourceId);
    setMessages([]);
    setIsError(false);
    setChatResetKey((current) => current + 1);
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
      const res = await apiRequest("POST", "/api/chat", {
        message,
        ...(spUser    ? { username:  spUser    } : {}),
        ...(sessionId ? { sessionId: sessionId } : {}),
        ...(selectedSourceId ? { sourceId: Number(selectedSourceId) } : {}),
      });
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
            style={{
              position: "fixed",
              bottom: 24,
              zIndex: 9999,
              backgroundColor: theme.color,
               color: theme.foreground,
              ...launcherSide,
            }}
            className={`flex items-center justify-center gap-2 shadow-lg transition-shadow hover:shadow-xl ${
              launcherStyle === "pill" ? "min-h-12 rounded-full px-5" : "h-14 w-14 rounded-full"
            }`}
            data-testid="button-open-chat"
            aria-label="Open chat"
          >
            <AssistantIcon className="w-6 h-6" style={{ color: theme.foreground }} />
            {launcherStyle === "pill" && <span className="text-sm font-semibold">{launcherLabel}</span>}
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
              bottom: 20,
              right: 20,
              zIndex: 9999,
              width: 420,
              maxWidth: "calc(100vw - 40px)",
              "--primary": theme.hsl,
               "--primary-foreground": hexToHsl(theme.foreground),
              ...dialogSide,
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
                  boxShadow: "0 2px 8px rgba(0,0,0,0.12)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  padding: "0 16px",
                  height: 56,
                  width: 280,
                }}
              >
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div
                    style={{
                      width: 28,
                      height: 28,
                      borderRadius: "50%",
                      background: theme.color,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                    }}
                  >
                    <AssistantIcon
                       style={{ width: 14, height: 14, color: theme.foreground }}
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
                  boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
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
                        background: theme.color,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                      }}
                    >
                      <AssistantIcon
                       style={{ width: 14, height: 14, color: theme.foreground }}
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
                    knowledgeSources={knowledgeSources}
                    selectedSourceId={selectedSourceId}
                     onSourceChange={handleSourceChange}
                     resetInputKey={chatResetKey}
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
