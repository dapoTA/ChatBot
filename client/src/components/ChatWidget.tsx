import { useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { MessageCircle, X, Minus, Maximize2 } from "lucide-react";
import Chat from "@/pages/Chat";

export function ChatWidget() {
  const [isOpen, setIsOpen] = useState(false);
  const [isMinimized, setIsMinimized] = useState(false);

  const handleOpen = () => {
    setIsOpen(true);
    setIsMinimized(false);
  };

  const handleClose = () => {
    setIsOpen(false);
    setIsMinimized(false);
  };

  const handleMinimize = () => {
    setIsMinimized(true);
  };

  const handleMaximize = () => {
    setIsMinimized(false);
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
            {/* Minimized pill */}
            {isMinimized ? (
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
                    style={{
                      fontSize: 14,
                      fontWeight: 600,
                      color: "hsl(var(--foreground))",
                    }}
                  >
                    ON-PNT® Assistant
                  </span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                  <button
                    onClick={handleMaximize}
                    data-testid="button-maximize"
                    aria-label="Maximize chat"
                    style={{
                      width: 32,
                      height: 32,
                      borderRadius: 8,
                      border: "none",
                      background: "transparent",
                      cursor: "pointer",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      color: "hsl(var(--foreground))",
                    }}
                    onMouseOver={(e) =>
                      (e.currentTarget.style.background = "hsl(var(--muted))")
                    }
                    onMouseOut={(e) =>
                      (e.currentTarget.style.background = "transparent")
                    }
                  >
                    <Maximize2 style={{ width: 14, height: 14 }} />
                  </button>
                  <button
                    onClick={handleClose}
                    data-testid="button-close-minimized"
                    aria-label="Close chat"
                    style={{
                      width: 32,
                      height: 32,
                      borderRadius: 8,
                      border: "none",
                      background: "transparent",
                      cursor: "pointer",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      color: "hsl(var(--foreground))",
                    }}
                    onMouseOver={(e) =>
                      (e.currentTarget.style.background = "hsl(var(--muted))")
                    }
                    onMouseOut={(e) =>
                      (e.currentTarget.style.background = "transparent")
                    }
                  >
                    <X style={{ width: 14, height: 14 }} />
                  </button>
                </div>
              </div>
            ) : (
              /* Full expanded dialog */
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
                        style={{
                          width: 14,
                          height: 14,
                          color: "hsl(var(--primary-foreground))",
                        }}
                      />
                    </div>
                    <span
                      style={{
                        fontSize: 14,
                        fontWeight: 600,
                        color: "hsl(var(--foreground))",
                      }}
                    >
                      ON-PNT® Assistant
                    </span>
                  </div>

                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <button
                      onClick={handleMinimize}
                      data-testid="button-minimize"
                      aria-label="Minimize chat"
                      style={{
                        width: 32,
                        height: 32,
                        borderRadius: 8,
                        border: "none",
                        background: "transparent",
                        cursor: "pointer",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        color: "hsl(var(--foreground))",
                      }}
                      onMouseOver={(e) =>
                        (e.currentTarget.style.background = "hsl(var(--muted))")
                      }
                      onMouseOut={(e) =>
                        (e.currentTarget.style.background = "transparent")
                      }
                    >
                      <Minus style={{ width: 14, height: 14 }} />
                    </button>
                    <button
                      onClick={handleClose}
                      data-testid="button-close-chat"
                      aria-label="Close chat"
                      style={{
                        width: 32,
                        height: 32,
                        borderRadius: 8,
                        border: "none",
                        background: "transparent",
                        cursor: "pointer",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        color: "hsl(var(--foreground))",
                      }}
                      onMouseOver={(e) =>
                        (e.currentTarget.style.background = "hsl(var(--muted))")
                      }
                      onMouseOut={(e) =>
                        (e.currentTarget.style.background = "transparent")
                      }
                    >
                      <X style={{ width: 14, height: 14 }} />
                    </button>
                  </div>
                </div>

                {/* Chat Body */}
                <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
                  <Chat isWidget />
                </div>
              </div>
            )}
          </motion.div>
        )}
      </AnimatePresence>
    </>
  );
}
