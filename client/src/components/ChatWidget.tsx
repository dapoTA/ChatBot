import { useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { MessageCircle, X, Minus } from "lucide-react";
import { Button } from "@/components/ui/button";
import Chat from "@/pages/Chat";

export function ChatWidget() {
  const [isOpen, setIsOpen] = useState(false);
  const [isMinimized, setIsMinimized] = useState(false);

  return (
    <>
      {/* Floating Button */}
      <motion.button
        initial={{ scale: 0 }}
        animate={{ scale: 1 }}
        onClick={() => {
          setIsOpen(true);
          setIsMinimized(false);
        }}
        className="fixed bottom-6 right-6 z-50 p-4 rounded-full bg-primary text-primary-foreground shadow-lg hover:shadow-xl transition-shadow"
        data-testid="button-open-chat"
        aria-label="Open chat"
      >
        <MessageCircle className="w-6 h-6" />
      </motion.button>

      {/* Chat Dialog */}
      <AnimatePresence>
        {isOpen && (
          <motion.div
            initial={{ opacity: 0, scale: 0.9, y: 100 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.9, y: 100 }}
            className="fixed bottom-6 right-6 z-50 w-full max-w-lg h-[600px] bg-background border border-border rounded-2xl shadow-2xl flex flex-col overflow-hidden"
            data-testid="dialog-chat"
          >
            {/* Header with Controls */}
            <div className="h-16 border-b border-border bg-background flex items-center justify-between px-6 flex-shrink-0">
              <h1 className="text-base font-semibold text-foreground">
                ON-PNT® Assistant
              </h1>
              <div className="flex gap-3">
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => setIsMinimized(true)}
                  data-testid="button-minimize"
                  className="h-9 w-9"
                >
                  <Minus className="w-5 h-5" />
                </Button>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => setIsOpen(false)}
                  data-testid="button-close-chat"
                  className="h-9 w-9"
                >
                  <X className="w-5 h-5" />
                </Button>
              </div>
            </div>

            {/* Chat Content */}
            {!isMinimized && <Chat isWidget />}
          </motion.div>
        )}
      </AnimatePresence>

      {/* Minimized State - Shows Icon with Badge */}
      <AnimatePresence>
        {isOpen && isMinimized && (
          <motion.button
            initial={{ opacity: 0, scale: 0.9 }}
            animate={{ opacity: 1, scale: 1 }}
            exit={{ opacity: 0, scale: 0.9 }}
            onClick={() => setIsMinimized(false)}
            className="fixed bottom-24 right-6 z-50 p-3 rounded-full bg-primary text-primary-foreground shadow-lg hover:shadow-xl transition-shadow"
            data-testid="button-restore-chat"
            aria-label="Restore chat"
          >
            <MessageCircle className="w-5 h-5" />
          </motion.button>
        )}
      </AnimatePresence>
    </>
  );
}
