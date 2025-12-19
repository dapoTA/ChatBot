import { useState, useRef, useEffect } from "react";
import { useChatHistory, useSendMessage } from "@/hooks/use-chat";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Send, Bot, User, AlertCircle } from "lucide-react";
import { cn } from "@/lib/utils";
import { motion, AnimatePresence } from "framer-motion";
import ReactMarkdown from "react-markdown";

interface ChatProps {
  isWidget?: boolean;
}

export default function Chat({ isWidget = false }: ChatProps) {
  const { data: messages } = useChatHistory();
  const sendMessage = useSendMessage();
  
  const [input, setInput] = useState("");
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollIntoView({ behavior: "smooth" });
    }
  }, [messages, sendMessage.isPending]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!input.trim() || sendMessage.isPending) return;
    sendMessage.mutate({ message: input });
    setInput("");
  };

  return (
    <div className={cn(
      "flex flex-col bg-background border-border",
      isWidget 
        ? "h-full flex-1" 
        : "h-screen w-full max-w-2xl mx-auto border-l border-r"
    )}>
      {/* Header - only show when not a widget */}
      {!isWidget && (
        <header className="h-14 border-b border-border bg-background flex items-center px-6 flex-shrink-0">
          <h1 className="text-base font-semibold text-foreground">ON-PNT® Assistant</h1>
        </header>
      )}

      {/* Chat Area */}
      <ScrollArea className="flex-1 overflow-hidden">
        <div className="p-6 space-y-4 flex flex-col">
          <AnimatePresence initial={false}>
            {messages?.map((msg) => (
              <motion.div
                key={msg.id}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className={cn(
                  "flex gap-3 max-w-full",
                  msg.role === "user" ? "justify-end" : "justify-start"
                )}
              >
                {msg.role === "assistant" && (
                  <div className="w-6 h-6 rounded-full bg-primary flex items-center justify-center flex-shrink-0 mt-0.5">
                    <Bot className="w-3 h-3 text-primary-foreground" />
                  </div>
                )}
                <div className={cn(
                  "max-w-sm p-3 rounded-lg text-sm leading-relaxed",
                  msg.role === "user" 
                    ? "bg-primary text-primary-foreground rounded-br-none" 
                    : "bg-white dark:bg-secondary text-foreground border border-border rounded-bl-none"
                )}>
                  {msg.role === "assistant" ? (
                    <div className="prose prose-sm dark:prose-invert max-w-none">
                      <ReactMarkdown>{msg.content}</ReactMarkdown>
                    </div>
                  ) : (
                    <p>{msg.content}</p>
                  )}
                </div>
                {msg.role === "user" && (
                  <div className="w-6 h-6 rounded-full bg-foreground/20 flex items-center justify-center flex-shrink-0 mt-0.5">
                    <User className="w-3 h-3 text-foreground" />
                  </div>
                )}
              </motion.div>
            ))}
          </AnimatePresence>

          {sendMessage.isPending && (
            <motion.div
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="flex gap-3 justify-start"
            >
              <div className="w-6 h-6 rounded-full bg-primary flex items-center justify-center flex-shrink-0 mt-0.5">
                <Bot className="w-3 h-3 text-primary-foreground" />
              </div>
              <div className="p-3 bg-white dark:bg-secondary border border-border text-foreground rounded-lg rounded-bl-none flex items-center gap-2">
                <span className="w-2 h-2 bg-foreground/40 rounded-full animate-bounce"></span>
              </div>
            </motion.div>
          )}

          {sendMessage.isError && (
            <div className="flex gap-3 justify-start">
              <div className="w-6 h-6 rounded-full bg-destructive flex items-center justify-center flex-shrink-0 mt-0.5">
                <AlertCircle className="w-3 h-3 text-white" />
              </div>
              <div className="p-3 bg-destructive/10 text-destructive border border-destructive/20 rounded-lg rounded-bl-none text-sm">
                Failed to send. Try again.
              </div>
            </div>
          )}
          
          <div ref={scrollRef} />
        </div>
      </ScrollArea>

      {/* Input Area */}
      <div className="border-t border-border bg-background p-4 flex-shrink-0">
        <form onSubmit={handleSubmit} className="relative flex items-center gap-2">
          <Input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="Type a message..."
            className="flex-1 py-2 text-sm rounded-full border-border focus-visible:ring-primary/20 bg-card"
            disabled={sendMessage.isPending}
            data-testid="input-message"
          />
          <Button 
            type="submit" 
            size="icon" 
            disabled={!input.trim() || sendMessage.isPending}
            className="rounded-full h-8 w-8 flex-shrink-0"
            data-testid="button-send"
          >
            <Send className="w-4 h-4" />
          </Button>
        </form>
      </div>
    </div>
  );
}
