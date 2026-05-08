import { useState, useRef, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Send, Bot, User, AlertCircle } from "lucide-react";
import { cn } from "@/lib/utils";
import { motion, AnimatePresence } from "framer-motion";
import ReactMarkdown from "react-markdown";
import rehypeRaw from "rehype-raw";
// rehypeRaw allows the AI to emit styled HTML (e.g. <span style="...">).
// Content comes from the OpenAI API, not directly from end-users.

export default function Chat({ isWidget = false, messages, isPending, isError, onSend, welcomeMessage = "Ask me anything about your SharePoint documents.", responseStyle = null }) {

  // Convert responseStyle object → a CSS style object applied to assistant message containers.
  // CSS inheritance means child elements with their own inline styles (e.g. the red bold prefix)
  // will naturally override the inherited global colour/weight.
  const globalMsgStyle = responseStyle ? {
    ...(responseStyle.color      && { color: responseStyle.color }),
    ...(responseStyle.bold       && { fontWeight: 'bold' }),
    ...(responseStyle.italic     && { fontStyle: 'italic' }),
    ...(responseStyle.underline  && { textDecoration: 'underline' }),
    ...(responseStyle.fontSize   && { fontSize: responseStyle.fontSize.replace('font-size:', '').replace(';', '') }),
  } : {};
  const [input, setInput] = useState("");
  const scrollRef = useRef(null);

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollIntoView({ behavior: "smooth" });
    }
  }, [messages, isPending]);

  const handleSubmit = (e) => {
    e.preventDefault();
    const trimmed = input.trim();
    if (!trimmed || isPending) return;
    setInput("");
    onSend(trimmed);
  };

  return (
    <div
      className={cn(
        "flex flex-col bg-background border-border",
        isWidget ? "h-full flex-1" : "h-screen w-full max-w-2xl mx-auto border-l border-r"
      )}
    >
      {!isWidget && (
        <header className="h-14 border-b border-border bg-background flex items-center px-6 flex-shrink-0">
          <h1 className="text-base font-semibold text-foreground">ON-PNT® Assistant</h1>
        </header>
      )}

      <ScrollArea className="flex-1 overflow-hidden">
        <div className="p-6 space-y-4 flex flex-col">
          {/* Empty state */}
          {messages.length === 0 && !isPending && (
            <div className="flex flex-col items-center justify-center h-40 text-center gap-2">
              <div className="w-10 h-10 rounded-full bg-primary/10 flex items-center justify-center">
                <Bot className="w-5 h-5 text-primary" />
              </div>
              <p className="text-sm text-muted-foreground">
                {welcomeMessage}
              </p>
            </div>
          )}

          <AnimatePresence initial={false}>
            {messages.map((msg) => (
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
                <div
                  className={cn(
                    "max-w-sm p-3 rounded-lg text-sm leading-relaxed",
                    msg.role === "user"
                      ? "bg-primary text-primary-foreground rounded-br-none"
                      : "bg-white dark:bg-secondary text-foreground border border-border rounded-bl-none"
                  )}
                >
                  {msg.role === "assistant" ? (
                    <div className="prose prose-sm dark:prose-invert max-w-none" style={globalMsgStyle}>
                  <ReactMarkdown
                    rehypePlugins={[rehypeRaw]}
                    components={{
                      a: ({ href, children }) => (
                        <a href={href} target={href?.startsWith("mailto:") ? "_self" : "_blank"} rel="noreferrer">{children}</a>
                      )
                    }}
                  >
                    {msg.content.replace(/\b([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})\b/g, '[$1](mailto:$1)')}
                  </ReactMarkdown>
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

          {/* Typing indicator */}
          {isPending && (
            <motion.div
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="flex gap-3 justify-start"
            >
              <div className="w-6 h-6 rounded-full bg-primary flex items-center justify-center flex-shrink-0 mt-0.5">
                <Bot className="w-3 h-3 text-primary-foreground" />
              </div>
              <div className="p-3 bg-white dark:bg-secondary border border-border text-foreground rounded-lg rounded-bl-none flex items-center gap-1.5">
                <span className="w-1.5 h-1.5 bg-foreground/40 rounded-full animate-bounce [animation-delay:0ms]" />
                <span className="w-1.5 h-1.5 bg-foreground/40 rounded-full animate-bounce [animation-delay:150ms]" />
                <span className="w-1.5 h-1.5 bg-foreground/40 rounded-full animate-bounce [animation-delay:300ms]" />
              </div>
            </motion.div>
          )}

          {/* Error state */}
          {isError && (
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

      {/* Input */}
      <div className="border-t border-border bg-background p-4 flex-shrink-0">
        <form onSubmit={handleSubmit} className="relative flex items-center gap-2">
          <Input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="Type a message..."
            className="flex-1 py-2 text-sm rounded-full border-border focus-visible:ring-primary/20 bg-card"
            disabled={isPending}
            data-testid="input-message"
          />
          <Button
            type="submit"
            size="icon"
            disabled={!input.trim() || isPending}
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
