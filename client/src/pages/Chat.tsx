import { useState, useRef, useEffect } from "react";
import { useChatHistory, useSendMessage, useClearChat } from "@/hooks/use-chat";
import { Sidebar } from "@/components/Sidebar";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { 
  Send, 
  Bot, 
  Trash2, 
  Sparkles, 
  User, 
  AlertCircle 
} from "lucide-react";
import { cn } from "@/lib/utils";
import { motion, AnimatePresence } from "framer-motion";
import { format } from "date-fns";
import ReactMarkdown from "react-markdown";

export default function Chat() {
  const { data: messages, isLoading: isHistoryLoading } = useChatHistory();
  const sendMessage = useSendMessage();
  const clearChat = useClearChat();
  
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
    <div className="min-h-screen bg-background flex overflow-hidden">
      <Sidebar />
      <main className="flex-1 pl-64 flex flex-col h-screen relative">
        {/* Header */}
        <header className="h-16 border-b border-border bg-background/80 backdrop-blur-md flex items-center justify-between px-6 sticky top-0 z-10">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-primary/10 rounded-lg">
              <Sparkles className="w-5 h-5 text-primary" />
            </div>
            <div>
              <h2 className="font-semibold">SharePoint Assistant</h2>
              <p className="text-xs text-muted-foreground">Connected to On-Premises Index</p>
            </div>
          </div>
          <Button 
            variant="ghost" 
            size="sm"
            onClick={() => clearChat.mutate()}
            disabled={!messages?.length || clearChat.isPending}
            className="text-muted-foreground hover:text-destructive"
          >
            <Trash2 className="w-4 h-4 mr-2" />
            Clear Chat
          </Button>
        </header>

        {/* Chat Area */}
        <ScrollArea className="flex-1 p-6">
          <div className="max-w-3xl mx-auto space-y-6 pb-6">
            {!messages?.length && !isHistoryLoading && (
              <motion.div 
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="text-center py-20"
              >
                <div className="w-20 h-20 bg-primary/5 rounded-3xl flex items-center justify-center mx-auto mb-6">
                  <Bot className="w-10 h-10 text-primary" />
                </div>
                <h3 className="text-2xl font-bold text-foreground mb-3">How can I help you?</h3>
                <p className="text-muted-foreground max-w-md mx-auto">
                  Ask me about documents, list items, or general information stored in your SharePoint collection.
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-8 max-w-lg mx-auto">
                  {[
                    "Summarize project alpha specs",
                    "Find documents about HR policies",
                    "List latest marketing assets",
                    "What is the status of ticket #42?"
                  ].map((suggestion) => (
                    <button
                      key={suggestion}
                      onClick={() => setInput(suggestion)}
                      className="text-sm p-3 rounded-xl border border-border bg-card hover:bg-secondary/50 hover:border-primary/30 transition-all text-left"
                    >
                      "{suggestion}"
                    </button>
                  ))}
                </div>
              </motion.div>
            )}

            <AnimatePresence initial={false}>
              {messages?.map((msg) => (
                <motion.div
                  key={msg.id}
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className={cn(
                    "flex gap-4 max-w-2xl",
                    msg.role === "user" ? "ml-auto flex-row-reverse" : "mr-auto"
                  )}
                >
                  <div className={cn(
                    "w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 mt-1",
                    msg.role === "user" ? "bg-primary text-primary-foreground" : "bg-green-600 text-white"
                  )}>
                    {msg.role === "user" ? <User className="w-4 h-4" /> : <Bot className="w-4 h-4" />}
                  </div>
                  <div className={cn(
                    "p-4 text-sm leading-relaxed",
                    msg.role === "user" 
                      ? "chat-bubble-user" 
                      : "chat-bubble-ai prose prose-sm dark:prose-invert max-w-none"
                  )}>
                    {msg.role === "assistant" ? (
                      <ReactMarkdown>{msg.content}</ReactMarkdown>
                    ) : (
                      <p>{msg.content}</p>
                    )}
                    <span className={cn(
                      "text-[10px] block mt-2 opacity-70",
                      msg.role === "user" ? "text-primary-foreground/80" : "text-muted-foreground"
                    )}>
                      {msg.createdAt && format(new Date(msg.createdAt), "h:mm a")}
                    </span>
                  </div>
                </motion.div>
              ))}
            </AnimatePresence>

            {sendMessage.isPending && (
              <motion.div
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="flex gap-4 max-w-2xl mr-auto"
              >
                <div className="w-8 h-8 rounded-full bg-green-600 flex items-center justify-center flex-shrink-0 mt-1 text-white">
                  <Bot className="w-4 h-4" />
                </div>
                <div className="p-4 bg-secondary rounded-2xl rounded-bl-none flex items-center gap-2">
                  <span className="w-2 h-2 bg-foreground/40 rounded-full animate-bounce [animation-delay:-0.3s]"></span>
                  <span className="w-2 h-2 bg-foreground/40 rounded-full animate-bounce [animation-delay:-0.15s]"></span>
                  <span className="w-2 h-2 bg-foreground/40 rounded-full animate-bounce"></span>
                </div>
              </motion.div>
            )}

            {sendMessage.isError && (
              <div className="flex gap-4 max-w-2xl mr-auto">
                 <div className="w-8 h-8 rounded-full bg-destructive flex items-center justify-center flex-shrink-0 mt-1 text-white">
                  <AlertCircle className="w-4 h-4" />
                </div>
                <div className="p-4 bg-destructive/10 text-destructive border border-destructive/20 rounded-2xl rounded-bl-none">
                  Failed to get response. Please try again.
                </div>
              </div>
            )}
            
            <div ref={scrollRef} />
          </div>
        </ScrollArea>

        {/* Input Area */}
        <div className="p-6 bg-background/80 backdrop-blur-md border-t border-border">
          <form onSubmit={handleSubmit} className="max-w-3xl mx-auto relative flex items-center gap-3">
            <Input
              value={input}
              onChange={(e) => setInput(e.target.value)}
              placeholder="Ask a question about your documents..."
              className="pr-12 py-6 text-base shadow-sm rounded-xl border-border/60 focus-visible:ring-primary/20 bg-card"
              disabled={sendMessage.isPending}
            />
            <Button 
              type="submit" 
              size="icon" 
              disabled={!input.trim() || sendMessage.isPending}
              className={cn(
                "absolute right-2 rounded-lg transition-all duration-200",
                input.trim() ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground"
              )}
            >
              <Send className="w-4 h-4" />
            </Button>
          </form>
          <div className="text-center mt-3">
            <p className="text-[10px] text-muted-foreground">
              AI can make mistakes. Please verify important information.
            </p>
          </div>
        </div>
      </main>
    </div>
  );
}
