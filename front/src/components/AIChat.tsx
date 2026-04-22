import * as React from "react";
import {
  Textarea,
  Button,
  Spinner,
  Avatar,
} from "@fluentui/react-components";
import {
  Send24Filled,
  Person24Regular,
  Bot24Regular,
  ErrorCircle24Regular,
} from "@fluentui/react-icons";
import type { ChatMessage } from "../types";
import QuickActions from "./QuickActions";

interface AIChatProps {
  documentName: string;
  sessionId: string;
  messages: ChatMessage[];
  onSend: (content: string) => void;
}

const AIChat: React.FC<AIChatProps> = ({
  messages,
  onSend,
}) => {
  const [input, setInput] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const bottomRef = React.useRef<HTMLDivElement>(null);
  const textareaRef = React.useRef<HTMLTextAreaElement>(null);

  // 有消息在 pending 时显示 loading
  React.useEffect(() => {
    const last = messages[messages.length - 1];
    setLoading(last?.role === "user");
  }, [messages]);

  React.useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const handleSend = React.useCallback(() => {
    const trimmed = input.trim();
    if (!trimmed || loading) return;
    setInput("");
    onSend(trimmed);
    textareaRef.current?.focus();
  }, [input, loading, onSend]);

  const handleKeyDown = React.useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        handleSend();
      }
    },
    [handleSend]
  );

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        minHeight: 0,
      }}
    >
      {/* 消息列表 */}
      <div
        style={{
          flex: 1,
          overflowY: "auto",
          overflowX: "hidden",
          padding: "12px 16px",
          display: "flex",
          flexDirection: "column",
          gap: "12px",
          minHeight: 0,
        }}
      >
        {messages.length === 0 && !loading && (
          <div
            style={{
              flex: 1,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              gap: "8px",
              color: "#999",
            }}
          >
            <Bot24Regular style={{ fontSize: "32px" }} />
            <span style={{ fontSize: "14px", textAlign: "center" }}>
              在下方输入问题或指令，AI 会结合当前文档上下文回复
            </span>
          </div>
        )}

        {messages.map((msg) => (
          <div key={msg.id}>
            <div
              style={{
                display: "flex",
                gap: "10px",
                alignItems: "flex-start",
                flexDirection: msg.role === "user" ? "row-reverse" : "row",
              }}
            >
              <Avatar
                size={24}
                icon={
                  msg.role === "user" ? (
                    <Person24Regular />
                  ) : msg.error ? (
                    <ErrorCircle24Regular />
                  ) : (
                    <Bot24Regular />
                  )
                }
                color={msg.role === "user" ? "brand" : "neutral"}
              />
              <div
                style={{
                  maxWidth: "80%",
                  padding: "8px 12px",
                  borderRadius: "12px",
                  fontSize: "14px",
                  lineHeight: "1.5",
                  whiteSpace: "pre-wrap",
                  wordBreak: "break-word",
                  backgroundColor:
                    msg.role === "user"
                      ? "#0078d4"
                      : msg.error
                      ? "#fde7e9"
                      : "#f3f2f1",
                  color:
                    msg.role === "user"
                      ? "#fff"
                      : msg.error
                      ? "#a80000"
                      : "#323130",
                  borderBottomRightRadius: msg.role === "user" ? "4px" : "12px",
                  borderBottomLeftRadius: msg.role === "user" ? "12px" : "4px",
                }}
              >
                {msg.content}
              </div>
            </div>
            <div
              style={{
                fontSize: "11px",
                color: "#999",
                marginTop: "2px",
                paddingLeft: msg.role === "user" ? 0 : "34px",
                textAlign: msg.role === "user" ? "right" : "left",
                paddingRight: msg.role === "user" ? "34px" : 0,
              }}
            >
              {formatTime(msg.timestamp)}
            </div>
          </div>
        ))}

        {loading && (
          <div style={{ display: "flex", gap: "10px", alignItems: "flex-start" }}>
            <Avatar size={24} icon={<Bot24Regular />} color="neutral" />
            <div
              style={{
                padding: "8px 12px",
                borderRadius: "12px",
                fontSize: "14px",
                backgroundColor: "#f3f2f1",
                borderBottomLeftRadius: "4px",
              }}
            >
              <Spinner size="tiny" label="AI 思考中..." labelPosition="after" />
            </div>
          </div>
        )}

        <div ref={bottomRef} />
      </div>

      <QuickActions onSend={onSend} disabled={loading} />

      {/* 输入区 */}
      <div
        style={{
          flexShrink: 0,
          padding: "12px 16px",
          borderTop: "1px solid #edebe9",
          display: "flex",
          flexDirection: "column",
          gap: "8px",
        }}
      >
        <div style={{ display: "flex", gap: "8px", alignItems: "flex-end" }}>
          <Textarea
            ref={textareaRef as React.RefObject<HTMLTextAreaElement>}
            style={{ flex: 1, resize: "none" }}
            value={input}
            onChange={(_, d) => setInput(d.value)}
            onKeyDown={handleKeyDown}
            placeholder="输入您的指令，例如：将第一段设为黑体三号字"
            rows={2}
            disabled={loading}
          />
          <Button
            icon={<Send24Filled />}
            appearance="primary"
            onClick={handleSend}
            disabled={loading || !input.trim()}
            aria-label="发送"
          />
        </div>
        <span style={{ fontSize: "12px", color: "#999" }}>
          按 Enter 发送，Shift+Enter 换行
        </span>
      </div>
    </div>
  );
};

function formatTime(date: Date): string {
  return date.toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" });
}

export default AIChat;
