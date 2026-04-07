import * as React from "react";
import {
  Textarea,
  Button,
  Spinner,
  Text,
  makeStyles,
  tokens,
  Avatar,
} from "@fluentui/react-components";
import {
  Send24Filled,
  Person24Regular,
  Bot24Regular,
  ErrorCircle24Regular,
} from "@fluentui/react-icons";
import { sendChat, getWordSelection, getDocumentName, ChatMessage } from "../api";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    minHeight: 0,
  },
  messages: {
    flex: 1,
    overflowY: "auto",
    overflowX: "hidden",
    padding: "12px 16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    minHeight: 0,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  messageRow: {
    display: "flex",
    gap: "10px",
    alignItems: "flex-start",
  },
  messageRowUser: {
    flexDirection: "row-reverse",
  },
  bubble: {
    maxWidth: "80%",
    padding: "8px 12px",
    borderRadius: "12px",
    fontSize: "14px",
    lineHeight: "1.5",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  bubbleUser: {
    backgroundColor: tokens.colorBrandBackground,
    // 品牌背景上须用 OnBrand 前景；BrandForeground1 是「品牌色字在浅底上」，易与背景撞色看不见
    color: tokens.colorNeutralForegroundOnBrand,
    borderBottomRightRadius: "4px",
  },
  bubbleAssistant: {
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1,
    borderBottomLeftRadius: "4px",
  },
  bubbleError: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
    borderBottomLeftRadius: "4px",
  },
  meta: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    marginTop: "2px",
    paddingLeft: "34px",
  },
  inputArea: {
    flexShrink: 0,
    padding: "12px 16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  inputRow: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-end",
  },
  textarea: {
    flex: 1,
    resize: "none",
  },
  emptyState: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "8px",
    color: tokens.colorNeutralForeground3,
  },
});

function formatTime(date: Date): string {
  return date.toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" });
}

function uuid(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

const AIChat: React.FC = () => {
  const styles = useStyles();
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);
  const [input, setInput] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const bottomRef = React.useRef<HTMLDivElement>(null);
  const textareaRef = React.useRef<HTMLTextAreaElement>(null);
  // sessionId 在组件首次挂载时生成，同一窗口内所有消息共享该 ID
  const [sessionId] = React.useState(() => uuid());

  React.useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const handleSend = React.useCallback(async () => {
    const trimmed = input.trim();
    if (!trimmed || loading) return;

    const userMsg: ChatMessage = {
      id: uuid(),
      role: "user",
      content: trimmed,
      timestamp: new Date(),
    };
    setMessages((prev) => [...prev, userMsg]);
    setInput("");
    setLoading(true);

    try {
      const [selectionText, docName] = await Promise.all([
        getWordSelection(),
        getDocumentName(),
      ]);

      const result = await sendChat(trimmed, selectionText, docName, sessionId);

      const assistantMsg: ChatMessage = {
        id: uuid(),
        role: "assistant",
        content: result.response,
        timestamp: new Date(),
        error: !result.success,
      };
      setMessages((prev) => [...prev, assistantMsg]);
    } catch (err) {
      const assistantMsg: ChatMessage = {
        id: uuid(),
        role: "assistant",
        content: `请求失败：${err instanceof Error ? err.message : String(err)}`,
        timestamp: new Date(),
        error: true,
      };
      setMessages((prev) => [...prev, assistantMsg]);
    } finally {
      setLoading(false);
      textareaRef.current?.focus();
    }
  }, [input, loading]);

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
    <div className={styles.container}>
      <div className={styles.messages}>
        {messages.length === 0 && !loading && (
          <div className={styles.emptyState}>
            <Bot24Regular style={{ fontSize: "32px" }} />
            <Text size={300} align="center">
              在下方输入问题或指令，AI 会结合当前选区与文档上下文回复；对话记录会保留在本面板中。
            </Text>
          </div>
        )}

        {messages.map((msg) => (
          <div key={msg.id}>
            <div
              className={`${styles.messageRow} ${msg.role === "user" ? styles.messageRowUser : ""}`}
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
                className={`${styles.bubble} ${msg.role === "user" ? styles.bubbleUser : msg.error ? styles.bubbleError : styles.bubbleAssistant}`}
              >
                {msg.content}
              </div>
            </div>
            <div className={styles.meta}>{formatTime(msg.timestamp)}</div>
          </div>
        ))}

        {loading && (
          <div className={styles.messageRow}>
            <Avatar size={24} icon={<Bot24Regular />} color="neutral" />
            <div className={`${styles.bubble} ${styles.bubbleAssistant}`}>
              <Spinner size="tiny" label="AI 思考中..." labelPosition="after" />
            </div>
          </div>
        )}

        <div ref={bottomRef} />
      </div>

      <div className={styles.inputArea}>
        <div className={styles.inputRow}>
          <Textarea
            ref={textareaRef as any}
            className={styles.textarea}
            value={input}
            onChange={(_, d) => setInput(d.value)}
            onKeyDown={handleKeyDown}
            placeholder="输入您的指令，例如：把选中文字设为黑体三号字"
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
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          按 Enter 发送，Shift+Enter 换行
        </Text>
      </div>
    </div>
  );
};

export default AIChat;
