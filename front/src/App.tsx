import * as React from "react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import Header from "./components/Header";
import AIChat from "./components/AIChat";
import StatusBar from "./components/StatusBar";
import {
  getSessions,
  getWordStatus,
  getWordDocuments,
  getChatHistory,
  clearChat,
  sendChat as apiSendChat,
} from "./api";
import type { SessionInfo, ChatMessage } from "./types";

function uuid(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function convertHistory(turns: Array<{ 用户需求?: string; 回答?: string; 内容?: string; 轮次: number; 角色?: string }>): ChatMessage[] {
  const messages: ChatMessage[] = [];
  for (const turn of turns) {
    if (turn["用户需求"]) {
      messages.push({
        id: uuid(), role: "user", content: turn["用户需求"], timestamp: new Date(),
      });
    }
    if (turn["回答"] || turn["内容"]) {
      messages.push({
        id: uuid(), role: "assistant",
        content: turn["回答"] || turn["内容"] || "",
        timestamp: new Date(),
      });
    }
  }
  return messages;
}

const App: React.FC = () => {
  // ── 会话状态 ────────────────────────────────────────────────────────────
  const [sessions, setSessions] = React.useState<SessionInfo[]>([]);
  const [activeSessionId, setActiveSessionId] = React.useState<string>(() => uuid());
  const [activeDocumentName, setActiveDocumentName] = React.useState<string>("");

  // ── 消息历史（提升到 App 层，切换会话时重新加载）────────────────────────
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);

  // ── Word 连接状态 ─────────────────────────────────────────────────────
  const [wordConnected, setWordConnected] = React.useState(false);
  const [wordDocumentName, setWordDocumentName] = React.useState<string>("");
  const [wordDocuments, setWordDocuments] = React.useState<{ name: string; path: string }[]>([]);
  const [wordLoading] = React.useState(false);
  const [wordError, setWordError] = React.useState<string>("");

  // ── 加载指定会话的历史 ────────────────────────────────────────────────
  const loadSessionHistory = React.useCallback(async (sessionId: string) => {
    try {
      const data = await getChatHistory(sessionId);
      const converted = convertHistory(data.turns);
      setMessages(converted);
    } catch {
      setMessages([]);
    }
  }, []);

  // ── 启动时加载 + Word 状态轮询 ───────────────────────────────────────
  React.useEffect(() => {
    loadSessions();
    checkWordStatus();
    const interval = setInterval(checkWordStatus, 5000);
    return () => clearInterval(interval);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Word 连接后自动同步文档名到当前会话 ────────────────────────────────
  React.useEffect(() => {
    if (wordConnected && wordDocumentName && !activeDocumentName) {
      setActiveDocumentName(wordDocumentName);
    }
  }, [wordConnected, wordDocumentName]);

  const loadSessions = React.useCallback(async () => {
    try {
      const data = await getSessions();
      setSessions(data.sessions);
    } catch {
      // 静默失败
    }
  }, []);

  const checkWordStatus = React.useCallback(async () => {
    try {
      const status = await getWordStatus();
      setWordConnected(status.connected);
      setWordDocumentName(status.document_name || "");
      setWordError("");

      // 同时获取所有打开的文档列表
      if (status.connected) {
        try {
          const docs = await getWordDocuments();
          setWordDocuments(docs.documents || []);
        } catch {
          setWordDocuments([]);
        }
      } else {
        setWordDocuments([]);
      }
    } catch {
      setWordConnected(false);
      setWordDocumentName("");
      setWordError("无法连接后端服务");
      setWordDocuments([]);
    }
  }, []);

  // ── 发送消息（由 AIChat 回调上来，走后端 API）──────────────────────────
  const handleSendMessage = React.useCallback(
    async (content: string) => {
      const userMsg: ChatMessage = {
        id: uuid(), role: "user", content, timestamp: new Date(),
      };
      setMessages((prev) => [...prev, userMsg]);

      try {
        const result = await apiSendChat(
          content,
          activeDocumentName,
          activeSessionId,
          ""
        );
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
      }
    },
    [activeDocumentName, activeSessionId]
  );

  // ── 会话操作 ───────────────────────────────────────────────────────────
  const handleNewSession = React.useCallback(() => {
    const newId = uuid();
    // 优先使用 Word 活动文档名，如果没有则用文档列表第一个，均无则提示"未打开文档"
    const docName = wordDocumentName || (wordDocuments.length > 0 ? wordDocuments[0].name : "未打开文档");
    setActiveSessionId(newId);
    setActiveDocumentName(docName);
    setMessages([]);
    setSessions((prev) => [
      {
        session_id: newId,
        document_name: docName,
        turn_count: 0,
        last_message: "",
        last_updated: new Date().toISOString(),
      },
      ...prev,
    ]);
  }, [wordDocumentName, wordDocuments]);

  const handleSessionChange = React.useCallback(
    async (sessionId: string, documentName: string) => {
      setActiveSessionId(sessionId);
      setActiveDocumentName(documentName);
      await loadSessionHistory(sessionId);
    },
    [loadSessionHistory]
  );

  const handleDeleteSession = React.useCallback(
    async (sessionId: string) => {
      try {
        await clearChat(sessionId);
        setSessions((prev) => prev.filter((s) => s.session_id !== sessionId));
        if (activeSessionId === sessionId) {
          handleNewSession();
        }
      } catch {
        // 静默失败
      }
    },
    [activeSessionId, handleNewSession]
  );

  return (
    <FluentProvider theme={webLightTheme}>
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          height: "100vh",
          minHeight: 0,
          overflow: "hidden",
          backgroundColor: "#fff",
        }}
      >
        <Header
          documentName={activeDocumentName}
          sessions={sessions}
          activeSessionId={activeSessionId}
          wordDocuments={wordDocuments}
          onSessionChange={handleSessionChange}
          onDeleteSession={handleDeleteSession}
          onDocumentSelect={(docName) => {
            setActiveDocumentName(docName);
          }}
        />

        <div style={{ flex: 1, minHeight: 0, overflow: "hidden" }}>
          <AIChat
            documentName={activeDocumentName}
            sessionId={activeSessionId}
            messages={messages}
            onSend={handleSendMessage}
          />
        </div>

        <StatusBar
          connected={wordConnected}
          documentName={wordDocumentName}
          loading={wordLoading}
          error={wordError}
        />
      </div>
    </FluentProvider>
  );
};

export default App;
