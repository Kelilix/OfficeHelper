/**
 * OfficeHelper AI Agent API Client
 * 与 pservice FastAPI 后端通信
 */

const _origin = process.env.ADDIN_API_ORIGIN ?? "";
const API_BASE =
  _origin.length > 0 ? `${_origin.replace(/\/$/, "")}/api` : "/api";

// ── 类型定义 ─────────────────────────────────────────────────────

export interface ChatRequest {
  message: string;
  selection_text: string;
  document_name: string;
  session_id: string;
}

export interface ChatResponse {
  response: string;
  success: boolean;
  error?: string;
  session_id: string;
}

export interface WordStatus {
  connected: boolean;
  document_name?: string;
  has_selection: boolean;
  selection_text: string;
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  timestamp: Date;
  error?: boolean;
}

// ── 工具函数 ─────────────────────────────────────────────────────

function uuid(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

// ── API 调用 ─────────────────────────────────────────────────────

/**
 * 发送聊天消息到 AI Agent
 * @param sessionId  前端生成的 UUID，同一窗口内所有消息共享同一个 sessionId
 */
export async function sendChat(
  message: string,
  selectionText: string,
  documentName: string,
  sessionId: string
): Promise<ChatResponse> {
  const body: ChatRequest = {
    message,
    selection_text: selectionText,
    document_name: documentName,
    session_id: sessionId,
  };

  const resp = await fetch(`${API_BASE}/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const text = await resp.text().catch(() => "网络请求失败");
    throw new Error(`HTTP ${resp.status}: ${text}`);
  }

  return resp.json() as Promise<ChatResponse>;
}

/**
 * 获取 Word 连接状态
 */
export async function getWordStatus(): Promise<WordStatus> {
  const resp = await fetch(`${API_BASE}/word/status`);
  if (!resp.ok) {
    throw new Error(`HTTP ${resp.status}`);
  }
  return resp.json() as Promise<WordStatus>;
}

/**
 * 主动连接 Word
 */
export async function connectWord(): Promise<{ success: boolean; error?: string }> {
  const resp = await fetch(`${API_BASE}/word/connect`, { method: "POST" });
  return resp.json() as Promise<{ success: boolean; error?: string }>;
}

/**
 * 断开 Word 连接
 */
export async function disconnectWord(
  save?: boolean
): Promise<{ success: boolean; error?: string }> {
  const resp = await fetch(`${API_BASE}/word/disconnect?save=${save ?? false}`, {
    method: "POST",
  });
  return resp.json() as Promise<{ success: boolean; error?: string }>;
}

// ── Office.js 工具 ────────────────────────────────────────────────

/**
 * 获取当前 Word 选中的文本
 */
export async function getWordSelection(): Promise<string> {
  return new Promise((resolve) => {
    try {
      Word.run(async (context) => {
        const sel = context.document.getSelection();
        context.load(sel, "text");
        await context.sync();
        resolve((sel as any).text?.trim() ?? "");
      }).catch(() => resolve(""));
    } catch {
      resolve("");
    }
  });
}

/**
 * 获取当前文档名称
 */
export async function getDocumentName(): Promise<string> {
  return new Promise((resolve) => {
    try {
      Word.run(async (context) => {
        const props = context.document.properties;
        context.load(props, "title");
        await context.sync();
        resolve((props as any).title || document.title || "未命名文档");
      }).catch(() => resolve("未命名文档"));
    } catch {
      resolve("未命名文档");
    }
  });
}

/**
 * 在文档末尾插入文本（复用原有的 taskpane.ts 方法）
 */
export async function insertTextToDoc(text: string): Promise<void> {
  const { insertText } = await import("./taskpane");
  return insertText(text);
}
