/**
 * OfficeHelper API Client
 * 与 pservice FastAPI 后端通信（独立于 Word Add-in，无 Office.js 依赖）
 */

import type {
  ChatRequest,
  ChatResponse,
  WordStatus,
  SessionInfo,
} from "./types";

const API_BASE = "http://127.0.0.1:8765/api";

// ── 工具函数 ─────────────────────────────────────────────────────────────

function uuid(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

async function fetchJson<T>(url: string, init?: RequestInit): Promise<T> {
  const resp = await fetch(url, {
    ...init,
    headers: {
      "Content-Type": "application/json",
      ...(init?.headers ?? {}),
    },
  });
  if (!resp.ok) {
    const text = await resp.text().catch(() => "网络请求失败");
    throw new Error(`HTTP ${resp.status}: ${text}`);
  }
  return resp.json() as Promise<T>;
}

// ── API 调用 ─────────────────────────────────────────────────────────────

/**
 * 发送聊天消息到 AI Agent
 */
export async function sendChat(
  message: string,
  documentName: string,
  sessionId: string,
  selectionText: string = ""
): Promise<ChatResponse> {
  const body: ChatRequest = {
    message,
    selection_text: selectionText,
    document_name: documentName,
    session_id: sessionId,
  };

  return fetchJson<ChatResponse>(`${API_BASE}/chat`, {
    method: "POST",
    body: JSON.stringify(body),
  });
}

/**
 * 获取 Word 连接状态
 */
export async function getWordStatus(): Promise<WordStatus> {
  return fetchJson<WordStatus>(`${API_BASE}/word/status`);
}

/**
 * 主动连接 Word
 */
export async function connectWord(): Promise<{ success: boolean; error?: string }> {
  return fetchJson(`${API_BASE}/word/connect`, { method: "POST" });
}

/**
 * 断开 Word 连接
 */
export async function disconnectWord(
  save?: boolean
): Promise<{ success: boolean; error?: string }> {
  return fetchJson(`${API_BASE}/word/disconnect?save=${save ?? false}`, {
    method: "POST",
  });
}

/**
 * 获取所有会话列表
 */
export async function getSessions(): Promise<{ sessions: SessionInfo[] }> {
  return fetchJson(`${API_BASE}/sessions`);
}

/**
 * 获取指定会话的对话历史
 * 后端返回 to_user_json() 格式（中文 key）
 */
export async function getChatHistory(
  sessionId: string
): Promise<{ session_id: string; turns: TurnJson[] }> {
  return fetchJson(`${API_BASE}/chat/history?session_id=${encodeURIComponent(sessionId)}`);
}

interface TurnJson {
  "轮次": number;
  "用户需求"?: string;
  "回答"?: string;
  "内容"?: string;
  "描述"?: string;
  "执行结果"?: unknown;
}

/**
 * 清除指定会话
 */
export async function clearChat(sessionId: string): Promise<{ success: boolean }> {
  return fetchJson(`${API_BASE}/chat/clear?session_id=${encodeURIComponent(sessionId)}`, {
    method: "DELETE",
  });
}

/**
 * 获取当前所有打开的 Word 文档
 */
export async function getWordDocuments(): Promise<
  { documents: { name: string; path: string }[] }
> {
  return fetchJson(`${API_BASE}/word/documents`);
}

export { uuid };
