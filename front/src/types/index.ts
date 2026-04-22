// ── API 请求/响应类型 ─────────────────────────────────────────────────────

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
  turn: number;
  stage: string;
  skill_selected?: string;
  executed: ExecutedAction[];
}

export interface ExecutedAction {
  step: number;
  skill: string;
  action: string;
  params: Record<string, unknown>;
  success: boolean;
  error?: string;
  description?: string;
  result?: unknown;
}

export interface WordStatus {
  connected: boolean;
  document_name?: string;
  has_selection: boolean;
  selection_text: string;
}

export interface SessionInfo {
  session_id: string;
  document_name: string;
  turn_count: number;
  last_message: string;
  last_updated: string;
}

// ── 聊天消息类型 ────────────────────────────────────────────────────────────

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  timestamp: Date;
  error?: boolean;
}

// ── 主题类型 ────────────────────────────────────────────────────────────────

export type ThemeMode = "light" | "dark";
