import * as React from "react";
import { Spinner } from "@fluentui/react-components";
import { ErrorCircle24Regular } from "@fluentui/react-icons";

interface StatusBarProps {
  connected: boolean;
  documentName?: string;
  loading?: boolean;
  error?: string;
}

const StatusBar: React.FC<StatusBarProps> = ({
  connected,
  documentName,
  loading,
  error,
}) => {
  if (loading) {
    return (
      <div
        style={{
          flexShrink: 0,
          padding: "6px 16px",
          borderTop: "1px solid #edebe9",
          display: "flex",
          alignItems: "center",
          gap: "6px",
          fontSize: "12px",
          color: "#666",
          backgroundColor: "#faf9f8",
        }}
      >
        <Spinner size="tiny" />
        <span>检查 Word 连接...</span>
      </div>
    );
  }

  if (error) {
    return (
      <div
        style={{
          flexShrink: 0,
          padding: "6px 16px",
          borderTop: "1px solid #fde7e9",
          display: "flex",
          alignItems: "center",
          gap: "6px",
          fontSize: "12px",
          color: "#a80000",
          backgroundColor: "#fde7e9",
        }}
      >
        <ErrorCircle24Regular style={{ color: "#a80000" }} />
        <span>{error}</span>
      </div>
    );
  }

  return (
    <div
      style={{
        flexShrink: 0,
        padding: "6px 16px",
        borderTop: "1px solid #edebe9",
        display: "flex",
        alignItems: "center",
        gap: "6px",
        fontSize: "12px",
        color: "#666",
        backgroundColor: "#faf9f8",
      }}
    >
      <span
        style={{
          width: "8px",
          height: "8px",
          borderRadius: "50%",
          backgroundColor: connected ? "#107c10" : "#d83b01",
          flexShrink: 0,
        }}
      />
      <span style={{ color: connected ? "#107c10" : "#d83b01" }}>
        {connected ? "已连接" : "未连接"}
      </span>
      {connected && documentName && (
        <span style={{ color: "#999", marginLeft: "4px" }}>
          · {documentName}
        </span>
      )}
      {connected && !documentName && (
        <span style={{ color: "#999", marginLeft: "4px" }}>
          · 未打开文档
        </span>
      )}
    </div>
  );
};

export default StatusBar;
