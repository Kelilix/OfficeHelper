import * as React from "react";
import {
  Button,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  MenuDivider,
  Text,
} from "@fluentui/react-components";
import {
  DocumentRegular,
  CaretDown24Regular,
  Chat20Regular,
  Delete20Regular,
} from "@fluentui/react-icons";

interface HeaderProps {
  documentName: string;
  sessions: { session_id: string; document_name: string }[];
  activeSessionId: string;
  wordDocuments: { name: string; path: string }[];
  onSessionChange: (sessionId: string, documentName: string) => void;
  onDeleteSession: (sessionId: string) => void;
  onDocumentSelect?: (docName: string) => void;
}

const Header: React.FC<HeaderProps> = ({
  documentName,
  sessions,
  activeSessionId,
  wordDocuments,
  onSessionChange,
  onDeleteSession,
  onDocumentSelect,
}) => {
  const activeSession = sessions.find((s) => s.session_id === activeSessionId);
  const displayName = activeSession?.document_name || documentName || "未打开文档";

  return (
    <div
      style={{
        flexShrink: 0,
        padding: "10px 16px",
        borderBottom: "1px solid #edebe9",
        display: "flex",
        alignItems: "center",
        gap: "8px",
        backgroundColor: "#faf9f8",
      }}
    >
      {/* Logo */}
      <img
        src="/icon.svg"
        alt="Logo"
        style={{ width: "28px", height: "28px", borderRadius: "6px" }}
        onError={(e) => {
          (e.target as HTMLImageElement).style.display = "none";
        }}
      />

      {/* 会话选择下拉 */}
      <div style={{ flex: 1, minWidth: 0 }}>
        <Menu>
          <MenuTrigger>
            <Button
              appearance="subtle"
              icon={<DocumentRegular />}
              iconPosition="before"
              style={{
                maxWidth: "220px",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
              }}
            >
              <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
                {displayName}
              </span>
              <CaretDown24Regular />
            </Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              {wordDocuments.length > 0 && (
                <>
                  <MenuDivider />
                  <Text
                    size={100}
                    style={{
                      display: "block",
                      padding: "4px 12px 2px",
                      color: "#666",
                      fontSize: "11px",
                    }}
                  >
                    已打开的文档
                  </Text>
                  {wordDocuments.map((doc) => (
                    <MenuItem
                      key={doc.path || doc.name}
                      icon={<DocumentRegular />}
                      onClick={() => onDocumentSelect?.(doc.name)}
                    >
                      {doc.name}
                    </MenuItem>
                  ))}
                </>
              )}

              {/* 已有会话列表 */}
              {sessions.length > 0 && (
                <>
                  <MenuDivider />
                  <Text
                    size={100}
                    style={{
                      display: "block",
                      padding: "4px 12px 2px",
                      color: "#666",
                      fontSize: "11px",
                    }}
                  >
                    历史会话
                  </Text>
                  {sessions.map((s) => (
                    <MenuItem
                      key={s.session_id}
                      icon={<Chat20Regular />}
                      secondaryContent={
                        s.session_id !== activeSessionId ? (
                          <Delete20Regular
                            style={{ cursor: "pointer", color: "#999" }}
                            onClick={(e) => {
                              e.stopPropagation();
                              onDeleteSession(s.session_id);
                            }}
                          />
                        ) : undefined
                      }
                      style={
                        s.session_id === activeSessionId
                          ? { backgroundColor: "#eff6fc", fontWeight: 600 }
                          : undefined
                      }
                      onClick={() =>
                        onSessionChange(s.session_id, s.document_name)
                      }
                    >
                      {s.document_name || "未打开文档"}
                    </MenuItem>
                  ))}
                </>
              )}
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>
    </div>
  );
};

export default Header;
