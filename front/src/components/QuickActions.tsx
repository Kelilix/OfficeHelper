import * as React from "react";
import { Button } from "@fluentui/react-components";
import {
  TextT24Regular,
  TextFont24Regular,
  TextAlignLeft24Regular,
  TextHeader2Regular,
} from "@fluentui/react-icons";

interface QuickActionsProps {
  onSend: (message: string) => void;
  disabled?: boolean;
}

const ACTIONS = [
  { label: "黑体", icon: <TextFont24Regular />, prompt: "将选中的文字设为黑体" },
  { label: "三号字", icon: <TextT24Regular />, prompt: "将选中的文字字号设为三号" },
  { label: "加粗", icon: <TextT24Regular />, prompt: "将选中的文字设为加粗" },
  { label: "首行缩进", icon: <TextAlignLeft24Regular />, prompt: "将选中段落设为首行缩进2字符" },
  { label: "添加标题", icon: <TextHeader2Regular />, prompt: "将选中段落设为标题1样式" },
  { label: "两端对齐", icon: <TextAlignLeft24Regular />, prompt: "将选中段落设为两端对齐" },
];

const QuickActions: React.FC<QuickActionsProps> = ({ onSend, disabled }) => {
  const [visible, setVisible] = React.useState(false);

  return (
    <div style={{ padding: "4px 16px 0" }}>
      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: "4px",
          paddingBottom: visible ? "8px" : 0,
          overflow: "hidden",
          transition: "all 0.2s",
          maxHeight: visible ? "200px" : "0px",
        }}
      >
        {ACTIONS.map((action) => (
          <Button
            key={action.label}
            size="small"
            appearance="subtle"
            icon={action.icon}
            onClick={() => onSend(action.prompt)}
            disabled={disabled}
            style={{ fontSize: "12px" }}
          >
            {action.label}
          </Button>
        ))}
      </div>
      <Button
        size="small"
        appearance="subtle"
        onClick={() => setVisible((v) => !v)}
        style={{ fontSize: "11px", color: "#999", padding: "0" }}
      >
        {visible ? "收起快捷指令" : "显示快捷指令"}
      </Button>
    </div>
  );
};

export default QuickActions;
