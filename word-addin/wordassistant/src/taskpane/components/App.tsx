import * as React from "react";
import { makeStyles, tokens, Text } from "@fluentui/react-components";
import AIChat from "./AIChat";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    flex: 1,
    minHeight: 0,
    height: "100%",
    maxHeight: "100%",
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    flexShrink: 0,
    padding: "12px 16px 10px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    alignItems: "center",
    gap: "10px",
  },
  logo: {
    width: "28px",
    height: "28px",
    borderRadius: "6px",
    objectFit: "contain",
  },
  titleArea: {
    flex: 1,
    minWidth: 0,
  },
  chatHost: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    minHeight: 0,
    overflow: "hidden",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <img className={styles.logo} src="assets/logo-filled.png" alt="Logo" />
        <div className={styles.titleArea}>
          <Text weight="semibold" size={400}>
            {props.title}
          </Text>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            AI 智能文档助手 · 对话
          </Text>
        </div>
      </div>

      <div className={styles.chatHost}>
        <AIChat />
      </div>
    </div>
  );
};

export default App;
