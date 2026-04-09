import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const title = "智能文档助手";

const rootElement: HTMLElement | null = document.getElementById("container");
const root: Root | undefined = rootElement ? createRoot(rootElement) : undefined;

const appTree = (
  <FluentProvider
    theme={webLightTheme}
    style={{
      display: "flex",
      flexDirection: "column",
      flex: 1,
      minHeight: 0,
      height: "100%",
      overflow: "hidden",
    }}
  >
    <App title={title} />
  </FluentProvider>
);

function renderApp() {
  root?.render(appTree);
}

/* Render after Office.js 就绪；若 Office 未注入则直接渲染，避免 ReferenceError 导致整页白屏 */
let rendered = false;
function tryRenderOnce() {
  if (rendered || !root) return;
  rendered = true;
  renderApp();
}

const officeGlobal = (typeof globalThis !== "undefined" && (globalThis as unknown as { Office?: unknown }).Office) as
  | { onReady: (cb: () => void) => void }
  | undefined;

if (officeGlobal && typeof officeGlobal.onReady === "function") {
  officeGlobal.onReady(() => tryRenderOnce());
  window.setTimeout(() => {
    if (!rendered) {
      console.warn("[wordassistant] Office.onReady 超时，仍尝试渲染界面（部分 Word API 可能不可用）");
      tryRenderOnce();
    }
  }, 8000);
} else {
  console.warn("[wordassistant] 未检测到 Office.js，在非宿主环境或脚本加载失败时仍可预览界面");
  tryRenderOnce();
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
