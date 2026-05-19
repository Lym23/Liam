# Word 数字统计插件（初始版本）

这是一个基于 Office.js 的 Word 插件示例。它会在右下角（任务窗格底部）实时显示当前选中内容中的：

- 求和（Sum）
- 平均数（Average）
- 计数（Count）

## 功能说明

1. 支持在 Word 文档中选中单个或多个数字后，实时统计 Sum / Average / Count。
2. 支持千分位数字（如 `1,234,567.89`），统计时自动忽略逗号。

## 运行方式（本地）

1. 启动一个本地静态服务器（例如 `npx http-server . -p 3000 --cors`）。
2. 将 `manifest.xml` 旁加载（sideload）到 Word。
3. 打开任务窗格，选中文档中的数字，观察底部统计栏实时变化。

## 目录结构

- `manifest.xml`：Office Add-in 清单。
- `taskpane.html`：任务窗格页面。
- `taskpane.js`：统计逻辑与 Word 选择监听。
- `styles.css`：简易样式。
