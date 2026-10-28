# QingSi · 皮瓣状态智能辅助监测系统

面向皮瓣移植术后监测的多角色微信小程序 + 边缘数采 + 云平台系统，AI 接口预留；架构按“平台 + 领域包”设计，可复用到烟草在线监测等工业边云监测场景。

当前阶段：**开发前架构评审**。

## 文档

| 文档 | 说明 |
| --- | --- |
| [`docs/architecture.md`](docs/architecture.md) | 系统架构与开发组织架构设计说明（v0.1，待评审） |
| [`docs/系统架构与开发组织讲解.txt`](docs/系统架构与开发组织讲解.txt) | 文字讲解版：设计思路、逐模块讲解、汇报指南与问答准备 |
| [`docs/ppt/皮瓣状态智能辅助监测系统-架构与开发组织-v0.1.pptx`](docs/ppt/皮瓣状态智能辅助监测系统-架构与开发组织-v0.1.pptx) | 评审用 PPT（26 页，含演讲备注） |
| [`docs/ppt/build-deck.js`](docs/ppt/build-deck.js) | PPT 生成脚本 |

## 重新生成 PPT

```bash
cd docs/ppt
npm install
npm run build
```

修改 `build-deck.js` 中的文字或版式后重新运行即可。PPT 使用“微软雅黑”字体，建议在 Windows / macOS 的 PowerPoint 或 WPS 中打开。
