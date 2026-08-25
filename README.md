# 红点时光探索之旅 App — 测试项目

这是配好的 React + Vite + Tailwind 项目脚手架，`src/App.jsx` 就是我们这几周一起做的那个 App。

## 本地先跑一下（可选，确认代码没问题）

需要先装好 [Node.js](https://nodejs.org/)（建议 18 或以上版本），然后：

```bash
npm install
npm run dev
```

打开终端提示的网址（通常是 http://localhost:5173），能看到 App 正常运行就说明没问题。

## 推上 GitHub、部署测试

详细步骤见对话里的说明，简单版：

1. 在 GitHub 建一个新仓库
2. 把这个文件夹的内容推上去
3. 去 [vercel.com](https://vercel.com) 用 GitHub 账号登录，选这个仓库，点 Deploy
4. 几十秒后拿到一个 `https://xxx.vercel.app` 的网址，这就是可以发给朋友测试的正式网址（不再是 claude.ai 那个受限的预览链接）

## 关于 Google Sheet 报名同步

`src/App.jsx` 顶部有个 `GOOGLE_SHEET_WEBHOOK_URL` 常量，部署好 Apps Script 后把网址填进去即可（详见 `GoogleSheetBackend.gs` 文件里的部署说明）。
