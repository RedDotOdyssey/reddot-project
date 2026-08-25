import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: "/" 用于 Vercel/Netlify 这类自定义域名部署
// 如果改用 GitHub Pages（网址形如 https://用户名.github.io/仓库名/），
// 需要把下面这行改成 base: "/仓库名/"
export default defineConfig({
  plugins: [react()],
  base: "/",
});
