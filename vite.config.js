import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { viteSingleFile } from "vite-plugin-singlefile";
// 모든 JS/CSS를 index.html 한 파일에 넣어, 더블클릭(file://)으로도 열리게 함
export default defineConfig({ plugins: [react(), viteSingleFile()], base: "./", build: { chunkSizeWarningLimit: 2000 } });
