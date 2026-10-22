export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      // 깜빡임 효과 (드론 추격전 앱과 동일)
      keyframes: {
        blink: { "0%, 100%": { opacity: 1 }, "50%": { opacity: 0.15 } },
      },
      animation: {
        blink: "blink 1s ease-in-out infinite",
      },
    },
  },
  plugins: [],
};
