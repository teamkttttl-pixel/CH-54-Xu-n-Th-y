/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      fontFamily: {
        sans: ['"Be Vietnam Pro"', "system-ui", "-apple-system", "sans-serif"],
        mono: ['"IBM Plex Mono"', "ui-monospace", "SFMono-Regular", "monospace"],
      },
      colors: {
        // Nâu cà phê — không đụng màu nào trong hệ ngữ nghĩa của app
        // (xanh lá đã thu, vàng còn nợ, đỏ trả hàng, tím nội bộ, xanh dương NCC).
        brand: {
          50:  "#FAF6F2",
          100: "#F0E7DE",
          200: "#DFCBB6",
          300: "#C4A78C",
          400: "#A67A54",
          500: "#8C5C31",
          600: "#7C4A21",   // nút chính — chữ trắng đạt 7,4:1
          700: "#633A19",
          800: "#4A2C13",
          900: "#331E0D",
        },
      },
      boxShadow: {
        card: "0 1px 2px 0 rgb(51 30 13 / 0.04), 0 1px 3px 0 rgb(51 30 13 / 0.05)",
        lift: "0 4px 14px -4px rgb(51 30 13 / 0.12)",
      },
    },
  },
  plugins: [],
}
