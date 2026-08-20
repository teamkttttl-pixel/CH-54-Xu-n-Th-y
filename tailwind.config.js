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
        // Xanh mực — đậm và trầm hơn xanh mặc định, hợp phần mềm sổ sách.
        // Giữ nguyên thang 50–900 nên mọi class brand-* cũ vẫn chạy.
        brand: {
          50:  "#EEF3FA",
          100: "#D8E3F2",
          200: "#B2C6E3",
          300: "#83A2CD",
          400: "#547DB3",
          500: "#335E97",
          600: "#24487A",   // nút chính
          700: "#1B3760",
          800: "#142948",
          900: "#0E1D33",   // nền thanh bên
        },
      },
      boxShadow: {
        card: "0 1px 2px 0 rgb(16 30 51 / 0.04), 0 1px 3px 0 rgb(16 30 51 / 0.06)",
        lift: "0 4px 12px -2px rgb(16 30 51 / 0.10)",
      },
    },
  },
  plugins: [],
}
