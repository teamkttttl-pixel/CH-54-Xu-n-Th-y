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
        // Xám đá — dựng quanh #94A3B8 (bậc 300). Không phải xám thuần:
        // có chút ánh xanh nên không bị chết và không tranh chỗ với hệ màu
        // ngữ nghĩa (xanh lá đã thu, vàng còn nợ, đỏ trả hàng, tím nội bộ).
        brand: {
          50:  "#F8FAFC",
          100: "#EEF2F6",
          200: "#CBD5E1",
          300: "#94A3B8",   // màu đã chọn
          400: "#7C8BA1",
          500: "#64748B",
          600: "#4A5A6E",   // nút chính — chữ trắng đạt 7,0:1
          700: "#334155",
          800: "#26303F",
          900: "#1B222D",
        },
      },
      boxShadow: {
        card: "0 1px 2px 0 rgb(27 34 45 / 0.03), 0 1px 3px 0 rgb(27 34 45 / 0.04)",
        lift: "0 4px 14px -4px rgb(27 34 45 / 0.10)",
      },
    },
  },
  plugins: [],
}
