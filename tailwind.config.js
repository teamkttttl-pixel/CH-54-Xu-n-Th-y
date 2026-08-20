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
        // Xám đậm — trung tính có chút ánh xanh, không tranh chỗ với hệ màu
        // ngữ nghĩa (xanh lá đã thu, vàng còn nợ, đỏ trả hàng, tím nội bộ).
        brand: {
          50:  "#F5F7FA",
          100: "#E8EDF2",
          200: "#CFD8E3",
          300: "#93A3B5",
          400: "#6B7D92",
          500: "#4C5F76",
          600: "#35475C",   // nút chính — chữ trắng đạt 9,0:1
          700: "#2A3949",
          800: "#1E2833",
          900: "#151C25",
        },
      },
      boxShadow: {
        card: "0 1px 2px 0 rgb(21 28 37 / 0.04), 0 1px 3px 0 rgb(21 28 37 / 0.05)",
        lift: "0 4px 14px -4px rgb(21 28 37 / 0.12)",
      },
    },
  },
  plugins: [],
}
