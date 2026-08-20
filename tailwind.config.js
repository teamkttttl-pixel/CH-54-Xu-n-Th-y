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
        // Xanh dương — thẻ trắng nổi trên nền xanh nhạt.
        brand: {
          50:  "#EFF6FE",
          100: "#DBEAFE",
          200: "#BEDBFC",
          300: "#93C5FD",
          400: "#5FA5F5",
          500: "#3B82F6",
          600: "#1D6FD6",   // nút chính, mục đang chọn — chữ trắng 4,9:1
          700: "#175BAF",
          800: "#16447F",
          900: "#132F55",
        },
        // Nền trang
        canvas: "#EAF3FC",
      },
      boxShadow: {
        card: "0 1px 2px 0 rgb(19 47 85 / 0.04), 0 1px 3px 0 rgb(19 47 85 / 0.06)",
        lift: "0 4px 14px -4px rgb(19 47 85 / 0.12)",
      },
    },
  },
  plugins: [],
}
