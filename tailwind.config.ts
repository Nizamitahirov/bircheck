import type { Config } from "tailwindcss";

const config: Config = {
  content: ["./app/**/*.{js,ts,jsx,tsx,mdx}", "./components/**/*.{js,ts,jsx,tsx,mdx}"],
  theme: {
    extend: {
      fontFamily: {
        sans: ["Inter", "system-ui", "sans-serif"],
      },
      colors: {
        brand: {
          50: "#eef4ff",
          100: "#d9e6ff",
          200: "#bcd3ff",
          300: "#8eb5ff",
          400: "#598dff",
          500: "#3366ff",
          600: "#1f47e6",
          700: "#1838b8",
          800: "#16308f",
          900: "#152c72",
        },
      },
      boxShadow: {
        soft: "0 10px 30px -10px rgba(31, 71, 230, 0.25)",
      },
    },
  },
  plugins: [],
};
export default config;
