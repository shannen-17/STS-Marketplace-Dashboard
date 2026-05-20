import type { Config } from "tailwindcss";

const config: Config = {
  content: ["./app/**/*.{js,ts,jsx,tsx}", "./components/**/*.{js,ts,jsx,tsx}"],
  theme: {
    extend: {
      fontFamily: { sans: ["Inter", "ui-sans-serif", "system-ui"] },
      colors: {
        sts: {
          navy: "#102A43",
          blue: "#1F6FEB",
          sky: "#EAF4FF",
          cream: "#FFF8EF",
          gold: "#C8912D"
        }
      }
    }
  },
  plugins: []
};
export default config;