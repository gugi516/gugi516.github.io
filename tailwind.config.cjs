/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./src/**/*.{astro,html,js,jsx,md,mdx,svelte,ts,tsx,vue}'],
  darkMode: ['class', '[data-theme="shlee-dark"]'],
  theme: {
    extend: {
      fontFamily: {
        sans: ['"Pretendard Variable"', 'Pretendard', 'Inter', '-apple-system', 'BlinkMacSystemFont', 'system-ui', 'Roboto', '"Helvetica Neue"', '"Segoe UI"', '"Apple SD Gothic Neo"', '"Noto Sans KR"', '"Malgun Gothic"', 'sans-serif'],
        serif: ['"Playfair Display"', 'Georgia', 'serif'],
        mono: ['"JetBrains Mono"', 'monospace'],
      },
      colors: {
        shlee: {
          'carbon-black':  '#0c1420',
          'carbon':        '#13202f',
          'carbon-deep':   '#0f1b2a',
          'green-deep':    '#1e3a5f',
          'green':         '#2d5380',
          'green-mid':     '#3d68a0',
          'mint':          '#6fb0e0',
          'mint-soft':     '#a0cfec',
          'honey':         '#41a0e0',
          'honey-deep':    '#2074b8',
          'honey-soft':    '#8ecdf2',
          'paper':         '#eef2f8',
          'paper-deep':    '#dde6f0',
          /* Urban-landscape accent (secondary to the blue; ~8:2 blue:green) */
          'leaf-deep':     '#2e7d4f',
          'leaf':          '#4aa96c',
          'leaf-soft':     '#a8dcc0',
          'leaf-bright':   '#6fcf97',   /* dark-mode tone */
        },
      },
      keyframes: {
        'fade-up': {
          '0%': { opacity: '0', transform: 'translateY(8px)' },
          '100%': { opacity: '1', transform: 'translateY(0)' },
        },
      },
      animation: {
        'fade-up': 'fade-up 0.5s ease-out',
      },
    },
  },
  plugins: [require("@tailwindcss/typography"), require("daisyui")],
  daisyui: {
    themes: [
      {
        "shlee-light": {
          "primary":           "#1e3a5f",
          "primary-content":   "#ffffff",
          "secondary":         "#3d68a0",
          "secondary-content": "#ffffff",
          "accent":            "#41a0e0",
          "accent-content":    "#0f1b2a",
          "neutral":           "#1e3a5f",
          "neutral-content":   "#eef2f8",
          "base-100":          "#ffffff",
          "base-200":          "#eef2f8",
          "base-300":          "#c9d7e6",
          "base-content":      "#0f1b2a",
          "info":              "#3d68a0",
          "success":           "#3d68a0",
          "warning":           "#2074b8",
          "error":             "#ef4444",
        },
      },
      {
        "shlee-dark": {
          "primary":           "#6fb0e0",
          "primary-content":   "#0c1420",
          "secondary":         "#3d68a0",
          "secondary-content": "#e2e6ec",
          "accent":            "#41a0e0",
          "accent-content":    "#0c1420",
          "neutral":           "#13202f",
          "neutral-content":   "#e2e6ec",
          "base-100":          "#0c1420",
          "base-200":          "#13202f",
          "base-300":          "#24364e",
          "base-content":      "#e2e6ec",
          "info":              "#3d68a0",
          "success":           "#6fb0e0",
          "warning":           "#41a0e0",
          "error":             "#ef4444",
        },
      },
    ],
    darkTheme: "shlee-dark",
    logs: false,
  }
}
