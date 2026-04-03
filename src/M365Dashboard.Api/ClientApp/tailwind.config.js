/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        // Microsoft Fluent Design colors
        'ms-blue': {
          50: '#e6f2ff',
          100: '#cce5ff',
          200: '#99cbff',
          300: '#66b0ff',
          400: '#3396ff',
          500: '#0078d4',
          600: '#0066b4',
          700: '#004d87',
          800: '#00335a',
          900: '#001a2d',
        },
        'ms-gray': {
          50: '#faf9f8',
          100: '#f3f2f1',
          200: '#edebe9',
          300: '#e1dfdd',
          400: '#d2d0ce',
          500: '#c8c6c4',
          600: '#a19f9d',
          700: '#605e5c',
          800: '#323130',
          900: '#201f1e',
        },
      },
      fontFamily: {
        sans: ['Segoe UI', 'system-ui', 'sans-serif'],
      },
    },
  },
  plugins: [],
}
