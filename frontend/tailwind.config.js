/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      colors: {
        brand: {
          50: '#fff1ea',
          100: '#ecd0c2',
          200: '#c4a894',
          300: '#7a828a',
          400: '#5a626a',
          500: '#3a4a5a',
        },
      },
    },
  },
  plugins: [],
}
