import type { Config } from 'tailwindcss';

const config: Config = {
  content: [
    './app/**/*.{js,ts,jsx,tsx,mdx}',
    './components/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  theme: {
    extend: {
      colors: {
        brand: {
          DEFAULT: '#3b5bdb',
          50: '#eef0fd',
          100: '#d5daf9',
          200: '#aab3f3',
          300: '#7f8ced',
          400: '#5466e7',
          500: '#3b5bdb',
          600: '#2f49b0',
          700: '#233785',
          800: '#17255a',
          900: '#0b122f',
        },
      },
    },
  },
  plugins: [],
};

export default config;
