import { defineConfig } from 'vite'
import react, { reactCompilerPreset } from '@vitejs/plugin-react'
import babel from '@rolldown/plugin-babel'

export default defineConfig({
  base: '/mockup-dw/',
  plugins: [
    react(),
    babel({ presets: [reactCompilerPreset()] })
  ],
})