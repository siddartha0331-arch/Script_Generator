import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Vite's job: take our JSX/React source, bundle it, and spit out a plain
// dist/ folder of static HTML/CSS/JS. That static folder is literally all
// Netlify needs to serve — it doesn't run Node or React itself, it's just
// a very fast static file host.
export default defineConfig({
  plugins: [react()],
});
