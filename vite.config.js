import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Em `build`, caminhos relativos (`./assets/...`) evitam página em branco no GitHub Pages
// quando o URL real do site não coincide com `/Automata` (maiúsculas, redirecionamentos).
// Em `dev`, `base` é `/` para o servidor em http://localhost:5173/
export default defineConfig(({ command }) => ({
  plugins: [react()],
  base: command === "build" ? "./" : "/",
}));
