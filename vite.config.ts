import { cloudflare } from "@cloudflare/vite-plugin";
import tailwindcss from "@tailwindcss/vite";
import { tanstackStart } from "@tanstack/react-start/plugin/vite";
import { defineConfig } from "vite";
import viteReact from "@vitejs/plugin-react";
import tsconfigPaths from "vite-tsconfig-paths";

export default defineConfig(({ command }) => {
  const useCloudflareRuntime = command === "build";

  if (!useCloudflareRuntime) {
    // Keep local dev off the Workers runtime; this avoids Miniflare/Cloudflare TLS
    // setup on Windows while still letting the production build keep its adapter.
    process.env.CLOUDFLARE_CF_FETCH_ENABLED = "false";
  }

  return {
    plugins: [
      ...(useCloudflareRuntime
        ? [
            cloudflare({
              viteEnvironment: { name: "ssr" },
              remoteBindings: false,
            }),
          ]
        : []),
      tailwindcss(),
      tsconfigPaths(),
      tanstackStart(),
      viteReact(),
    ],
  };
});
