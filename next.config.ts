import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  ...(process.env.AZURE_STATIC_EXPORT === "1" ? { output: "export" as const } : {}),
  ...(process.env.AZURE_STATIC_EXPORT === "1" ? { distDir: ".next-azure" } : {}),
};

export default nextConfig;
