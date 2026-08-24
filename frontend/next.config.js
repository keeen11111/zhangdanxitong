/** @type {import('next').NextConfig} */
const nextConfig = {
  output: "standalone",
  // Keep a running dev server isolated from production/desktop builds. Next
  // otherwise rewrites the same .next module manifest and produces missing
  // chunk errors while users are uploading files.
  distDir: process.env.NEXT_DIST_DIR || (process.env.NODE_ENV === "development" ? ".next-dev" : ".next"),
  // 允许开发期跨域请求后端
  async rewrites() {
    return [];
  },
};

module.exports = nextConfig;
