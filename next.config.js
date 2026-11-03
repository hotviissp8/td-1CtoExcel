/** @type {import('next').NextConfig} */
// Fully static site (all processing happens in the browser).
// PAGES_BASE_PATH is set by the GitHub Pages workflow (e.g. "/td-1CtoExcel").
const basePath = process.env.PAGES_BASE_PATH || '';

const nextConfig = {
  output: 'export',
  basePath,
  images: { unoptimized: true },
};
module.exports = nextConfig;
