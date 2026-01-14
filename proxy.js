const http = require("http");
const https = require("https");
const url = require("url");

const PORT = 3002;
const ADS_API_BASE = "https://api.adsabs.harvard.edu";

const server = http.createServer((req, res) => {
  // CORS headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type");

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  // Only proxy /v1/* paths
  if (!req.url.startsWith("/v1/")) {
    res.writeHead(404);
    res.end("Not found");
    return;
  }

  const targetUrl = ADS_API_BASE + req.url;
  const parsedUrl = url.parse(targetUrl);

  const options = {
    hostname: parsedUrl.hostname,
    port: 443,
    path: parsedUrl.path,
    method: req.method,
    headers: {
      ...req.headers,
      host: parsedUrl.hostname,
    },
  };

  const proxyReq = https.request(options, (proxyRes) => {
    res.writeHead(proxyRes.statusCode, proxyRes.headers);
    proxyRes.pipe(res);
  });

  proxyReq.on("error", (err) => {
    console.error("Proxy error:", err.message);
    res.writeHead(502);
    res.end("Proxy error");
  });

  req.pipe(proxyReq);
});

server.listen(PORT, () => {
  console.log(`ADS proxy running at http://localhost:${PORT}`);
  console.log(`Proxying to ${ADS_API_BASE}`);
});
