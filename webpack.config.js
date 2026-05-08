const path = require("path");
const http = require("http");
const https = require("https");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const { getHttpsServerOptions } = require("office-addin-dev-certs");

function proxyCapivRequest(req, res) {
  const requestUrl = new URL(req.url, "https://localhost");
  const targetRaw = requestUrl.searchParams.get("url");
  if (!targetRaw) {
    res.writeHead(400);
    res.end("Missing url");
    return;
  }

  let target;
  try {
    target = new URL(targetRaw);
  } catch {
    res.writeHead(400);
    res.end("Invalid url");
    return;
  }

  if (!["datos.gob.ar", "datos.energia.gob.ar"].includes(target.hostname)) {
    res.writeHead(403);
    res.end("Host not allowed");
    return;
  }

  const client = target.protocol === "https:" ? https : http;
  const upstream = client.get(
    target,
    {
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
      }
    },
    (upstreamRes) => {
      const headers = {
        ...upstreamRes.headers,
        "access-control-allow-origin": "*"
      };
      res.writeHead(upstreamRes.statusCode || 502, headers);
      upstreamRes.pipe(res);
    }
  );

  upstream.on("error", (error) => {
    res.writeHead(502);
    res.end(error.message);
  });
}

module.exports = async (_env, argv) => {
  const mode = argv.mode || "development";
  const isDev = mode === "development";
  const httpsOptions = isDev ? await getHttpsServerOptions() : undefined;

  return {
    mode,
    devtool: isDev ? "inline-source-map" : "source-map",
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts"
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"]
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/index.html",
        chunks: ["taskpane"]
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"]
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "manifest.xml", to: "manifest.xml" },
          { from: "assets", to: "assets", noErrorOnMissing: false }
        ]
      })
    ],
    devServer: {
      hot: false,
      server: {
        type: "https",
        options: httpsOptions
      },
      port: Number(process.env.PORT || 3002),
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      allowedHosts: "all",
      setupMiddlewares: (middlewares, devServer) => {
        devServer.app.get("/capiv-proxy", proxyCapivRequest);
        return middlewares;
      }
    }
  };
};
