const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const { getHttpsServerOptions } = require("office-addin-dev-certs");

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
      allowedHosts: "all"
    }
  };
};
