const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const isProduction = process.env.NODE_ENV === "production";
const baseUrl = process.env.BASE_URL || "https://localhost:3000";

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.tsx",
    functions: "./src/functions/functions.ts",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].bundle.js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js", ".jsx"],
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/taskpane.html",
      filename: "taskpane.html",
      chunks: ["taskpane", "functions"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "manifest.xml",
          to: "manifest.xml",
          transform(content) {
            return content.toString().replace(/\{\{BASE_URL\}\}/g, baseUrl);
          },
        },
        { from: "assets", to: "assets", noErrorOnMissing: true },
        { from: "src/functions/functions.json", to: "functions.json" },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: {
      type: "https",
      options: {
        key: undefined, // office-addin-dev-certs provides these
        cert: undefined,
      },
    },
    port: 3000,
    hot: true,
  },
  devtool: isProduction ? "source-map" : "eval-source-map",
};
