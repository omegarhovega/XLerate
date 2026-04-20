/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = process.env.XLERATE_PROD_URL || "https://omegarhovega.github.io/XLerate/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      traceDialog: ["./src/taskpane/traceDialog.ts", "./src/taskpane/traceDialog.html"],
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets",
            // The icon-generation script (`npm run icons`) drops the
            // source SVGs alongside the rendered PNGs in assets/icons/
            // for debugging. They're never loaded at runtime — keep
            // them out of the shipped bundle.
            globOptions: { ignore: ["**/*.svg"] },
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "traceDialog.html",
        template: "./src/taskpane/traceDialog.html",
        // Dialog omits the `polyfill` chunk deliberately: it only targets
        // modern Excel hosts (Edge WebView2 / Chromium Online), the entry
        // code uses no features that core-js shims on this host, and
        // dropping it shaves parse time off the 1–2 s cold dialog spawn
        // measured in sideload. The main taskpane still ships polyfills
        // because it runs in the same iframe as legacy-compat code.
        chunks: ["traceDialog"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      host: "0.0.0.0",
      allowedHosts: "all",
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      // The full-screen client overlay injected by webpack-dev-server blocks
      // interaction inside the taskpane iframe AND the Office Dialog window
      // on every warning / runtime error — making it impossible to judge
      // live-nav smoothness in sideload. Disabled here; use the Edge DevTools
      // console (right-click taskpane → Inspect) for error visibility during
      // development. Re-enable selectively with
      //   client: { overlay: { errors: true, warnings: false, runtimeErrors: false } }
      // if a future change needs it.
      client: {
        overlay: false,
      },
    },
  };

  return config;
};
