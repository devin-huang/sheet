const CompressionPlugin = require("compression-webpack-plugin");
module.exports = {
  publicPath: process.env.NODE_ENV === "production" ? "/" : "/",
  configureWebpack: () => {
    var obj = {
      // 防止引用xlsx-style组件报错
      externals: {
        "./cptable": "var cptable",
        "../xlsx.js": "var _XLSX",
      },
    };
    if (process.env.NODE_ENV == "production") {
      obj.plugins = [
        new CompressionPlugin({
          test: /\.js$|\.html$|\.css/,
          threshold: 10240,
          deleteOriginalAssets: false,
        }),
      ];
    }
    return obj;
  },
};
