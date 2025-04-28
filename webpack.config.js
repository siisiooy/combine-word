const path = require('path');

module.exports = {
  entry: './src/index.js',  // 入口文件
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: 'combine-word.js',
    library: 'CombineWord',  // 全局挂载的变量
    libraryTarget: 'umd',  // 支持多种模块化方式 (CommonJS, AMD, browser)
    globalObject: 'this',  // 解决在 Node.js 和浏览器环境中通用的问题
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: 'babel-loader',
      },
    ],
  },
  resolve: {
    extensions: ['.js'],
  },
  devtool: 'source-map',
  externals: {
    'jszip': 'JSZip',
  }
};
