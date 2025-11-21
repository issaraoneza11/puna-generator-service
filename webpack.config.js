const path = require('path');
const webpack = require('webpack');
var nodeExternals = require('webpack-node-externals');

module.exports = {
  mode: 'development',
  entry: './/app/bin/www.js',
  plugins: [new webpack.ProgressPlugin()],
  target: 'node',
  node:{
    __dirname: true
  },
  externals: [nodeExternals()],
  module: {
    rules: [{
      test: /\.(js|jsx)$/,
      include: [path.resolve(__dirname, '/app/bin')],
      exclude:[path.resolve(__dirname, 'node_modules')],
      loader: 'babel-loader'
    }]
  },

}