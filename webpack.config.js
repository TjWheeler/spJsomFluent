const path = require('path');
const webpack = require('webpack');
const version = 'v0.0.7';
var banner = 'spJsomFluent ' + version + ' - https://github.com/TjWheeler/spJsomFluent';
module.exports = {
    entry: { spJsomFluent: './src/fluent.ts'},
    devtool: 'source-map',
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                use: 'ts-loader',
                exclude: path.resolve(__dirname, '/node_modules')
            }
        ]
    },
    resolve: {
        extensions: ['.ts']
    },
    plugins: [new webpack.BannerPlugin(banner)],
    output: {
        filename: '[name].js',
        path: path.resolve(__dirname, './dist'),
        libraryTarget: 'var',
        library: 'spJsom'
    }
};