const path = require('path');
const webpack = require('webpack');
module.exports = {
    entry: './myApp.ts',
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
    plugins: [],
    output: {
        filename: 'myApp.js',
        path: path.resolve(__dirname, './')
    }
};