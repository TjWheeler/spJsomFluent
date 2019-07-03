var path = require('path');
var webpack = require('webpack');
var version = JSON.stringify(require("./package.json").version);
var banner = 'spJsomFluent ' + version + ' - https://github.com/TjWheeler/spJsomFluent';
function DtsBundlePlugin() { }
DtsBundlePlugin.prototype.apply = function (compiler) {
    compiler.plugin('done', function () {
        var dts = require('dts-bundle');
        dts.bundle({
            name: 'spJsomFluent',
            main: './build/src/**/*.d.ts',
            out: '../../dist/index.d.ts',
            removeSource: true,
            outputAsModuleFolder: true
        });
    });
};
var baseConfiguration = {
    entry: { spJsomFluent: './src/fluent.ts', spJsomExamples: './examples/spJsomExamples-typescript.ts' },
    devtool: 'source-map',
    module: {
        rules: [
            {
                test: /\.ts?$/,
                use: 'ts-loader',
                exclude: /node_modules/
            }
        ]
    },
    resolve: {
        extensions: ['.ts']
    }
};

var libraryConfig = {
    ...baseConfiguration, ...
    {
        output: {
            filename: '[name].js',
            path: path.resolve(__dirname, './dist'),
            libraryTarget: 'var',
            library: 'spJsom'
        },
        plugins: [
            new webpack.BannerPlugin(banner)
            ]
        
    }
};
var umdConfig = {
    ...baseConfiguration, ...
    {
        output: {
            filename: '[name].umd.js',
            path: path.resolve(__dirname, './dist'),
            libraryTarget: 'umd',
            library: 'spJsom'
        },
        plugins: [
            new DtsBundlePlugin(),
            new webpack.BannerPlugin(banner)
            ]
    }
};
module.exports = [libraryConfig, umdConfig];