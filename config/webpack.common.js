var path = require('path');
var webpack = require('webpack');
var HtmlWebpackPlugin = require('html-webpack-plugin');
var ExtractTextPlugin = require('extract-text-webpack-plugin');

module.exports = {
    entry: {
        'app': './src/app.tsx',
        'config': './src/config.tsx'
    },

    output: {
        filename: '[name].js',
        path: path.resolve(__dirname, "dist")
    },

    // Enable sourcemaps for debugging webpack's output.
    devtool: 'source-map',

    resolve: {
        // Add '.ts' and '.tsx' as resolvable extensions.
        extensions: ['', '.css', '.ts', '.tsx', '.js', '.jsx']
    },

    module: {
        loaders: [
            // All files with a '.ts' or '.tsx' extension will be handled by 'ts-loader'.
            {
                test: /\.tsx?$/,
                loader: 'ts-loader',
                exclude: /node_modules/
            },

            // All files with a '.css' extension will be handled by 'extract-text-plugin and css-loader'.
            {
                test: /\.css$/,
                loader: ExtractTextPlugin.extract('css')
            }
        ],

        preLoaders: [
            // All output '.js' files will have any sourcemaps re-processed by 'source-map-loader'.
            { test: /\.js$/, loader: 'source-map-loader' }
        ]
    },

    plugins: [
        new HtmlWebpackPlugin({
            title: 'React • TodoMVC',
            filename: 'index.html',
            template: path.resolve('index.html'),
            chunks: ['app']
        }),

        new HtmlWebpackPlugin({
            title: 'Config • TodoMVC',
            filename: 'config.html',
            template: path.resolve('config.html'),
            chunks: ['config']
        }),

        new ExtractTextPlugin('[name].css'),

        new webpack.ProvidePlugin({
            Router: 'director'
        })
    ]
};