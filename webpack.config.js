const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

module.exports = [
	{
		// Browser bundle configuration
		entry: "./src/odsp-app.ts",
		mode: "development",
		output: {
			filename: "[name].[contenthash].js",
			path: path.resolve(__dirname, "dist"),
		},
		plugins: [
			new webpack.ProvidePlugin({
				process: "process/browser",
			}),
			new HtmlWebpackPlugin({ template: "./src/index.html" }),
		],
		module: {
			rules: [
				{
					test: /\.tsx?$/,
					use: "ts-loader",
					exclude: /node_modules/,
				},
			],
		},
		resolve: {
			extensions: [".ts", ".tsx", ".js"],
			fallback: {
				http: false,
				fs: false,
				constants: false,
			},
		},
		devServer: {
			open: false,
		},
	},
	{
		target: "node", // Tells webpack to target a Node.js environment
		mode: "development",
		entry: "./src/odsp-client/index.ts", // The entry point for your application
		output: {
			filename: "[name].[contenthash].js",
			path: path.resolve(__dirname, "dist"),
		},
		resolve: {
			extensions: [".ts", ".js"], // Allows importing both .ts and .js files without specifying extension
		},
		module: {
			rules: [
				{
					test: /\.ts$/, // Use ts-loader for .ts files
					use: "ts-loader",
				},
			],
		},
	},
];
