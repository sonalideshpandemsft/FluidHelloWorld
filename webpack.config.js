/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

// // webpack config working for Azure apps
// const { CleanWebpackPlugin } = require("clean-webpack-plugin");
// const HtmlWebpackPlugin = require("html-webpack-plugin");
// const webpack = require("webpack");

// module.exports = (env) => {
// 	const htmlTemplate = "./src/index.html";
// 	const plugins =
// 		env && env.clean
// 			? [new CleanWebpackPlugin(), new HtmlWebpackPlugin({ template: htmlTemplate })]
// 			: [new HtmlWebpackPlugin({ template: htmlTemplate })];

// 	const mode = env && env.prod ? "production" : "development";

// 	return {
// 		devtool: "inline-source-map",
// 		entry: {
// 			app: "./src/odsp-app.ts",
// 		},
// 		mode,
// 		output: {
// 			filename: "[name].[contenthash].js",
// 		},
// 		plugins,
// 		module: {
// 			rules: [
// 				{
// 					test: /\.tsx?$/,
// 					use: "ts-loader",
// 					exclude: /node_modules/,
// 				},
// 			],
// 		},
// 		resolve: {
// 			extensions: [".ts", ".tsx", ".js"],
// 			fallback: {
// 				http: false,
// 				fs: false,
// 			},
// 		},
// 		devServer: {
// 			open: false,
// 		},
// 	};
// };

// odsp client webpack
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

module.exports = (env) => {
	const htmlTemplate = "./src/index.html";

	const mode = env && env.prod ? "production" : "development";

	return {
		devtool: "inline-source-map",
		entry: {
			app: "./src/odsp-app.ts",
		},
		mode,
		output: {
			filename: "[name].[contenthash].js",
		},
		plugins: [
			new webpack.ProvidePlugin({
				process: "process/browser",
			}),
			new HtmlWebpackPlugin({ template: htmlTemplate }),
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
			},
		},
		devServer: {
			open: false,
		},
	};
};
