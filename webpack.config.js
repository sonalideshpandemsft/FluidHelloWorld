/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

// webpack config working for Azure apps
// const { CleanWebpackPlugin } = require("clean-webpack-plugin");
// const HtmlWebpackPlugin = require("html-webpack-plugin");

// module.exports = (env) => {
//     const htmlTemplate = "./src/index.html";
//     const plugins =
//         env && env.clean
//             ? [new CleanWebpackPlugin(), new HtmlWebpackPlugin({ template: htmlTemplate })]
//             : [new HtmlWebpackPlugin({ template: htmlTemplate })];

//     const mode = env && env.prod ? "production" : "development";

//     return {
//         devtool: "inline-source-map",
//         // target: "node",
//         entry: {
//             app: "./src/app.ts",
//         },
//         mode,
//         output: {
//             filename: "[name].[contenthash].js",
//         },
//         plugins,
//         module: {
//             rules: [
//                 {
//                     test: /\.tsx?$/,
//                     use: "ts-loader",
//                     exclude: /node_modules/,
//                 },
//             ],
//         },
//         resolve: {
//             extensions: [".ts", ".tsx", ".js"],
//         },
//         devServer: {
//             open: false,
//         },
//     };
// };

// webpack for odsp client
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const path = require("path");

module.exports = (env) => {
	const mode = env && env.prod ? "production" : "development";

	return {
		entry: "./src/odsp-app.ts",
		target: "node",
		mode,
		output: {
			filename: "odsp-app.js",
			path: path.resolve(__dirname, "dist"),
		},
		plugins: [
			new HtmlWebpackPlugin({ template: "./src/index.html" }),
			new CopyWebpackPlugin({
				patterns: [
					{
						from: "src/odsp-client",
						to: "odsp-client",
					},
				],
			}),
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
			extensions: [".tsx", ".ts", ".js"],
		},
		devServer: {
			contentBase: path.join(__dirname, "src"),
			compress: true,
			port: 9000,
		},
	};
};
