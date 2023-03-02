const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

module.exports = [
	{
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
			new webpack.DefinePlugin({
				"process.env.login__odsp__test__tenants": JSON.stringify(
					"login__odsp__test__tenants",
				),
				"process.env.login__odspdf__test__tenants": JSON.stringify(
					"login__odspdf__test__tenants",
				),
				"process.env.login__odsp__test__accounts": JSON.stringify(
					"login__odsp__test__accounts",
				),
				"process.env.login__microsoft__clientId": JSON.stringify(
					"login__microsoft__clientId",
				),
				"process.env.login__microsoft__secret": JSON.stringify("login__microsoft__secret"),
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
];
