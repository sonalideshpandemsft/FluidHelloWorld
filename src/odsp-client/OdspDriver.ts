/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import assert from "assert";
import os from "os";
import { getDriveId, IClientConfig } from "@fluidframework/odsp-doclib-utils";
import type {
	OdspResourceTokenFetchOptions,
	HostStoragePolicy,
} from "@fluidframework/odsp-driver-definitions";
import {
	getMicrosoftConfiguration,
	OdspTokenConfig,
	OdspTokenManager,
	odspTokensCache,
} from "@fluidframework/tool-utils";

const passwordTokenConfig = (username: string, password: string): OdspTokenConfig => ({
	type: "password",
	username,
	password,
});

interface IOdspTestLoginInfo {
	siteUrl: string;
	username: string;
	password: string;
	supportsBrowserAuth?: boolean;
}

type TokenConfig = IOdspTestLoginInfo & IClientConfig;

export interface IOdspTestDriverConfig extends TokenConfig {
	directory: string;
	driveId: string;
	options: HostStoragePolicy | undefined;
}

/**
 * This class is copied from @fluidframework/test-driver's OdspTestDriver, with
 * only the minimal functionality we need retained, and any additional
 * functionality added
 */
export class OdspDriver {
	// Share the tokens and driverId across multiple instance of the test driver
	private static readonly odspTokenManager = new OdspTokenManager(odspTokensCache);
	private static readonly driveIdPCache = new Map<string, Promise<string>>();
	private static async getDriveIdFromConfig(tokenConfig: TokenConfig): Promise<string> {
		const siteUrl = tokenConfig.siteUrl;
		try {
			return await getDriveId(siteUrl, "", undefined, {
				accessToken: await this.getStorageToken({ siteUrl, refresh: false }, tokenConfig),
				refreshTokenFn: async () =>
					this.getStorageToken(
						{ siteUrl, refresh: true, useBrowserAuth: true },
						tokenConfig,
					),
			});
		} catch (ex) {
			if (tokenConfig.supportsBrowserAuth !== true) {
				throw ex;
			}
		}
		return getDriveId(siteUrl, "", undefined, {
			accessToken: await this.getStorageToken(
				{ siteUrl, refresh: false, useBrowserAuth: true },
				tokenConfig,
			),
			refreshTokenFn: async () =>
				this.getStorageToken({ siteUrl, refresh: true, useBrowserAuth: true }, tokenConfig),
		});
	}

	private static async getDriveId(siteUrl: string, tokenConfig: TokenConfig) {
		let driveIdP = this.driveIdPCache.get(siteUrl);
		if (driveIdP) {
			return driveIdP;
		}

		driveIdP = this.getDriveIdFromConfig(tokenConfig);
		this.driveIdPCache.set(siteUrl, driveIdP);
		try {
			return await driveIdP;
		} catch (e) {
			this.driveIdPCache.delete(siteUrl);
			throw e;
		}
	}

	public static async createFromEnv(config?: {
		directory?: string;
		username?: string;
		options?: HostStoragePolicy;
		supportsBrowserAuth?: boolean;
	}) {
		const loginAccounts = process.env.login__odsp__test__accounts;
		assert(loginAccounts !== undefined, "Missing login__odsp__test__accounts");
		// Expected format of login__odsp__test__accounts is simply string key-value pairs of username and password
		const passwords: { [user: string]: string } = JSON.parse(loginAccounts);
		const username = config?.username ?? Object.keys(passwords)[0];
		assert(passwords[username], `No password for username: ${username}`);

		const emailServer = username.substr(username.indexOf("@") + 1);
		let siteUrl: string;
		let tenantName: string;
		if (emailServer.startsWith("http://") || emailServer.startsWith("https://")) {
			// it's already a site url
			tenantName = new URL(emailServer).hostname;
			siteUrl = emailServer;
		} else {
			tenantName = emailServer.substr(0, emailServer.indexOf("."));
			siteUrl = `https://${tenantName}.sharepoint.com`;
		}

		// force isolateSocketCache because we are using different users in a single context
		// and socket can't be shared between different users
		const options = config?.options ?? {};
		options.isolateSocketCache = true;

		console.log("create env server----", siteUrl);

		return this.create(
			{
				username,
				password: passwords[username],
				siteUrl,
				supportsBrowserAuth: config?.supportsBrowserAuth,
			},
			config?.directory ?? "",
			config?.options,
		);
	}

	// use this directly instead
	private static async create(
		loginConfig: IOdspTestLoginInfo,
		directory: string,
		options?: HostStoragePolicy,
	) {
		const tokenConfig: TokenConfig = {
			...loginConfig,
			...getMicrosoftConfiguration(),
		};

		console.log("create env tokenConfig----", tokenConfig);

		console.log("create env directoryParts----");

		const driveId = await this.getDriveId(loginConfig.siteUrl, tokenConfig);
		const directoryParts = [directory];

		// if we are in a azure dev ops build use the build id in the dir path
		if (process.env.BUILD_BUILD_ID !== undefined) {
			directoryParts.push(process.env.BUILD_BUILD_ID);
		} else {
			directoryParts.push(os.hostname());
		}

		const driverConfig: IOdspTestDriverConfig = {
			...tokenConfig,
			directory: directoryParts.join("/"),
			driveId,
			options,
		};

		console.log("create env driverConfig----", driverConfig);

		return new OdspDriver(driverConfig);
	}

	private static async getStorageToken(
		options: OdspResourceTokenFetchOptions & { useBrowserAuth?: boolean },
		config: IOdspTestLoginInfo & IClientConfig,
	) {
		console.log("site url-------", options.siteUrl);
		const host = new URL(options.siteUrl).host;

		console.log("site url host-------", host);

		if (options.useBrowserAuth === true) {
			console.log("site url 1-------");
			const browserTokens = await this.odspTokenManager.getOdspTokens(
				host,
				config,
				{
					type: "browserLogin",
					navigator: (openUrl) => {
						console.log(
							`Open the following url in a new private browser window, and login with user: ${config.username}`,
						);
						console.log(
							`Additional account details may be available in the environment variable login__odsp__test__accounts`,
						);
						console.log(`"${openUrl}"`);
					},
				},
				options.refresh,
			);
			return browserTokens.accessToken;
		}
		console.log("site url 2-------", passwordTokenConfig(config.username, config.password));
		// This function can handle token request for any multiple sites.
		// Where the test driver is for a specific site.
		const tokens = await this.odspTokenManager.getOdspTokens(
			host,
			config,
			passwordTokenConfig(config.username, config.password),
			options.refresh,
		);
		console.log("site url 2-------", passwordTokenConfig(config.username, config.password));
		return tokens.accessToken;
	}

	public get siteUrl(): string {
		return `https://${this.config.siteUrl}`;
	}
	public get driveId(): string {
		return this.config.driveId;
	}
	public get directory(): string {
		return this.config.directory;
	}

	private constructor(private readonly config: Readonly<IOdspTestDriverConfig>) {}

	public readonly getStorageToken = async (options: OdspResourceTokenFetchOptions) => {
		return OdspDriver.getStorageToken(options, this.config);
	};
	public readonly getPushToken = async (options: OdspResourceTokenFetchOptions) => {
		const tokens = await OdspDriver.odspTokenManager.getPushTokens(
			new URL(options.siteUrl).hostname,
			this.config,
			passwordTokenConfig(this.config.username, this.config.password),
			options.refresh,
		);

		return tokens.accessToken;
	};
}
