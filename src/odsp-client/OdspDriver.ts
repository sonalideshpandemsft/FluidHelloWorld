/* eslint-disable max-len */
/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
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
	OdspTokenConfig,
	OdspTokenManager,
	odspTokensCache,
	getMicrosoftConfiguration,
} from "@fluidframework/tool-utils";
import {
	OdspDocumentServiceFactory,
	createOdspCreateContainerRequest,
	createOdspUrl,
	OdspDriverUrlResolver,
} from "@fluidframework/odsp-driver";

const passwordTokenConfig = (username: string, password: string): OdspTokenConfig => ({
	type: "password",
	username,
	password,
});

// specific a range of user name from <prefix><start> to <prefix><start + count - 1> all having the same password
interface LoginTenantRange {
	prefix: string;
	start: number;
	count: number;
	password: string;
}

interface LoginTenants {
	[tenant: string]: {
		range: LoginTenantRange;
		// add different format here
	};
}

interface IOdspTestLoginInfo {
	siteUrl: string;
	username: string;
	password: string;
	supportsBrowserAuth?: boolean;
}

type OdspEndpoint = "odsp" | "odsp-df";

type TokenConfig = IOdspTestLoginInfo & IClientConfig;

export interface IOdspTestDriverConfig extends TokenConfig {
	directory: string;
	driveId: string;
	options: HostStoragePolicy | undefined;
}

export function assertOdspEndpoint(
	endpoint: string | undefined,
): asserts endpoint is OdspEndpoint | undefined {
	if (endpoint === undefined || endpoint === "odsp" || endpoint === "odsp-df") {
		return;
	}
	throw new TypeError("Not a odsp endpoint");
}

/**
 * Get from the env a set of credential to use from a single tenant
 * @param tenantIndex - interger to choose the tenant from an array
 * @param requestedUserName - specific user name to filter to
 */
function getCredentials(
	odspEndpointName: OdspEndpoint,
	tenantIndex: number,
	requestedUserName?: string,
) {
	const creds: { [user: string]: string } = {};
	const loginTenants =
		odspEndpointName === "odsp"
			? process.env.login__odsp__test__tenants
			: process.env.login__odspdf__test__tenants;
	if (loginTenants !== undefined) {
		const tenants: LoginTenants = JSON.parse(loginTenants);
		const tenantNames = Object.keys(tenants);
		const tenant = tenantNames[tenantIndex % tenantNames.length];
		const tenantInfo = tenants[tenant];
		// Translate all the user from that user to the full user principle name by appending the tenant domain
		const range = tenantInfo.range;

		// Return the set of account to choose from a single tenant
		for (let i = 0; i < range.count; i++) {
			const username = `${range.prefix}${range.start + i}@${tenant}`;
			if (requestedUserName === undefined || requestedUserName === username) {
				creds[username] = range.password;
			}
		}
	} else {
		const loginAccounts =
			odspEndpointName === "odsp"
				? process.env.login__odsp__test__accounts
				: process.env.login__odspdf__test__accounts;
		assert(loginAccounts !== undefined, "Missing login__odsp/odspdf__test__accounts");
		// Expected format of login__odsp__test__accounts is simply string key-value pairs of username and password
		const passwords: { [user: string]: string } = JSON.parse(loginAccounts);

		// Need to choose one out of the set as these account might be from different tenant
		const username = requestedUserName ?? Object.keys(passwords)[0];
		assert(passwords[username], `No password for username: ${username}`);
		creds[username] = passwords[username];
	}
	return creds;
}

export const OdspDriverApi = {
	version: "2.0.0-internal.3.2.0",
	OdspDocumentServiceFactory,
	OdspDriverUrlResolver,
	createOdspCreateContainerRequest,
	createOdspUrl, // REVIEW: does this need to be back compat?
};

export type OdspDriverApiType = typeof OdspDriverApi;

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
		const siteUrl = `${tokenConfig.siteUrl}`;
		console.log("siteUrl", siteUrl);
		return getDriveId(siteUrl, "", undefined, {
			accessToken: await this.getStorageToken({ siteUrl, refresh: false }, tokenConfig),
		});
	}
	public static async createFromEnv(
		config?: {
			directory?: string;
			username?: string;
			options?: HostStoragePolicy;
			supportsBrowserAuth?: boolean;
			tenantIndex?: number;
			odspEndpointName?: string;
		},
		api: OdspDriverApiType = OdspDriverApi,
	) {
		const tenantIndex = config?.tenantIndex ?? 0;
		assertOdspEndpoint(config?.odspEndpointName);
		const endpointName = config?.odspEndpointName ?? "odsp";
		const creds = getCredentials(endpointName, tenantIndex, config?.username);
		// Pick a random one on the list (only supported for >= 0.46)
		const users = Object.keys(creds);
		const randomUserIndex = Math.random();
		const userIndex = Math.floor(randomUserIndex * users.length);
		const username = users[userIndex];

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

		return this.create(
			{
				username,
				password: creds[username],
				siteUrl,
				supportsBrowserAuth: config?.supportsBrowserAuth,
			},
			config?.directory ?? "",
			api,
			options,
			tenantName,
			userIndex,
			endpointName,
		);
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

	private static async create(
		loginConfig: IOdspTestLoginInfo,
		directory: string,
		api = OdspDriverApi,
		options?: HostStoragePolicy,
		tenantName?: string,
		userIndex?: number,
		endpointName?: string,
	) {
		const tokenConfig: TokenConfig = {
			...loginConfig,
			...getMicrosoftConfiguration(),
		};

		const driveId = await this.getDriveId(loginConfig.siteUrl, tokenConfig);
		const directoryParts = [directory];

		const driverConfig: IOdspTestDriverConfig = {
			...tokenConfig,
			directory: directoryParts.join("/"),
			driveId,
			options,
		};

		console.log("API", api.version);

		return new OdspDriver(driverConfig, api, tenantName, userIndex, endpointName);
	}

	private static async getStorageToken(
		options: OdspResourceTokenFetchOptions & { useBrowserAuth?: boolean },
		config: IOdspTestLoginInfo & IClientConfig,
	) {
		// const host = new URL(options.siteUrl).host;
		// console.log("site url host-------", host);
		// if (options.useBrowserAuth === true) {
		// 	console.log("site url 1-------");
		// 	const browserTokens = await this.odspTokenManager.getOdspTokens(
		// 		host,
		// 		config,
		// 		{
		// 			type: "browserLogin",
		// 			navigator: (openUrl) => {
		// 				console.log(
		// 					`Open the following url in a new private browser window, and login with user: ${config.username}`,
		// 				);
		// 				console.log(
		// 					`Additional account details may be available in the environment variable login__odsp__test__accounts`,
		// 				);
		// 				console.log(`"${openUrl}"`);
		// 			},
		// 		},
		// 		options.refresh,
		// 	);
		// 	return browserTokens.accessToken;
		// }
		// console.log("site url 2-------", passwordTokenConfig(config.username, config.password));
		// // This function can handle token request for any multiple sites.
		// // Where the test driver is for a specific site.
		// const tokens = await this.odspTokenManager.getOdspTokens(
		// 	host,
		// 	config,
		// 	passwordTokenConfig(config.username, config.password),
		// 	options.refresh,
		// );
		// console.log("site url 2-------", passwordTokenConfig(config.username, config.password));
		// return tokens.accessToken;
		return "STORAGE_TOKEN";
	}

	public get siteUrl(): string {
		return `${this.config.siteUrl}`;
	}
	public get driveId(): string {
		return this.config.driveId;
	}
	public get directory(): string {
		return this.config.directory;
	}

	private constructor(
		private readonly config: Readonly<IOdspTestDriverConfig>,
		private readonly api = OdspDriverApi,
		public readonly tenantName?: string,
		public readonly userIndex?: number,
		public readonly endpointName?: string,
	) {
		console.log(this.api.version);
	}

	public readonly getStorageToken = async (options: OdspResourceTokenFetchOptions) => {
		return OdspDriver.getStorageToken(options, this.config);
	};
	public readonly getPushToken = async (options: OdspResourceTokenFetchOptions) => {
		// const tokens = await OdspDriver.odspTokenManager.getPushTokens(
		// 	new URL(options.siteUrl).hostname,
		// 	this.config,
		// 	passwordTokenConfig(this.config.username, this.config.password),
		// 	options.refresh,
		// );

		// return tokens.accessToken;
		return "PUSH_TOKEN";
	};
}
