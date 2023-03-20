/* eslint-disable max-len */
/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { getDriveId, IClientConfig } from "@fluidframework/odsp-doclib-utils";
import type {
	OdspResourceTokenFetchOptions,
	HostStoragePolicy,
} from "@fluidframework/odsp-driver-definitions";
import { OdspTokenConfig, getMicrosoftConfiguration } from "@fluidframework/tool-utils";
import {
	OdspDocumentServiceFactory,
	createOdspCreateContainerRequest,
	createOdspUrl,
	OdspDriverUrlResolver,
} from "@fluidframework/odsp-driver";

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
	private static readonly driveIdPCache = new Map<string, Promise<string>>();

	private static async getDriveIdFromConfig(tokenConfig: TokenConfig): Promise<string> {
		const siteUrl = `${tokenConfig.siteUrl}`;
		console.log("siteUrl", siteUrl);
		return getDriveId(siteUrl, "", undefined, {
			accessToken: (await this.getStorageToken(
				{ siteUrl, refresh: false },
				tokenConfig,
			)) as any,
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
		const options = config?.options ?? {};
		options.isolateSocketCache = true;

		return this.create(
			{
				username: "username",
				password: "password",
				siteUrl: "siteUrl",
				supportsBrowserAuth: config?.supportsBrowserAuth,
			},
			config?.directory ?? "",
			api,
			options,
			"tenant_name",
			"odsp",
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
		endpointName?: string,
	) {
		const tokenConfig: TokenConfig = {
			...loginConfig,
			...getMicrosoftConfiguration(),
		};

		const driveId = await this.getDriveId(loginConfig.siteUrl, tokenConfig);
		console.log("Drive id: ", driveId);
		const directoryParts = [directory];

		const driverConfig: IOdspTestDriverConfig = {
			...tokenConfig,
			directory: directoryParts.join("/"),
			driveId,
			options,
		};

		console.log("API", api.version);

		return new OdspDriver(driverConfig, api, tenantName, endpointName);
	}

	private static async getGraphToken(
		options: OdspResourceTokenFetchOptions & { useBrowserAuth?: boolean },
		config: IOdspTestLoginInfo & IClientConfig,
	) {
		return "GRAPH_TOKEN";
	}

	private static async getStorageToken(
		options: OdspResourceTokenFetchOptions & { useBrowserAuth?: boolean },
		config: IOdspTestLoginInfo & IClientConfig,
	) {
		return "STORAGE_TOKEN";
	}

	private static async getPushToken(
		options: OdspResourceTokenFetchOptions & { useBrowserAuth?: boolean },
		config: IOdspTestLoginInfo & IClientConfig,
	) {
		return "PUSH_TOKEN";
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
		public readonly userIndex?: string,
		public readonly endpointName?: string,
	) {
		console.log(this.api.version);
	}

	public readonly getStorageToken = async (options: OdspResourceTokenFetchOptions) => {
		return OdspDriver.getStorageToken(options, this.config);
	};
	public readonly getPushToken = async (options: OdspResourceTokenFetchOptions) => {
		return OdspDriver.getPushToken(options, this.config);
	};
	public readonly getGraphToken = async (options: OdspResourceTokenFetchOptions) => {
		return OdspDriver.getGraphToken(options, this.config);
	};
	public readonly getMicrosoftGraphToken = "GRAPH_TOKEN";
}
