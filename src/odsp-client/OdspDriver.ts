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
    OdspTokenConfig,
    OdspTokenManager,
    odspTokensCache,
    getMicrosoftConfiguration,
} from "@fluidframework/tool-utils";

const passwordTokenConfig = (username: string, password: string): OdspTokenConfig => ({
    type: "password",
    username,
    password,
});

interface IOdspTestLoginInfo {
    server: string;
    username: string;
    password: string;
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
    private static readonly driverIdPCache = new Map<string, Promise<string>>();
    private static async getDriveId(server: string, tokenConfig: TokenConfig): Promise<string> {
        const siteUrl = `https://${tokenConfig.server}`;
        return getDriveId(server, "", undefined, {
            accessToken: await this.getStorageToken({ siteUrl, refresh: false }, tokenConfig),
        });
    }

    public static async createFromEnv(config?: {
        directory?: string;
        username?: string;
        options?: HostStoragePolicy;
    }) {
        const loginAccounts = process.env.login__odsp__test__accounts;
        assert(loginAccounts !== undefined, "Missing login__odsp__test__accounts");
        // Expected format of login__odsp__test__accounts is simply string key-value pairs of username and password
        const passwords: { [user: string]: string } = JSON.parse(loginAccounts);
        const username = config?.username ?? Object.keys(passwords)[0];
        assert(passwords[username], `No password for username: ${username}`);

        const emailServer = username.substr(username.indexOf("@") + 1);
        const server = `${emailServer.substr(0, emailServer.indexOf("."))}.sharepoint.com`;

        return this.create(
            {
                username,
                password: passwords[username],
                server,
            },
            config?.directory ?? "",
            config?.options,
        );
    }

    private static async create(
        loginConfig: IOdspTestLoginInfo,
        directory: string,
        options?: HostStoragePolicy,
    ) {
        const tokenConfig: TokenConfig = {
            ...loginConfig,
            ...getMicrosoftConfiguration(),
        };

        let driveIdP = this.driverIdPCache.get(loginConfig.server);
        if (!driveIdP) {
            driveIdP = this.getDriveId(loginConfig.server, tokenConfig);
        }

        const driveId = await driveIdP;
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

        return new OdspDriver(driverConfig);
    }

    private static async getStorageToken(
        options: OdspResourceTokenFetchOptions,
        config: IOdspTestLoginInfo & IClientConfig,
    ) {
        // This function can handle token request for any multiple sites. Where the test driver is for a specific site.
        const tokens = await this.odspTokenManager.getOdspTokens(
            new URL(options.siteUrl).hostname,
            config,
            passwordTokenConfig(config.username, config.password),
            options.refresh,
        );
        return tokens.accessToken;
    }

    public get siteUrl(): string {
        return `https://${this.config.server}`;
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
