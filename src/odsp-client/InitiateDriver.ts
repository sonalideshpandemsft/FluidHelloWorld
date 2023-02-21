/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { OdspConnectionConfig } from "./interfaces";
import { OdspClient } from "./OdspClient";
import { OdspDriver } from "./OdspDriver";

export const initDriver = async () => {
	const driver: OdspDriver = await OdspDriver.createFromEnv();
	const connectionConfig: OdspConnectionConfig = {
		getSharePointToken: driver.getStorageToken,
		getPushServiceToken: driver.getPushToken,
	};

	OdspClient.init(connectionConfig);
	return driver;
};
