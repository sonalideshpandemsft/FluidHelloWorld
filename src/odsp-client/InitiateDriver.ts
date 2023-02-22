/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { OdspConnectionConfig } from "./interfaces";
import { OdspClient } from "./OdspClient";
import { OdspDriver } from "./OdspDriver";

const initDriver = async () => {
	console.log("Driver init------");
	const driver: OdspDriver = await OdspDriver.createFromEnv({
		directory: "OdspFluidHelloWorld",
		supportsBrowserAuth: true,
	});
	console.log("Driver------", driver);
	const connectionConfig: OdspConnectionConfig = {
		getSharePointToken: driver.getStorageToken,
		getPushServiceToken: driver.getPushToken,
	};

	OdspClient.init(connectionConfig);
	return driver;
};

export const getodspDriver = async () => {
	const odspDriver = await initDriver();
	console.log("INITIAL DRIVER", odspDriver);
	return odspDriver;
};
