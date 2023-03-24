/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { v4 as uuid } from "uuid";
import { SharedMap } from "@fluidframework/map";
import { ContainerSchema } from "@fluidframework/fluid-static";
import { OdspCreateContainerConfig, OdspGetContainerConfig } from "./odsp-client/interfaces";
import { OdspClient } from "./odsp-client/OdspClient";
import { getodspDriver } from "./odsp-client/InitiateDriver";
import { OdspDriver } from "./odsp-client";

export const diceValueKey = "dice-value-key";
// let window: { [key: string]: any } = {};
let sharingLink: string;
const documentId = uuid();

const containerSchema: ContainerSchema = {
	initialObjects: { diceMap: SharedMap },
};

const root = document.getElementById("content");

const createDice = async (odspDriver: OdspDriver) => {
	console.log("CREATING THE CONTAINER");
	const containerConfig: OdspCreateContainerConfig = {
		siteUrl: odspDriver.siteUrl,
		driveId: odspDriver.driveId,
		folderName: odspDriver.directory,
		fileName: documentId,
	};

	console.log("CONTAINER CONFIG", containerConfig);

	const { fluidContainer, containerServices } = await OdspClient.createContainer(
		containerConfig,
		containerSchema,
	);

	console.log("CONTAINER CREATED");

	sharingLink = await containerServices.generateLink();

	console.log("SHARING LINK", sharingLink);

	const map = fluidContainer.initialObjects.diceMap as SharedMap;
	map.set(diceValueKey, 1);
	renderDiceRoller(map, root);
	return sharingLink;
};

const loadDice = async (url: string) => {
	const containerConfig: OdspGetContainerConfig = {
		fileUrl: url, //pass file url
	};

	const { fluidContainer } = await OdspClient.getContainer(containerConfig, containerSchema);

	const map = fluidContainer.initialObjects.diceMap as SharedMap;
	renderDiceRoller(map, root);
};

async function start() {
	console.log("Initiating the driver------");
	const odspDriver = await getodspDriver();
	console.log("INITIAL DRIVER", odspDriver);

	if (location.hash) {
		// const containerId = /^#(https?|ftp):\/\//i.test(location.hash)
		// 	? decodeURI(location.hash)
		// 	: location.hash;
		const hash = decodeURI(location.hash);
		const id = hash.charAt(0) === "#" ? hash.substring(1) : hash;
		console.log("hash---", id);
		await loadDice(id);
	} else {
		const containerId = await createDice(odspDriver);
		console.log("container id: ", containerId);
		// const itemId = getItemIdFromUrl(containerId);
		/**
		 * The encodeURI() function is used to encode a URI and the decodeURI() function is used to decode the encoded URI.
		 * In this code, url is encoded using encodeURI() and then stored in the window.location.hash property. Finally, the location.hash
		 * property is decoded using decodeURI() and logged to the console.
		 */
		window.location.hash = encodeURI(containerId); //The encodeURI function is used to encode the URL string so that it can be safely used in the hash fragment of the URL. The location.hash property returns the value of the hash fragment of the URL. In this case, it returns the encoded URL string as the hash fragment.
		// window.location.hash = itemId !== undefined ? itemId : encodeURI(containerId);
	}
}

start().catch((error) => console.error(error));

// Define the view
const template = document.createElement("template");

template.innerHTML = `
  <style>
    .wrapper { text-align: center }
    .dice { font-size: 200px }
    .roll { font-size: 50px;}
  </style>
  <div class="wrapper">
    <div class="dice"></div>
    <button class="roll"> Roll </button>
  </div>
`;

const renderDiceRoller = (diceMap: SharedMap, elem: any) => {
	elem.appendChild(template.content.cloneNode(true));

	const rollButton = elem.querySelector(".roll");
	const dice = elem.querySelector(".dice");

	// Set the value at our dataKey with a random number between 1 and 6.
	rollButton.onclick = () => diceMap.set(diceValueKey, Math.floor(Math.random() * 6) + 1);

	// Get the current value of the shared data to update the view whenever it changes.
	const updateDice = () => {
		const diceValue = diceMap.get(diceValueKey);
		// Unicode 0x2680-0x2685 are the sides of a dice (⚀⚁⚂⚃⚄⚅)
		dice.textContent = String.fromCodePoint(0x267f + diceValue);
		dice.style.color = `hsl(${diceValue * 60}, 70%, 30%)`;
	};
	updateDice();

	// Use the changed event to trigger the rerender whenever the value changes.
	diceMap.on("valueChanged", updateDice);

	// Setting "fluidStarted" is just for our test automation
	// window["fluidStarted"] = true;
};
