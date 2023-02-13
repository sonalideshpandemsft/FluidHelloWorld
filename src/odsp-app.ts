/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { v4 as uuid } from "uuid";
import { SharedMap } from "@fluidframework/map";
import { ContainerSchema } from "@fluidframework/fluid-static";
import { OdspDriver } from "./odsp-client/OdspDriver";
import {
    OdspConnectionConfig,
    OdspCreateContainerConfig,
    OdspGetContainerConfig,
} from "./odsp-client/interfaces";
import { OdspClient } from "./odsp-client/OdspClient";

export const diceValueKey = "dice-value-key";
let window: { [key: string]: any };
let driver: OdspDriver;
const documentId = uuid();

const initDriver = async () => {
    driver = await OdspDriver.createFromEnv({ directory: "OdspTestClient" });
    const connectionConfig: OdspConnectionConfig = {
        getSharePointToken: driver.getStorageToken,
        getPushServiceToken: driver.getPushToken,
    };

    OdspClient.init(connectionConfig);
    return driver;
};

const containerSchema: ContainerSchema = {
    initialObjects: { diceMap: SharedMap },
};

const root = document.getElementById("content");

const createNewDice = async (odspDriver: OdspDriver) => {
    const containerConfig: OdspCreateContainerConfig = {
        siteUrl: odspDriver.siteUrl,
        driveId: odspDriver.driveId,
        folderName: odspDriver.directory,
        fileName: documentId,
    };

    const { fluidContainer } = await OdspClient.createContainer(containerConfig, containerSchema);

    const map = fluidContainer.initialObjects.diceMap as SharedMap;
    map.set(diceValueKey, 1);
    const id = await fluidContainer.attach();
    renderDiceRoller(map, root);
    return id;
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
    const odspDriver = await initDriver();

    if (location.hash) {
        await loadDice(location.hash.substring(1));
    } else {
        const id = await createNewDice(odspDriver);
        location.hash = id;
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
    window["fluidStarted"] = true;
};
