/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { SharedMap } from "@fluidframework/map";
import { AzureClient, AzureLocalConnectionConfig } from "@fluidframework/azure-client";
import { ContainerSchema } from "@fluidframework/fluid-static";
import { InsecureTokenProvider } from "@fluidframework/test-client-utils";

export const diceValueKey = "dice-value-key";
let window: { [key: string]: any };

const userConfig = {
    id: "userId",
    name: "userName",
    additionalDetails: {
        email: "userName@example.com",
    },
};

const serviceConfig: AzureLocalConnectionConfig = {
    type: "local",
    tokenProvider: new InsecureTokenProvider("fooBar", userConfig),
    endpoint: "http://localhost:7070",
};

const client = new AzureClient({ connection: serviceConfig });

const containerSchema: ContainerSchema = {
    initialObjects: { diceMap: SharedMap },
};

const root = document.getElementById("content");

const createNewDice = async () => {
    const { container } = await client.createContainer(containerSchema);
    const map = container.initialObjects.diceMap as SharedMap;
    map.set(diceValueKey, 1);
    const id = await container.attach();
    renderDiceRoller(map, root);
    return id;
};

const loadExistingDice = async (id: string) => {
    const { container } = await client.getContainer(id, containerSchema);
    const map = container.initialObjects.diceMap as SharedMap;
    renderDiceRoller(map, root);
};

async function start() {
    if (location.hash) {
        await loadExistingDice(location.hash.substring(1));
    } else {
        const id = await createNewDice();
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

const renderDiceRoller = (diceMap: any, elem: any) => {
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
