"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.NiimbotBluetoothClient = exports.BleDefaultConfiguration = void 0;
const events_1 = require("../events");
const _1 = require(".");
const packets_1 = require("../packets");
const utils_1 = require("../utils");
const printer_models_1 = require("../printer_models");
const getAllModelFirstLetters = () => [...new Set(printer_models_1.modelsLibrary.map((m) => m.model[0]))];
/**
 * @category Client
 */
class BleDefaultConfiguration {
}
exports.BleDefaultConfiguration = BleDefaultConfiguration;
BleDefaultConfiguration.SERVICES = ["e7810a71-73ae-499d-8c15-faa9aef0c3f2"];
BleDefaultConfiguration.NAME_FILTERS = [
    ...getAllModelFirstLetters().map((l) => ({ namePrefix: l })),
];
/**
 * Uses [Web Bluetooth API](https://developer.mozilla.org/en-US/docs/Web/API/Web_Bluetooth_API)
 *
 * @category Client
 */
class NiimbotBluetoothClient extends _1.NiimbotAbstractClient {
    constructor() {
        super(...arguments);
        this.gattServer = undefined;
        this.channel = undefined;
        this.serviceUuidFilter = BleDefaultConfiguration.SERVICES;
    }
    getServiceUuidFilter() {
        return this.serviceUuidFilter;
    }
    setServiceUuidFilter(ids) {
        this.serviceUuidFilter = ids;
    }
    async connect() {
        await this.disconnect();
        const options = {
            filters: [
                ...BleDefaultConfiguration.NAME_FILTERS,
                { services: this.serviceUuidFilter ?? BleDefaultConfiguration.SERVICES },
            ],
        };
        const device = await navigator.bluetooth.requestDevice(options);
        if (device.gatt === undefined) {
            throw new Error("Device has no Bluetooth Generic Attribute Profile");
        }
        const disconnectListener = () => {
            this.gattServer = undefined;
            this.channel = undefined;
            this.info = {};
            this.emit("disconnect", new events_1.DisconnectEvent());
            device.removeEventListener("gattserverdisconnected", disconnectListener);
        };
        device.addEventListener("gattserverdisconnected", disconnectListener);
        const gattServer = await device.gatt.connect();
        const channel = await this.findSuitableBluetoothCharacteristic(gattServer);
        if (channel === undefined) {
            gattServer.disconnect();
            throw new Error("Suitable device characteristic not found");
        }
        console.log(`Found suitable characteristic ${channel.uuid}`);
        channel.addEventListener("characteristicvaluechanged", (event) => {
            const target = event.target;
            this.processRawPacket(target.value);
        });
        await channel.startNotifications();
        this.gattServer = gattServer;
        this.channel = channel;
        try {
            await this.initialNegotiate();
            await this.fetchPrinterInfo();
        }
        catch (e) {
            console.error("Unable to fetch printer info.");
            console.error(e);
        }
        const result = {
            deviceName: device.name,
            result: this.info.connectResult ?? packets_1.ConnectResult.FirmwareErrors,
        };
        this.emit("connect", new events_1.ConnectEvent(result));
        return result;
    }
    async findSuitableBluetoothCharacteristic(gattServer) {
        const services = await gattServer.getPrimaryServices();
        for (const service of services) {
            if (service.uuid.length < 5) {
                continue;
            }
            const characteristics = await service.getCharacteristics();
            for (const c of characteristics) {
                if (c.properties.notify && c.properties.writeWithoutResponse) {
                    return c;
                }
            }
        }
        return undefined;
    }
    isConnected() {
        return this.gattServer !== undefined && this.channel !== undefined;
    }
    // eslint-disable-next-line @typescript-eslint/require-await
    async disconnect() {
        this.stopHeartbeat();
        this.gattServer?.disconnect();
        this.gattServer = undefined;
        this.channel = undefined;
        this.info = {};
    }
    async sendRaw(data, force) {
        const send = async () => {
            if (this.channel === undefined) {
                throw new Error("Channel is closed");
            }
            await utils_1.Utils.sleep(this.packetIntervalMs);
            await this.channel.writeValueWithoutResponse(data.buffer);
            this.emit("rawpacketsent", new events_1.RawPacketSentEvent(data));
        };
        if (force) {
            await send();
        }
        else {
            await this.mutex.runExclusive(send);
        }
    }
}
exports.NiimbotBluetoothClient = NiimbotBluetoothClient;
