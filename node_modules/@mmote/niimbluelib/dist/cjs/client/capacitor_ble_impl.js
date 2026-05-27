"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.NiimbotCapacitorBleClient = void 0;
const events_1 = require("../events");
const _1 = require(".");
const packets_1 = require("../packets");
const utils_1 = require("../utils");
const bluetooth_le_1 = require("@capacitor-community/bluetooth-le");
/**
 * Uses [@capacitor-community/bluetooth-le](https://github.com/capacitor-community/bluetooth-le)
 *
 * @category Client
 */
class NiimbotCapacitorBleClient extends _1.NiimbotAbstractClient {
    async connect(options) {
        await this.disconnect();
        await bluetooth_le_1.BleClient.initialize({ androidNeverForLocation: true });
        const bluetoothEnabled = await bluetooth_le_1.BleClient.isEnabled();
        if (!bluetoothEnabled) {
            throw new Error("Bluetooth is not enabled");
        }
        let device;
        if (options?.deviceId !== undefined) {
            device = {
                deviceId: options.deviceId,
                name: options.deviceId,
            };
        }
        else {
            device = await bluetooth_le_1.BleClient.requestDevice();
        }
        await bluetooth_le_1.BleClient.connect(device.deviceId, () => this.onBleDisconnect());
        await bluetooth_le_1.BleClient.discoverServices(device.deviceId);
        const { service, characteristic } = await this.findSuitableCharacteristic(device.deviceId).finally(() => this.onBleDisconnect());
        this.deviceId = device.deviceId;
        this.serviceUUID = service;
        this.characteristicUUID = characteristic;
        if (this.debug) {
            console.log("Suitable channel found:", { service, characteristic });
        }
        await bluetooth_le_1.BleClient.startNotifications(this.deviceId, this.serviceUUID, this.characteristicUUID, (value) => {
            this.processRawPacket(value);
        });
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
    async findSuitableCharacteristic(devId) {
        const services = await bluetooth_le_1.BleClient.getServices(devId);
        for (const service of services) {
            if (service.uuid.length < 5) {
                continue;
            }
            const characteristics = service.characteristics;
            for (const ch of characteristics) {
                if (ch.properties.notify && ch.properties.writeWithoutResponse) {
                    return {
                        characteristic: ch.uuid,
                        service: service.uuid,
                    };
                }
            }
        }
        throw new Error("Unable to find suitable channel characteristic");
    }
    onBleDisconnect() {
        this.deviceId = undefined;
        this.serviceUUID = undefined;
        this.characteristicUUID = undefined;
        this.info = {};
        this.emit("disconnect", new events_1.DisconnectEvent());
    }
    isConnected() {
        return this.deviceId !== undefined;
    }
    async disconnect() {
        this.stopHeartbeat();
        if (this.deviceId !== undefined) {
            await bluetooth_le_1.BleClient.stopNotifications(this.deviceId, this.serviceUUID, this.characteristicUUID);
            await bluetooth_le_1.BleClient.disconnect(this.deviceId);
        }
        this.deviceId = undefined;
        this.info = {};
    }
    async sendRaw(data, force) {
        const send = async () => {
            if (!this.isConnected()) {
                throw new Error("Channel is closed");
            }
            await utils_1.Utils.sleep(this.packetIntervalMs);
            const dw = new DataView(data.buffer, data.byteOffset, data.byteLength);
            await bluetooth_le_1.BleClient.writeWithoutResponse(this.deviceId, this.serviceUUID, this.characteristicUUID, dw);
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
exports.NiimbotCapacitorBleClient = NiimbotCapacitorBleClient;
