"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.NiimbotSerialClient = void 0;
const events_1 = require("../events");
const _1 = require(".");
const packets_1 = require("../packets");
const utils_1 = require("../utils");
/**
 * Uses [Web Serial API](https://developer.mozilla.org/en-US/docs/Web/API/Web_Serial_API)
 *
 * @category Client
 **/
class NiimbotSerialClient extends _1.NiimbotAbstractClient {
    constructor() {
        super(...arguments);
        this.port = undefined;
        this.writer = undefined;
        this.reader = undefined;
    }
    async connect() {
        await this.disconnect();
        const _port = await navigator.serial.requestPort();
        _port.addEventListener("disconnect", () => {
            this.port = undefined;
            this.emit("disconnect", new events_1.DisconnectEvent());
        });
        await _port.open({ baudRate: 115200 });
        if (_port.readable === null) {
            throw new Error("Port is not readable");
        }
        if (_port.writable === null) {
            throw new Error("Port is not writable");
        }
        this.port = _port;
        const info = _port.getInfo();
        this.writer = _port.writable.getWriter();
        this.reader = _port.readable.getReader();
        setTimeout(() => {
            void (async () => {
                await this.waitSerialData();
            })();
        }, 1); // todo: maybe some other way exists
        try {
            await this.initialNegotiate();
            await this.fetchPrinterInfo();
        }
        catch (e) {
            console.error("Unable to fetch printer info (is it turned on?).");
            console.error(e);
        }
        const result = {
            deviceName: `Serial (VID:${info.usbVendorId?.toString(16)} PID:${info.usbProductId?.toString(16)})`,
            result: this.info.connectResult ?? packets_1.ConnectResult.FirmwareErrors,
        };
        this.emit("connect", new events_1.ConnectEvent(result));
        return result;
    }
    async waitSerialData() {
        while (true) {
            try {
                const result = await this.reader.read();
                if (result.value) {
                    if (this.debug) {
                        console.info(`<< serial chunk ${utils_1.Utils.bufToHex(result.value)}`);
                    }
                    this.processRawPacket(result.value);
                }
                if (result.done) {
                    console.log("done");
                    break;
                }
            }
            catch (_e) {
                break;
            }
        }
    }
    async disconnect() {
        this.stopHeartbeat();
        if (this.writer !== undefined) {
            this.writer.releaseLock();
        }
        if (this.reader !== undefined) {
            this.reader.releaseLock();
        }
        if (this.port !== undefined) {
            await this.port.close();
            this.emit("disconnect", new events_1.DisconnectEvent());
        }
        this.port = undefined;
        this.writer = undefined;
    }
    isConnected() {
        return this.port !== undefined && this.writer !== undefined && this.reader !== undefined;
    }
    async sendRaw(data, force) {
        const send = async () => {
            if (!this.isConnected()) {
                throw new Error("Port is not readable/writable");
            }
            await utils_1.Utils.sleep(this.packetIntervalMs);
            await this.writer.write(data);
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
exports.NiimbotSerialClient = NiimbotSerialClient;
