import { ConnectionInfo, NiimbotAbstractClient } from ".";
/**
 * @category Client
 */
export declare class BleDefaultConfiguration {
    static readonly SERVICES: string[];
    static readonly NAME_FILTERS: BluetoothLEScanFilter[];
}
/**
 * Uses [Web Bluetooth API](https://developer.mozilla.org/en-US/docs/Web/API/Web_Bluetooth_API)
 *
 * @category Client
 */
export declare class NiimbotBluetoothClient extends NiimbotAbstractClient {
    private gattServer?;
    private channel?;
    private serviceUuidFilter;
    getServiceUuidFilter(): string[];
    setServiceUuidFilter(ids: string[]): void;
    connect(): Promise<ConnectionInfo>;
    private findSuitableBluetoothCharacteristic;
    isConnected(): boolean;
    disconnect(): Promise<void>;
    sendRaw(data: Uint8Array, force?: boolean): Promise<void>;
}
