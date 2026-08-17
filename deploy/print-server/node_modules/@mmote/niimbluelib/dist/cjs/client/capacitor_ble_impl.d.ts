import { ConnectionInfo, NiimbotAbstractClient } from ".";
/**
 * @category Client
 */
export interface NiimbotCapacitorBleClientConnectOptions {
    /**
     * Skip device picker dialog and connect to given device ID.
     *
     * On **Android** this is the BLE MAC address.
     *
     * On **iOS** and **web** it is an identifier.
     */
    deviceId?: string;
}
/**
 * Uses [@capacitor-community/bluetooth-le](https://github.com/capacitor-community/bluetooth-le)
 *
 * @category Client
 */
export declare class NiimbotCapacitorBleClient extends NiimbotAbstractClient {
    private deviceId?;
    private serviceUUID?;
    private characteristicUUID?;
    connect(options?: NiimbotCapacitorBleClientConnectOptions): Promise<ConnectionInfo>;
    private findSuitableCharacteristic;
    private onBleDisconnect;
    isConnected(): boolean;
    disconnect(): Promise<void>;
    sendRaw(data: Uint8Array, force?: boolean): Promise<void>;
}
