import { NiimbotAbstractClient, ConnectionInfo } from "./abstract_client";
import { NiimbotBluetoothClient } from "./bluetooth_impl";
import { NiimbotCapacitorBleClient, NiimbotCapacitorBleClientConnectOptions } from "./capacitor_ble_impl";
import { NiimbotSerialClient } from "./serial_impl";
/** Client type for {@link instantiateClient} */
export type NiimbotClientType = "bluetooth" | "serial" | "capacitor-ble";
/** Create new client instance */
export declare const instantiateClient: (t: NiimbotClientType) => NiimbotAbstractClient;
export { NiimbotAbstractClient, ConnectionInfo, NiimbotBluetoothClient, NiimbotCapacitorBleClient, NiimbotCapacitorBleClientConnectOptions, NiimbotSerialClient, };
