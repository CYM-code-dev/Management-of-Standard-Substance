import { ConnectionInfo, NiimbotAbstractClient } from ".";
/**
 * Uses [Web Serial API](https://developer.mozilla.org/en-US/docs/Web/API/Web_Serial_API)
 *
 * @category Client
 **/
export declare class NiimbotSerialClient extends NiimbotAbstractClient {
    private port?;
    private writer?;
    private reader?;
    connect(): Promise<ConnectionInfo>;
    private waitSerialData;
    disconnect(): Promise<void>;
    isConnected(): boolean;
    sendRaw(data: Uint8Array, force?: boolean): Promise<void>;
}
