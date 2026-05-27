/** @category Image encoder */
export type ImageRow = {
    dataType: "void" | "pixels" | "check";
    rowNumber: number;
    repeat: number;
    blackPixelsCount: number;
    rowData?: Uint8Array;
};
/** @category Image encoder */
export type EncodedImage = {
    cols: number;
    rows: number;
    rowsData: ImageRow[];
};
/** @category Image encoder */
export type PrintDirection = "left" | "top";
/**
 * @category Helpers
 * @category Image encoder
 */
export declare class ImageEncoder {
    /** printDirection = "left" rotates image for 90 degrees clockwise */
    static encodeCanvas(canvas: HTMLCanvasElement, printDirection?: PrintDirection): EncodedImage;
    /** printDirection = "left" rotates image to 90 degrees clockwise */
    static isPixelNonWhite(iData: ImageData, x: number, y: number, printDirection?: PrintDirection): boolean;
    /**
     * @param data Pixels encoded by {@link encodeCanvas} (byte is 8 pixels)
     * @returns Array of indexes where every index stored in two bytes (big endian)
     */
    static indexPixels(data: Uint8Array): Uint8Array;
}
