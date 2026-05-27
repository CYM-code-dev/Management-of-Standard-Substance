"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ImageEncoder = void 0;
const _1 = require(".");
/**
 * @category Helpers
 * @category Image encoder
 */
class ImageEncoder {
    /** printDirection = "left" rotates image for 90 degrees clockwise */
    static encodeCanvas(canvas, printDirection = "left") {
        const ctx = canvas.getContext("2d");
        const iData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const rowsData = [];
        let cols = canvas.width;
        let rows = canvas.height;
        if (printDirection === "left") {
            cols = canvas.height;
            rows = canvas.width;
        }
        if (cols % 8 !== 0) {
            throw new Error("Column count must be multiple of 8");
        }
        for (let row = 0; row < rows; row++) {
            let isVoid = true;
            let blackPixelsCount = 0;
            const rowData = new Uint8Array(cols / 8);
            for (let colOct = 0; colOct < cols / 8; colOct++) {
                let pixelsOctet = 0;
                for (let colBit = 0; colBit < 8; colBit++) {
                    if (ImageEncoder.isPixelNonWhite(iData, colOct * 8 + colBit, row, printDirection)) {
                        pixelsOctet |= 1 << (7 - colBit);
                        isVoid = false;
                        blackPixelsCount++;
                    }
                }
                rowData[colOct] = pixelsOctet;
            }
            const newPart = {
                dataType: isVoid ? "void" : "pixels",
                rowNumber: row,
                repeat: 1,
                rowData: isVoid ? undefined : rowData,
                blackPixelsCount,
            };
            // Check previous row and increment repeats instead of adding new row if data is same
            if (rowsData.length === 0) {
                rowsData.push(newPart);
            }
            else {
                const lastPacket = rowsData[rowsData.length - 1];
                let same = newPart.dataType === lastPacket.dataType;
                if (same && newPart.dataType === "pixels") {
                    same = _1.Utils.u8ArraysEqual(newPart.rowData, lastPacket.rowData);
                }
                if (same) {
                    lastPacket.repeat++;
                }
                else {
                    rowsData.push(newPart);
                }
                const sendRowCheck = row % 200 === 199;
                if (sendRowCheck) {
                    rowsData.push({
                        dataType: "check",
                        rowNumber: row,
                        repeat: 0,
                        rowData: undefined,
                        blackPixelsCount: 0,
                    });
                }
            }
        }
        return { cols, rows, rowsData };
    }
    /** printDirection = "left" rotates image to 90 degrees clockwise */
    static isPixelNonWhite(iData, x, y, printDirection = "left") {
        let idx = y * iData.width + x;
        if (printDirection === "left") {
            idx = (iData.height - 1 - x) * iData.width + y;
        }
        idx *= 4;
        return iData.data[idx] !== 255 || iData.data[idx + 1] !== 255 || iData.data[idx + 2] !== 255;
    }
    /**
     * @param data Pixels encoded by {@link encodeCanvas} (byte is 8 pixels)
     * @returns Array of indexes where every index stored in two bytes (big endian)
     */
    static indexPixels(data) {
        const result = [];
        for (let bytePos = 0; bytePos < data.byteLength; bytePos++) {
            const b = data[bytePos];
            for (let bitPos = 0; bitPos < 8; bitPos++) {
                // iterate from most significant bit of byte
                if (b & (1 << (7 - bitPos))) {
                    result.push(..._1.Utils.u16ToBytes(bytePos * 8 + bitPos));
                }
            }
        }
        return new Uint8Array(result);
    }
}
exports.ImageEncoder = ImageEncoder;
