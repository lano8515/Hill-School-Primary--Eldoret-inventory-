declare module 'jsbarcode' {
  type BarcodeOptions = {
    format?: string;
    displayValue?: boolean;
    margin?: number;
    height?: number;
    width?: number;
  };
  function JsBarcode(element: SVGSVGElement, value: string, options?: BarcodeOptions): void;
  export default JsBarcode;
}
