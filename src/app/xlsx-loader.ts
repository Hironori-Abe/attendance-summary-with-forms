type XlsxModule = typeof import('xlsx');
type XlsxCalcModule = typeof import('xlsx-calc');

let xlsxModulePromise: Promise<XlsxModule> | undefined;
let xlsxCalcModulePromise: Promise<XlsxCalcModule> | undefined;

export function loadXlsx(): Promise<XlsxModule> {
  xlsxModulePromise ??= import('xlsx');
  return xlsxModulePromise;
}

export function loadXlsxCalc(): Promise<XlsxCalcModule> {
  xlsxCalcModulePromise ??= import('xlsx-calc');
  return xlsxCalcModulePromise;
}
