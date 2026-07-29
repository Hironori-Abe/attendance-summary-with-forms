declare module 'xlsx-calc' {
  import type { WorkBook } from 'xlsx';

  interface CalculationOptions {
    continue_after_error?: boolean;
    log_error?: boolean;
  }

  function calculateWorkbook(
    workbook: WorkBook,
    options?: CalculationOptions,
  ): void;

  export default calculateWorkbook;
}
