import { Component, Inject } from '@angular/core';
import { MAT_DIALOG_DATA, MatDialogModule } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';

export interface FormulaCalculationWarningDialogData {
  files: Array<{
    fileName: string;
    failureCount: number;
  }>;
}

@Component({
  selector: 'app-formula-calculation-warning-dialog',
  standalone: true,
  imports: [MatDialogModule, MatButtonModule],
  template: `
    <h2 mat-dialog-title>数式を計算できませんでした</h2>
    <mat-dialog-content>
      <p>
        一部の「合計点数」列の式が複雑、または未対応の形式のため、変換できませんでした。
        該当する得点は、出力ファイル内では0と表記されます。
      </p>
      <ul>
        @for (file of data.files; track file.fileName) {
          <li>{{ file.fileName }}：{{ file.failureCount }}件</li>
        }
      </ul>
    </mat-dialog-content>
    <mat-dialog-actions align="end">
      <button mat-button mat-dialog-close>閉じる</button>
    </mat-dialog-actions>
  `,
})
export class FormulaCalculationWarningDialogComponent {
  constructor(
    @Inject(MAT_DIALOG_DATA)
    readonly data: FormulaCalculationWarningDialogData,
  ) {}
}
