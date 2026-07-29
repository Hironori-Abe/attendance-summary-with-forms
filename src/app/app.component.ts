import { Component, OnInit, ChangeDetectorRef } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import moment from 'moment';
import { FileUploadControl } from '@iplab/ngx-file-upload';
import { MatDialog } from '@angular/material/dialog';
import {
  findUncalculatedNumericFormulas,
  readCalculatedWorkbook,
} from './workbook-formula-calculator';
import {
  FormulaCalculationWarningDialogComponent,
} from './formula-calculation-warning-dialog.component';

@Component({
  selector: 'app-root',
  standalone: false,
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  private readonly formsDateTimeFormats = [
    'MM/DD/YY HH:mm:ss',
    'MM/DD/YYYY HH:mm:ss',
    'M/D/YY H:mm:ss',
    'M/D/YYYY H:mm:ss',
    'MM-DD-YY HH:mm:ss',
    'MM-DD-YYYY HH:mm:ss',
    'M-D-YY H:mm:ss',
    'M-D-YYYY H:mm:ss',
    'YYYY/M/D H:mm:ss',
    'YYYY-M-D H:mm:ss',
    'MM/DD/YY HH:mm',
    'MM/DD/YYYY HH:mm',
    'M/D/YY H:mm',
    'M/D/YYYY H:mm',
    'MM-DD-YY HH:mm',
    'MM-DD-YYYY HH:mm',
    'M-D-YY H:mm',
    'M-D-YYYY H:mm',
    'YYYY/M/D H:mm',
    'YYYY-M-D H:mm',
  ];

  public fileUploadControl = new FileUploadControl().setListVisibility(false);
  array = Array.from;
  displayedColumns = ['position', 'name', 'lateTime', 'clear'];

  mapFileInfo: Map<string, {
    file: File;
    columnDate?: string;
    lateDate: moment.Moment;
    lateTime: string;
    order: number;
    data?: any;
  }> = new Map();

  readonly labelTitle = 'Formsの結果を「出席簿」として集約する';
  readonly labelAboutExplains = [
    'このWebアプリケーション（以下，アプリ）は，Microsoft 365のFormsから得られたエクセルファイルを１つのファイルにまとめた「出席簿」を作成します。',
    'エクセルファイルの「メール」・「名前」・「開始時刻」・「合計点数」列を抽出し，「メール」を軸として集約・集計します。',
    'Microsoft 365にログインしていない（＝メール列が「anonymous」になっている）受講者の記録は集計しません。',
    'アプリはこの画面が表示された端末内だけで処理が行われます。あなたや受講生の記録がどこかに送られることは決してありません。',
    'アプリはサーバーの無料枠で運用しています。利用者が多くなった場合，無料枠を超過して利用できないかもしれません。その場合は月が変わってから再度お試しください。',
  ];
  readonly labelHowToTitle = '使い方';
  readonly labelHowToExplains = [
    '以下の欄を通じて，クリックかドラッグでMicrosoft 365のFormsから得られたエクセルファイルを選択・投入します。投入されたファイルは最下部にリスト化されます。',
    '遅刻の時刻を設定したい場合は，右側の日時・時間用の入力欄をクリックして，日時を設定してください。入力された日時自体は「遅刻ではない」と判定されます。設定がない場合，すべての記録が認められます。',
  ];
  readonly labelStartTime = '開始時刻';
  readonly labelMail = 'メール';
  readonly labelName = '名前';
  readonly labelTotal = '合計点数';
  maxDate: Date;

  mapStudent: Map<string, string> = new Map();

  idTimeout: any = null;
  processing = false;
  private fileLoadGeneration = 0;
  private pendingFileLoads = 0;
  private formulaCalculationWarnings = new Map<string, number>();

  constructor(
    private ref: ChangeDetectorRef,
    private dialog: MatDialog,
  ) {
    this.maxDate = new Date();
  }

  ngOnInit() {
    this.fileUploadControl.acceptFiles('.xlsx');
    this.fileUploadControl.valueChanges.subscribe(files => {
      const loadGeneration = ++this.fileLoadGeneration;
      // console.log(files);
      this.mapStudent.clear();
      this.mapFileInfo.clear();
      this.formulaCalculationWarnings.clear();
      let duplicate = false;
      files.forEach(f => {
        if (!this.mapFileInfo.has(f.name)) {
          this.mapFileInfo.set(f.name, {
            file: f,
            lateDate: null,
            lateTime: '',
            order: 0,
          });
          this.ref.detectChanges();
        } else {
          this.fileUploadControl.removeFile(f);
          duplicate = true;
        }
      });
      this.pendingFileLoads = this.mapFileInfo.size;
      this.mapFileInfo.forEach((value, filename, map) => {
        const reader: FileReader = new FileReader();
        reader.onload = (e: any) => {
          if (loadGeneration !== this.fileLoadGeneration) {
            return;
          }

          const data = new Uint8Array(e.target.result);
          const wb = readCalculatedWorkbook(data);
          const wsname: string = wb.SheetNames[0];
          const ws: XLSX.WorkSheet = wb.Sheets[wsname];
          const formulaFailures = findUncalculatedNumericFormulas(ws, this.labelTotal);
          if (formulaFailures.length > 0) {
            this.formulaCalculationWarnings.set(filename, formulaFailures.length);
          }

          const jsonArray = XLSX.utils.sheet_to_json(ws, { raw: false });
          // console.log(jsonArray);
          jsonArray.forEach((j, jaIndex) => {
            if (jaIndex === 0) {
              this.mapFileInfo.get(filename).columnDate = this.formatLectureDate(
                j[this.labelStartTime],
              );
            }
            if (j && j[this.labelMail] && j[this.labelMail] !== 'anonymous') {
              this.mapStudent.set(j[this.labelMail], j[this.labelName]);
            }
          });
          this.mapFileInfo.get(filename).data = jsonArray;

          this.idTimeout = setTimeout(() => {
            clearTimeout(this.idTimeout);
            const newMap: Map<string, string> = new Map();
            const temp = Array.from(this.mapStudent.keys());
            temp.sort();
            temp.forEach(t => {
              newMap.set(t, this.mapStudent.get(t));
            });
            // console.log(temp);
            this.mapStudent = newMap;
          }, 100);
          this.finishFileLoad(loadGeneration);
        };
        reader.onerror = () => this.finishFileLoad(loadGeneration);
        reader.readAsArrayBuffer(value.file);
      });
    });
  }

  private finishFileLoad(loadGeneration: number) {
    if (loadGeneration !== this.fileLoadGeneration) {
      return;
    }

    this.pendingFileLoads--;
    if (this.pendingFileLoads !== 0 || this.formulaCalculationWarnings.size === 0) {
      return;
    }

    this.dialog.open(FormulaCalculationWarningDialogComponent, {
      data: {
        files: Array.from(this.formulaCalculationWarnings, ([fileName, failureCount]) => ({
          fileName,
          failureCount,
        })),
      },
      maxWidth: '640px',
    });
  }

  clickTimepicker(ev: Event) {
    ev.stopPropagation();
  }

  changeDate(name: string, date: moment.Moment) {
    // console.log('name: %s date(moment): %O', name, date);
    this.mapFileInfo.get(name).lateDate = date;
  }

  changeTime(name: string, time: string) {
    // console.log('name: %s time: %s', name, time);
    this.mapFileInfo.get(name).lateTime = time;
  }

  formatLectureDate(startTime: string): string {
    const parsed = moment(startTime, this.formsDateTimeFormats, true);
    if (parsed.isValid()) {
      return parsed.format('YY-MM-DD');
    }

    // Preserve the source value when an unexpected format is encountered.
    return String(startTime ?? '').split(' ')[0].replace(/\//g, '-');
  }

  isSubmissionOnTime(
    startTime: string,
    deadline: moment.Moment,
    deadlineTime: string,
  ): boolean {
    const start = moment(startTime, this.formsDateTimeFormats, true);
    if (!start.isValid()) {
      return false;
    }

    const cutoff = deadline.clone();
    if (deadlineTime) {
      const [hour, minute] = deadlineTime.split(':').map(Number);
      cutoff.hour(hour).minute(minute).second(0).millisecond(0);
    } else {
      cutoff.endOf('day');
    }

    return start.isSameOrBefore(cutoff);
  }

  clearFile(name: string) {
    this.mapFileInfo.delete(name);
    const d = this.fileUploadControl.value.filter(f => {
      return f.name === name;
    });
    this.fileUploadControl.removeFile(d[0]);
  }

  /**
    * Takes a positive integer and returns the corresponding column name.
    * @param {number} num The positive integer to convert to a column name.
    * @return {string} The column name.
    * by Chris West
    * https://cwestblog.com/2013/09/05/javascript-snippet-convert-number-to-column-name/
    */
  toColumnName(num: number): string {
    let ret = '';
    for (let a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
      ret = String.fromCharCode(parseInt(String((num % b) / a)) + 65) + ret;
    }
    return ret;
  }

  createSummary() {
    // console.log(this.mapStudent);
    this.processing = true;
    this.ref.markForCheck();

    const newWB = XLSX.utils.book_new();
    const wsData: Array<Array<any>> = [
      [this.labelMail, this.labelName, 'Total']
    ];
    const wsData2: Array<Array<any>> = JSON.parse(JSON.stringify(wsData));

    this.mapFileInfo.forEach(d => {
      wsData[0].push(d.columnDate);
      wsData2[0].push(d.columnDate);
    });
    wsData[0].push('Attend -> 2  Be late -> 1');
    const column = this.toColumnName(this.mapFileInfo.size + 3);
    // console.log('column:', column);
    let stIndex = 0;
    this.mapStudent.forEach((name, mail, map) => {
      wsData.push([mail, name, { f: `SUM(D${stIndex + 2}:${column}${stIndex + 2})` }]);
      wsData2.push([mail, name, { f: `SUM(D${stIndex + 2}:${column}${stIndex + 2})` }]);

      this.mapFileInfo.forEach(value => {
        // console.log(value);
        let exist = false;
        value.data.forEach(at => {
          if (mail === at[this.labelMail]) {
            if (value.lateDate) {
              if (this.isSubmissionOnTime(
                at[this.labelStartTime],
                value.lateDate,
                value.lateTime,
              )) {
                wsData[stIndex + 1].push(2);
              } else {
                wsData[stIndex + 1].push(1);
              }
            } else {
              wsData[stIndex + 1].push(2);
            }
            // console.log(at[this.labelTotal]);
            if (Number.isFinite(Number(at[this.labelTotal]))) {
              wsData2[stIndex + 1].push(Number(at[this.labelTotal]));
            } else {
              wsData2[stIndex + 1].push(0);
            }
            exist = true;
          }
        });
        if (!exist) {
          wsData[stIndex + 1].push(0);
          wsData2[stIndex + 1].push(0);
        }
      });
      if (stIndex === map.size - 1) {
        let newWS = XLSX.utils.aoa_to_sheet(wsData);
        let newWS2 = XLSX.utils.aoa_to_sheet(wsData2);

        XLSX.utils.book_append_sheet(newWB, newWS, 'Attendance');
        XLSX.utils.book_append_sheet(newWB, newWS2, 'Score');

        var wopts: XLSX.WritingOptions = { bookType: 'xlsx', bookSST: false, type: 'array' };
        const wbout = XLSX.write(newWB, wopts);
        const current = moment().format('YYYYMMDD-HHmmss')
        saveAs(new Blob([wbout], { type: "application/octet-stream" }), `attendance-${current}.xlsx`);
        this.processing = false;
        this.ref.markForCheck();
      }
      stIndex++;
    });
  }
}
