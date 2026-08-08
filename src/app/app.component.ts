import { afterNextRender, Component, DOCUMENT, inject, Injector, OnInit, computed, signal } from '@angular/core';
import type * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { FileUploadControl } from '@iplab/ngx-file-upload';
import { MatDialog } from '@angular/material/dialog';
import {
  findUncalculatedNumericFormulas,
  readCalculatedWorkbook,
} from './workbook-formula-calculator';
import {
  FormulaCalculationWarningDialogComponent,
} from './formula-calculation-warning-dialog.component';
import { formatShortDate, formatTimestamp, parseFormsDateTime } from './forms-date';
import { loadXlsx } from './xlsx-loader';

type AttendanceRow = Record<string, unknown>;

interface FileInfo {
  file: File;
  columnDate?: string;
  lateDate: Date | null;
  lateTime: string;
  order: number;
  data?: AttendanceRow[];
}

@Component({
  selector: 'app-root',
  standalone: false,
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  public fileUploadControl = new FileUploadControl().setListVisibility(false);
  displayedColumns = ['position', 'name', 'date', 'lateTime', 'clear'];

  private readonly injector = inject(Injector);
  private readonly document = inject(DOCUMENT);

  readonly mapFileInfo = signal(new Map<string, FileInfo>());
  readonly fileNames = computed(() => Array.from(this.mapFileInfo().keys()));

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

  readonly mapStudent = signal(new Map<string, string>());

  idTimeout: ReturnType<typeof setTimeout> | null = null;
  readonly processing = signal(false);
  private fileLoadGeneration = 0;
  private pendingFileLoads = 0;
  private formulaCalculationWarnings = new Map<string, number>();

  constructor(
    private dialog: MatDialog,
  ) {
    this.maxDate = new Date();
  }

  ngOnInit() {
    this.fileUploadControl.acceptFiles('.xlsx');
    this.fileUploadControl.valueChanges.subscribe(files => {
      const loadGeneration = ++this.fileLoadGeneration;
      // console.log(files);
      if (this.idTimeout !== null) {
        clearTimeout(this.idTimeout);
        this.idTimeout = null;
      }
      this.mapStudent.set(new Map());
      this.mapFileInfo.set(new Map());
      this.formulaCalculationWarnings.clear();
      const nextFileInfo = new Map<string, FileInfo>();
      files.forEach(f => {
        if (!nextFileInfo.has(f.name)) {
          nextFileInfo.set(f.name, {
            file: f,
            lateDate: null,
            lateTime: '',
            order: 0,
          });
        } else {
          this.fileUploadControl.removeFile(f);
        }
      });
      this.mapFileInfo.set(nextFileInfo);
      this.pendingFileLoads = nextFileInfo.size;
      nextFileInfo.forEach((value, filename) => {
        const reader: FileReader = new FileReader();
        reader.onload = async (e: any) => {
          if (loadGeneration !== this.fileLoadGeneration) {
            return;
          }

          try {
            const data = new Uint8Array(e.target.result);
            const wb = await readCalculatedWorkbook(data);
            const xlsx = await loadXlsx();
            if (loadGeneration !== this.fileLoadGeneration || !this.mapFileInfo().has(filename)) {
              return;
            }
            const wsname: string = wb.SheetNames[0];
            const ws: XLSX.WorkSheet = wb.Sheets[wsname];
            const formulaFailures = await findUncalculatedNumericFormulas(ws, this.labelTotal);
            if (formulaFailures.length > 0) {
              const nextWarnings = new Map(this.formulaCalculationWarnings);
              nextWarnings.set(filename, formulaFailures.length);
              this.formulaCalculationWarnings = nextWarnings;
            }

            const jsonArray = xlsx.utils.sheet_to_json<AttendanceRow>(ws, { raw: false });
            // console.log(jsonArray);
            let columnDate: string | undefined;
            const nextStudents = new Map(this.mapStudent());
            jsonArray.forEach((j, jaIndex) => {
              if (jaIndex === 0) {
                columnDate = this.formatLectureDate(j[this.labelStartTime] as string);
              }
              const mail = j[this.labelMail] as string | undefined;
              if (mail && mail !== 'anonymous') {
                nextStudents.set(mail, j[this.labelName] as string);
              }
            });
            this.mapStudent.set(nextStudents);

            const currentFileInfo = this.mapFileInfo().get(filename);
            if (!currentFileInfo) {
              return;
            }
            const nextFileInfo = new Map(this.mapFileInfo());
            nextFileInfo.set(filename, {
              ...currentFileInfo,
              columnDate,
              data: jsonArray,
            });
            this.mapFileInfo.set(nextFileInfo);

            if (this.idTimeout !== null) {
              clearTimeout(this.idTimeout);
            }
            this.idTimeout = setTimeout(() => {
              this.idTimeout = null;
              if (loadGeneration !== this.fileLoadGeneration) {
                return;
              }
              const newMap: Map<string, string> = new Map();
              const temp = Array.from(this.mapStudent().keys());
              temp.sort();
              temp.forEach(t => {
                const name = this.mapStudent().get(t);
                if (name !== undefined) {
                  newMap.set(t, name);
                }
              });
              // console.log(temp);
              this.mapStudent.set(newMap);
            }, 100);
          } finally {
            this.finishFileLoad(loadGeneration);
          }
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

  changeDate(name: string, date: Date) {
    // console.log('name: %s date: %O', name, date);
    const fileInfo = this.mapFileInfo().get(name);
    if (fileInfo) {
      const nextFileInfo = new Map(this.mapFileInfo());
      nextFileInfo.set(name, {...fileInfo, lateDate: date});
      this.mapFileInfo.set(nextFileInfo);
    }
  }

  changeTime(name: string, time: Date) {
    // console.log('name: %s time: %O', name, time);
    const hour = String(time.getHours()).padStart(2, '0');
    const minute = String(time.getMinutes()).padStart(2, '0');
    const fileInfo = this.mapFileInfo().get(name);
    if (fileInfo) {
      const nextFileInfo = new Map(this.mapFileInfo());
      nextFileInfo.set(name, {...fileInfo, lateTime: `${hour}:${minute}`});
      this.mapFileInfo.set(nextFileInfo);
    }
  }

  scrollTimepickerToMorning() {
    afterNextRender(() => {
      const panel = this.document.querySelector<HTMLElement>(
        '.mat-timepicker-panel:not(.mat-timepicker-panel-exit)',
      );
      const morningOption = Array.from(
        panel?.querySelectorAll<HTMLElement>('mat-option') ?? [],
      ).find(option => {
        const [hour, minute] = option.textContent?.trim().split(':') ?? [];
        return Number(hour) === 9 && minute === '00';
      });

      morningOption?.scrollIntoView({
        block: 'center',
        inline: 'nearest',
      });
    }, { injector: this.injector });
  }

  formatLectureDate(startTime: string): string {
    const parsed = parseFormsDateTime(startTime);
    if (parsed) {
      return formatShortDate(parsed);
    }

    // Preserve the source value when an unexpected format is encountered.
    return String(startTime ?? '').split(' ')[0].replace(/\//g, '-');
  }

  isSubmissionOnTime(
    startTime: string,
    deadline: Date,
    deadlineTime: string,
  ): boolean {
    const start = parseFormsDateTime(startTime);
    if (!start) {
      return false;
    }

    const cutoff = new Date(deadline);
    if (deadlineTime) {
      const [hour, minute] = deadlineTime.split(':').map(Number);
      cutoff.setHours(hour, minute, 0, 0);
    } else {
      cutoff.setHours(23, 59, 59, 999);
    }

    return start.getTime() <= cutoff.getTime();
  }

  clearFile(name: string) {
    const nextFileInfo = new Map(this.mapFileInfo());
    nextFileInfo.delete(name);
    this.mapFileInfo.set(nextFileInfo);
    const d = this.fileUploadControl.value.filter(f => {
      return f.name === name;
    });
    if (d[0]) {
      this.fileUploadControl.removeFile(d[0]);
    }
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

  async createSummary() {
    // console.log(this.mapStudent());
    this.processing.set(true);
    try {
      const xlsx = await loadXlsx();
      const newWB = xlsx.utils.book_new();
      const wsData: Array<Array<any>> = [
        [this.labelMail, this.labelName, 'Total']
      ];
      const wsData2: Array<Array<any>> = JSON.parse(JSON.stringify(wsData));
      const fileInfoMap = this.mapFileInfo();
      const studentMap = this.mapStudent();

      fileInfoMap.forEach(d => {
        wsData[0].push(d.columnDate);
        wsData2[0].push(d.columnDate);
      });
      wsData[0].push('Attend -> 2  Be late -> 1');
      const column = this.toColumnName(fileInfoMap.size + 3);
      // console.log('column:', column);
      let stIndex = 0;
      studentMap.forEach((name, mail) => {
        wsData.push([mail, name, { f: `SUM(D${stIndex + 2}:${column}${stIndex + 2})` }]);
        wsData2.push([mail, name, { f: `SUM(D${stIndex + 2}:${column}${stIndex + 2})` }]);

        fileInfoMap.forEach(value => {
          // console.log(value);
          let exist = false;
          (value.data ?? []).forEach(at => {
            if (mail === at[this.labelMail]) {
              if (value.lateDate) {
                if (this.isSubmissionOnTime(
                  at[this.labelStartTime] as string,
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
        stIndex++;
      });

      if (studentMap.size > 0) {
        const newWS = xlsx.utils.aoa_to_sheet(wsData);
        const newWS2 = xlsx.utils.aoa_to_sheet(wsData2);

        xlsx.utils.book_append_sheet(newWB, newWS, 'Attendance');
        xlsx.utils.book_append_sheet(newWB, newWS2, 'Score');

        const wopts: XLSX.WritingOptions = { bookType: 'xlsx', bookSST: false, type: 'array' };
        const wbout = xlsx.write(newWB, wopts);
        const current = formatTimestamp(new Date());
        saveAs(new Blob([wbout], { type: "application/octet-stream" }), `attendance-${current}.xlsx`);
      }
    } finally {
      this.processing.set(false);
    }
  }
}
