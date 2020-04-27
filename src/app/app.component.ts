import { Component, OnInit, ChangeDetectorRef } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import * as moment from 'moment';
import { FileUploadControl } from '@iplab/ngx-file-upload';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
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

  readonly labelTitle = 'Formsの結果を出席簿に集約する';
  readonly labelAboutExplains = [
    'このWebアプリケーション（以下，アプリ）は，Microsoft 365のFormsから得られるエクセルファイルを１つのファイルにまとめ，出欠簿を作成します。',
    'エクセルファイルの「メール」・「開始時刻」・「名前」・「合計得点」を抽出し，集計します。',
    '受講者（ユーザー）は，Microsoft 365にログインしていることが必須です。「anonymous」は集計しません。',
    'アプリは端末内だけで処理が行われ，あなたや受講生の記録がどこかに送られることはありません。',
    'アプリはサーバーの無料枠で運用しています。利用者が多くなった場合，利用ができないかもしれません。その場合は月が変わってからお試しください。',
  ];
  readonly labelHowToTitle = '使い方';
  readonly labelHowToExplains = [
    '以下の欄を通じて，クリックかドラッグでMicrosoft 365のFormsから得られたエクセルファイルを選択・投入します。投入されたファイルは最下部にリスト化されます。',
    '遅刻の時刻を設定したい場合は，右側の日時・時間用の入力欄をクリックして，日時を設定してください。設定された日時自体は「遅刻ではない」と判定されます。',
  ];
  readonly labelStartTime = '開始時刻';
  readonly labelMail = 'メール';
  readonly labelName = '名前';
  readonly labelTotal = '合計得点';
  maxDate: Date;

  mapStudent: Map<string, string> = new Map();

  idTimeout: any = null;
  processing = false;

  constructor(
    private ref: ChangeDetectorRef,
  ) {
    this.maxDate = new Date();
  }

  ngOnInit() {
    this.fileUploadControl.acceptFiles('.xlsx');
    this.fileUploadControl.valueChanges.subscribe(files => {
      // console.log(files);
      this.mapStudent.clear();
      this.mapFileInfo.clear();
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
      this.mapFileInfo.forEach((value, filename, map) => {
        const reader: FileReader = new FileReader();
        reader.onload = (e: any) => {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array' });
          const wsname: string = wb.SheetNames[0];
          const ws: XLSX.WorkSheet = wb.Sheets[wsname];

          const jsonArray = XLSX.utils.sheet_to_json(ws, { raw: false });
          // console.log(jsonArray);
          jsonArray.forEach((j, jaIndex) => {
            if (jaIndex === 0) {
              this.mapFileInfo.get(filename).columnDate = j[this.labelStartTime].split(' ')[0].replace(/\//g, '-');
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
        };
        reader.readAsArrayBuffer(value.file);
      });
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
              const cd = value.lateDate;
              const ct = value.lateTime.split(':');
              // cd.add(Number(ct[0]), 'hour').add(Number(ct[1]) + 1, 'minute');
              if (value.lateTime) {
                cd.hour(Number(ct[0]));
                cd.minute(Number(ct[1]));
              } else {
                cd.add(1, 'days');
              }
              console.log(at[this.labelStartTime]);
              const start = moment(at[this.labelStartTime], 'MM/DD/YY HH:mm:ss');
              console.log(cd.format('YYYY-MM-DD HH:mm:ss'));
              console.log(start.format('YYYY-MM-DD HH:mm:ss'));
              if (start.isBefore(cd)) {
                wsData[stIndex + 1].push(2);
              } else {
                wsData[stIndex + 1].push(1);
              }
            } else {
              wsData[stIndex + 1].push(2);
            }
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
