import { TestBed } from '@angular/core/testing';
import { vi } from 'vitest';
import * as XLSX from 'xlsx';
import { MatDialog } from '@angular/material/dialog';
import { AppComponent } from './app.component';
import { AppModule } from './app.module';
import {
  findUncalculatedNumericFormulas,
  readCalculatedWorkbook,
} from './workbook-formula-calculator';
import { formatTimestamp, parseFormsDateTime } from './forms-date';

async function waitForLoadedFile(app: AppComponent, filename: string): Promise<void> {
  const timeout = Date.now() + 5000;
  while (app.mapFileInfo().get(filename)?.data === undefined) {
    if (Date.now() >= timeout) {
      throw new Error(`Timed out waiting for ${filename} to finish loading`);
    }
    await new Promise(resolve => setTimeout(resolve, 25));
  }
}

describe('AppComponent', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AppModule],
    }).compileComponents();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('should create the app', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    expect(app).toBeTruthy();
  });

  it('should convert a positive integer to an Excel column name', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    expect(app.toColumnName(1)).toBe('A');
    expect(app.toColumnName(27)).toBe('AA');
  });

  it('should format a Forms timestamp as YY-MM-DD', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;

    expect(app.formatLectureDate('07/29/26 10:05:00')).toBe('26-07-29');
    expect(app.formatLectureDate('7/9/2026 9:05:00')).toBe('26-07-09');
  });

  it('should treat the exact deadline as on time and later submissions as late', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    const deadline = new Date(2026, 6, 29);

    expect(app.isSubmissionOnTime('07/29/26 10:00:00', deadline, '10:00')).toBe(true);
    expect(app.isSubmissionOnTime('07/29/26 10:00:01', deadline, '10:00')).toBe(false);
    // 締切時刻の指定がない場合はその日の終わりまでを許容する。
    expect(app.isSubmissionOnTime('07/29/26 23:59:59', deadline, '')).toBe(true);
    expect(app.isSubmissionOnTime('07/30/26 00:00:00', deadline, '')).toBe(false);
    // 日付ピッカーが渡す値は日付の 0 時ちょうどで、比較で書き換えられない。
    expect(formatTimestamp(deadline)).toBe('20260729-000000');
  });

  it('should reject Forms timestamps that are not a supported format', () => {
    expect(parseFormsDateTime('2026-07-29 10:05:00')).toEqual(new Date(2026, 6, 29, 10, 5, 0));
    expect(parseFormsDateTime('07/29/26 10:05')).toEqual(new Date(2026, 6, 29, 10, 5, 0));
    // 存在しない日付・時刻や、区切り文字が混在する値は解析しない。
    expect(parseFormsDateTime('02/30/26 10:05:00')).toBeNull();
    expect(parseFormsDateTime('07/29/26 24:00:00')).toBeNull();
    expect(parseFormsDateTime('07-29/26 10:05:00')).toBeNull();
    expect(parseFormsDateTime('')).toBeNull();
  });

  it('should calculate formula cells without cached values', async () => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['合計点数', null, null, null, null, null, null, '回答'],
      [{ t: 'n', f: 'IF(LEN(H2)>99, 3, IF(LEN(H2)>49, 2, 0))' }, null, null, null, null, null, null, 'あ'.repeat(100)],
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const input = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const calculated = await readCalculatedWorkbook(input);
    const scoreCell = calculated.Sheets['Sheet1']['A2'];

    expect(scoreCell.f).toBe('IF(LEN(H2)>99, 3, IF(LEN(H2)>49, 2, 0))');
    expect(scoreCell.v).toBe(3);
    expect(scoreCell.t).toBe('n');
    expect(await findUncalculatedNumericFormulas(
      calculated.Sheets['Sheet1'],
      '合計点数',
    )).toEqual([]);
  });

  it('should identify score formulas that cannot be calculated', async () => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['合計点数', '回答'],
      [{ t: 'n', f: 'UNSUPPORTED_FUNCTION(B2)' }, '回答内容'],
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const input = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const calculated = await readCalculatedWorkbook(input);
    const failures = await findUncalculatedNumericFormulas(
      calculated.Sheets['Sheet1'],
      '合計点数',
    );

    expect(failures).toEqual([
      {
        address: 'A2',
        formula: 'UNSUPPORTED_FUNCTION(B2)',
      },
    ]);
  });

  it('renders an uploaded workbook row and its lecture date in the mat-table', async () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    fixture.detectChanges();

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['メール', '名前', '開始時刻', '合計点数'],
      ['student@example.com', '山田 太郎', '07/29/26 10:05:00', 100],
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const input = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    app.fileUploadControl.addFile(new File([input], 'a.xlsx'));
    await waitForLoadedFile(app, 'a.xlsx');
    fixture.detectChanges();

    const table = fixture.nativeElement.querySelector('table.mat-mdc-table') as HTMLTableElement;
    expect(table).toBeTruthy();
    const headers = Array.from(table.querySelectorAll('thead th'))
      .map(cell => cell.textContent?.trim());
    expect(headers).toContain('Date');

    const rows = table.querySelectorAll('tbody tr.mat-mdc-row');
    expect(rows.length).toBe(1);
    expect(rows[0].textContent).toContain('a.xlsx');
    expect(rows[0].textContent).toContain('26-07-29');
  });

  it('keeps long file names from causing document horizontal overflow', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    const longFileNames = [
      '心理的アセスメント2・第1回・課題 (2026 心理的アセスメント2（阿部宏徳)V(1-97).xlsx',
      '心理的アセスメント2・第1回・課題 (2026 心理的アセスメント2（阿部宏徳)V(1-98).xlsx',
      '心理的アセスメント2・第1回・課題 (2026 心理的アセスメント2（阿部宏徳)V(1-99).xlsx',
    ];

    fixture.detectChanges();
    app.mapFileInfo.set(new Map(longFileNames.map((fileName, order) => [fileName, {
      file: new File([], fileName),
      lateDate: null,
      lateTime: '',
      order,
    }])));
    fixture.detectChanges();

    const appRoot = fixture.nativeElement as HTMLElement;
    const tableWrapper = appRoot.querySelector<HTMLElement>('.table-scroll-container');
    expect(tableWrapper).not.toBeNull();
    expect(tableWrapper?.clientWidth).toBeGreaterThan(0);
    expect(getComputedStyle(tableWrapper as HTMLElement).overflowX).toBe('auto');
    expect(tableWrapper?.clientWidth).toBeLessThanOrEqual(
      tableWrapper?.parentElement?.clientWidth ?? 0,
    );
    expect(appRoot.scrollWidth).toBeLessThanOrEqual(appRoot.clientWidth);

    expect(document.documentElement.scrollWidth)
      .toBeLessThanOrEqual(document.documentElement.clientWidth);
  });

  it('should show one warning dialog after all files finish loading', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    const dialog = TestBed.inject(MatDialog);
    const openSpy = vi.spyOn(dialog, 'open').mockReturnValue(undefined as any);
    const internalApp = app as any;

    internalApp.pendingFileLoads = 1;
    internalApp.formulaCalculationWarnings.set('未対応の数式.xlsx', 2);
    internalApp.finishFileLoad(0);

    expect(openSpy).toHaveBeenCalledTimes(1);
    const dialogConfig = openSpy.mock.lastCall?.[1] as any;
    expect(dialogConfig.data.files).toEqual([
      {
        fileName: '未対応の数式.xlsx',
        failureCount: 2,
      },
    ]);
  });
});
