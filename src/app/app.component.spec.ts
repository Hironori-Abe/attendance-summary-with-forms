import { TestBed } from '@angular/core/testing';
import * as XLSX from 'xlsx';
import { MatDialog } from '@angular/material/dialog';
import moment from 'moment';
import { AppComponent } from './app.component';
import { AppModule } from './app.module';
import {
  findUncalculatedNumericFormulas,
  readCalculatedWorkbook,
} from './workbook-formula-calculator';

describe('AppComponent', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AppModule],
    }).compileComponents();
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
    const deadline = moment('2026-07-29');

    expect(app.isSubmissionOnTime('07/29/26 10:00:00', deadline, '10:00')).toBeTrue();
    expect(app.isSubmissionOnTime('07/29/26 10:00:01', deadline, '10:00')).toBeFalse();
    expect(deadline.format('YYYY-MM-DD HH:mm:ss')).toBe('2026-07-29 00:00:00');
  });

  it('should calculate formula cells without cached values', () => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['合計点数', null, null, null, null, null, null, '回答'],
      [{ t: 'n', f: 'IF(LEN(H2)>99, 3, IF(LEN(H2)>49, 2, 0))' }, null, null, null, null, null, null, 'あ'.repeat(100)],
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const input = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const calculated = readCalculatedWorkbook(input);
    const scoreCell = calculated.Sheets['Sheet1']['A2'];

    expect(scoreCell.f).toBe('IF(LEN(H2)>99, 3, IF(LEN(H2)>49, 2, 0))');
    expect(scoreCell.v).toBe(3);
    expect(scoreCell.t).toBe('n');
    expect(findUncalculatedNumericFormulas(calculated.Sheets['Sheet1'], '合計点数')).toEqual([]);
  });

  it('should identify score formulas that cannot be calculated', () => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['合計点数', '回答'],
      [{ t: 'n', f: 'UNSUPPORTED_FUNCTION(B2)' }, '回答内容'],
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const input = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const calculated = readCalculatedWorkbook(input);
    const failures = findUncalculatedNumericFormulas(
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

  it('should show one warning dialog after all files finish loading', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    const dialog = TestBed.inject(MatDialog);
    const openSpy = spyOn(dialog, 'open');
    const internalApp = app as any;

    internalApp.pendingFileLoads = 1;
    internalApp.formulaCalculationWarnings.set('未対応の数式.xlsx', 2);
    internalApp.finishFileLoad(0);

    expect(openSpy).toHaveBeenCalledTimes(1);
    const dialogConfig = openSpy.calls.mostRecent().args[1] as any;
    expect(dialogConfig.data.files).toEqual([
      {
        fileName: '未対応の数式.xlsx',
        failureCount: 2,
      },
    ]);
  });
});
