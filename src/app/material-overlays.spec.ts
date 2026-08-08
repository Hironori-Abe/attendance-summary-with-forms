import { ChangeDetectionStrategy, Component } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { By } from '@angular/platform-browser';
import { OverlayContainer } from '@angular/cdk/overlay';
import { MatNativeDateModule } from '@angular/material/core';
import { MatDatepicker } from '@angular/material/datepicker';
import { MatDatepickerModule } from '@angular/material/datepicker';
import { MatDialog } from '@angular/material/dialog';
import { MatDialogModule } from '@angular/material/dialog';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatTimepicker } from '@angular/material/timepicker';
import { MatTimepickerModule } from '@angular/material/timepicker';
import { AppComponent } from './app.component';
import { AppModule } from './app.module';
import { FormulaCalculationWarningDialogComponent } from './formula-calculation-warning-dialog.component';

@Component({
  selector: 'test-material-overlay-host',
  standalone: false,
  changeDetection: ChangeDetectionStrategy.Eager,
  template: `
    <mat-form-field>
      <mat-label>Date</mat-label>
      <input matInput [matDatepicker]="datePicker" readonly>
      <mat-datepicker #datePicker></mat-datepicker>
    </mat-form-field>

    <mat-form-field>
      <mat-label>Time</mat-label>
      <input matInput [matTimepicker]="timePicker" readonly>
      <mat-timepicker #timePicker interval="10min"></mat-timepicker>
    </mat-form-field>

    <button type="button" (click)="openDialog()">Open dialog</button>
  `,
})
class MaterialOverlayHostComponent {
  constructor(private readonly dialog: MatDialog) {}

  openDialog() {
    this.dialog.open(FormulaCalculationWarningDialogComponent, {
      data: {
        files: [{fileName: 'sample.xlsx', failureCount: 1}],
      },
    });
  }
}

describe('Material overlays without the legacy animation module', () => {
  let overlayContainer: OverlayContainer;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [
        AppModule,
        FormulaCalculationWarningDialogComponent,
        MatDatepickerModule,
        MatDialogModule,
        MatFormFieldModule,
        MatInputModule,
        MatNativeDateModule,
        MatTimepickerModule,
      ],
      declarations: [MaterialOverlayHostComponent],
    }).compileComponents();

    overlayContainer = TestBed.inject(OverlayContainer);
  });

  afterEach(() => {
    TestBed.inject(MatDialog).closeAll();
    overlayContainer.getContainerElement().innerHTML = '';
  });

  it('opens the datepicker and renders calendar date cells', async () => {
    const fixture = TestBed.createComponent(MaterialOverlayHostComponent);
    fixture.detectChanges();

    const datepickerDebugElement = fixture.debugElement.query(By.directive(MatDatepicker));
    expect(datepickerDebugElement).toBeTruthy();
    const datepicker = datepickerDebugElement.componentInstance as MatDatepicker<Date>;

    datepicker.open();
    fixture.detectChanges();
    await fixture.whenStable();
    fixture.detectChanges();

    expect(datepicker.opened).toBeTrue();
    const overlay = overlayContainer.getContainerElement();
    expect(overlay.querySelector('mat-datepicker-content')).not.toBeNull();
    expect(overlay.querySelectorAll('.mat-calendar-body-cell').length).toBeGreaterThan(0);
  });

  it('opens the timepicker and renders time options', async () => {
    const fixture = TestBed.createComponent(MaterialOverlayHostComponent);
    fixture.detectChanges();

    const timepickerDebugElement = fixture.debugElement.query(By.directive(MatTimepicker));
    expect(timepickerDebugElement).toBeTruthy();
    const timepicker = timepickerDebugElement.componentInstance as MatTimepicker<Date>;

    timepicker.open();
    fixture.detectChanges();
    await fixture.whenStable();
    fixture.detectChanges();

    expect(timepicker.isOpen()).toBeTrue();
    const timepickerPanel = document.body.querySelector('.mat-timepicker-panel');
    expect(timepickerPanel).not.toBeNull();
    const options = timepickerPanel?.querySelectorAll('mat-option') ?? [];
    expect(options.length).toBeGreaterThan(0);
    expect(Array.from(options).some(option => option.textContent?.trim() === '10:30')).toBeTrue();
  });

  it('scrolls the app timepicker to 09:00 after it opens', async () => {
    const fixture = TestBed.createComponent(AppComponent);
    fixture.detectChanges();
    fixture.componentInstance.mapFileInfo.set(new Map([
      ['sample.xlsx', {
        file: new File([], 'sample.xlsx'),
        lateDate: null,
        lateTime: '',
        order: 0,
      }],
    ]));
    fixture.detectChanges();

    const timepickerDebugElement = fixture.debugElement.query(By.directive(MatTimepicker));
    expect(timepickerDebugElement).toBeTruthy();
    const timepicker = timepickerDebugElement.componentInstance as MatTimepicker<Date>;

    timepicker.open();
    fixture.detectChanges();
    await fixture.whenStable();
    fixture.detectChanges();

    const panel = document.body.querySelector<HTMLElement>('.mat-timepicker-panel');
    expect(panel).not.toBeNull();
    const morningOption = Array.from(panel?.querySelectorAll<HTMLElement>('mat-option') ?? [])
      .find(option => /^0?9:00$/.test(option.textContent?.trim() ?? ''));
    expect(morningOption).toBeTruthy();

    const optionTop = morningOption?.offsetTop ?? 0;
    const optionBottom = optionTop + (morningOption?.offsetHeight ?? 0);
    expect(panel?.scrollTop ?? 0).toBeGreaterThan(0);
    expect(optionTop).toBeLessThan((panel?.scrollTop ?? 0) + (panel?.clientHeight ?? 0));
    expect(optionBottom).toBeGreaterThan(panel?.scrollTop ?? 0);

    timepicker.close();
  });

  it('scrolls the newly opened timepicker while the previous panel is exiting', async () => {
    const fixture = TestBed.createComponent(AppComponent);
    fixture.detectChanges();
    fixture.componentInstance.mapFileInfo.set(new Map([
      ['first.xlsx', {
        file: new File([], 'first.xlsx'),
        lateDate: null,
        lateTime: '',
        order: 0,
      }],
      ['second.xlsx', {
        file: new File([], 'second.xlsx'),
        lateDate: null,
        lateTime: '',
        order: 0,
      }],
    ]));
    fixture.detectChanges();

    const timepickers = fixture.debugElement.queryAll(By.directive(MatTimepicker))
      .map(debugElement => debugElement.componentInstance as MatTimepicker<Date>);
    expect(timepickers.length).toBe(2);

    timepickers[0].open();
    fixture.detectChanges();
    await fixture.whenStable();
    fixture.detectChanges();

    timepickers[0].close();
    fixture.detectChanges();
    timepickers[1].open();
    fixture.detectChanges();

    const panelsDuringTransition = document.body.querySelectorAll<HTMLElement>('.mat-timepicker-panel');
    expect(panelsDuringTransition.length).toBeGreaterThan(1);
    expect(document.body.querySelector('.mat-timepicker-panel.mat-timepicker-panel-exit'))
      .not.toBeNull();

    await fixture.whenStable();
    fixture.detectChanges();

    const panel = document.body.querySelector<HTMLElement>(
      '.mat-timepicker-panel:not(.mat-timepicker-panel-exit)',
    );
    expect(panel).not.toBeNull();
    const morningOption = Array.from(panel?.querySelectorAll<HTMLElement>('mat-option') ?? [])
      .find(option => /^0?9:00$/.test(option.textContent?.trim() ?? ''));
    expect(morningOption).toBeTruthy();

    const optionTop = morningOption?.offsetTop ?? 0;
    const optionBottom = optionTop + (morningOption?.offsetHeight ?? 0);
    expect(panel?.scrollTop ?? 0).toBeGreaterThan(0);
    expect(optionTop).toBeLessThan((panel?.scrollTop ?? 0) + (panel?.clientHeight ?? 0));
    expect(optionBottom).toBeGreaterThan(panel?.scrollTop ?? 0);

    timepickers[1].close();
  });

  it('opens the warning dialog and renders its dialog container', async () => {
    const fixture = TestBed.createComponent(MaterialOverlayHostComponent);
    fixture.detectChanges();

    fixture.componentInstance.openDialog();
    fixture.detectChanges();
    await fixture.whenStable();
    fixture.detectChanges();

    const overlay = overlayContainer.getContainerElement();
    expect(overlay.querySelector('mat-dialog-container')).not.toBeNull();
    expect(overlay.querySelector('mat-dialog-container')?.textContent).toContain('数式を計算できませんでした');
  });
});
