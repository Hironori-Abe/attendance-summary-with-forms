import { Component, ChangeDetectionStrategy } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { By } from '@angular/platform-browser';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatTimepickerModule } from '@angular/material/timepicker';
import { AppComponent } from './app.component';
import { AppModule } from './app.module';

@Component({
  selector: 'test-deadline-time-input',
  standalone: false,
  changeDetection: ChangeDetectionStrategy.Eager,
  template: `
    <mat-form-field>
      <mat-label>Time</mat-label>
      <input matInput [matTimepicker]="timePicker" readonly>
      <mat-timepicker-toggle matSuffix [for]="timePicker"></mat-timepicker-toggle>
      <mat-timepicker #timePicker interval="10min"
        (selected)="selectTime($event.value)"></mat-timepicker>
    </mat-form-field>
  `,
})
class DeadlineTimeInputHostComponent {
  selectedTime: Date | null = null;

  selectTime(time: Date) {
    this.selectedTime = time;
  }
}

describe('deadline time input', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AppModule, MatFormFieldModule, MatInputModule, MatTimepickerModule],
      declarations: [DeadlineTimeInputHostComponent],
    }).compileComponents();
  });

  it('selects a Material time option and emits its value as a Date', () => {
    const fixture = TestBed.createComponent(DeadlineTimeInputHostComponent);
    const host = fixture.componentInstance;
    fixture.detectChanges();

    const timeInput = fixture.nativeElement.querySelector('.mat-timepicker-input') as HTMLInputElement;
    expect(timeInput).toBeTruthy();
    const timepicker = fixture.debugElement.query(By.css('mat-timepicker')).componentInstance;
    timepicker.open();
    fixture.detectChanges();

    const timeOption = Array.from(document.body.querySelectorAll('mat-option'))
      .find(option => option.textContent?.trim() === '10:30') as HTMLElement;
    expect(timeOption).toBeTruthy();
    timeOption.click();
    fixture.detectChanges();

    expect(timeInput.value).toBe('10:30');
    expect(host.selectedTime instanceof Date).toBeTrue();
    expect(host.selectedTime?.getHours()).toBe(10);
    expect(host.selectedTime?.getMinutes()).toBe(30);
  });

  it('generates time options at ten-minute intervals', () => {
    const fixture = TestBed.createComponent(DeadlineTimeInputHostComponent);
    fixture.detectChanges();

    const timepicker = fixture.debugElement.query(By.css('mat-timepicker')).componentInstance;
    timepicker.open();
    fixture.detectChanges();

    const panel = document.body.querySelector('.mat-timepicker-panel');
    expect(panel).not.toBeNull();
    const labels = Array.from(panel?.querySelectorAll('mat-option') ?? [])
      .map(option => option.textContent?.trim());
    expect(labels).toContain('10:10');
    expect(labels).not.toContain('10:05');

    timepicker.close();
  });

  it('stores a selected time as HH:mm through AppComponent.changeTime()', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    app.mapFileInfo.set(new Map([
      ['sample.xlsx', {
        file: new File([], 'sample.xlsx'),
        lateDate: null,
        lateTime: '',
        order: 0,
      }],
    ]));

    app.changeTime('sample.xlsx', new Date(2026, 0, 1, 9, 5));

    expect(app.mapFileInfo().get('sample.xlsx')?.lateTime).toBe('09:05');
  });
});
