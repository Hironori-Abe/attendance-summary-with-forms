import { TestBed } from '@angular/core/testing';
import { AppComponent } from './app.component';
import { AppModule } from './app.module';

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
});
