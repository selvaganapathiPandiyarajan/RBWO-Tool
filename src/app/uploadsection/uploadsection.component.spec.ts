import { ComponentFixture, TestBed } from '@angular/core/testing';

import { UploadsectionComponent } from './uploadsection.component';

describe('UploadsectionComponent', () => {
  let component: UploadsectionComponent;
  let fixture: ComponentFixture<UploadsectionComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ UploadsectionComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(UploadsectionComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
