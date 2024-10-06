import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { NgToastModule } from 'ng-angular-popup';
import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { UploadsectionComponent } from './uploadsection/uploadsection.component';
import { HttpClientModule } from '@angular/common/http';
import { CustomPipe } from './custom.pipe';

@NgModule({

  declarations: [
    AppComponent,
    UploadsectionComponent,
    CustomPipe
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    NgToastModule,
    HttpClientModule,    



  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
