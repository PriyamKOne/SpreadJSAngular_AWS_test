import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import '@grapecity/spread-sheets-designer-resources-en';
import { DesignerModule } from '@grapecity/spread-sheets-designer-angular';
import { SpreadSheetsModule } from "@grapecity/spread-sheets-angular";

import { AppComponent } from './app.component';

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    DesignerModule,
    SpreadSheetsModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
