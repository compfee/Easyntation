import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppComponent } from './app.component';
import {TuiButtonModule, TuiLinkModule, TuiLoaderModule, TuiRootModule} from "@taiga-ui/core";
import {TuiInputFilesModule, TuiTextAreaModule} from "@taiga-ui/kit";
import {FormsModule, ReactiveFormsModule} from "@angular/forms";

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    TuiRootModule,
    TuiTextAreaModule,
    TuiInputFilesModule,
    TuiButtonModule,
    FormsModule,
    ReactiveFormsModule,
    TuiLoaderModule,
    TuiLinkModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
