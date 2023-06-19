import {Component, OnInit} from '@angular/core';
import {TuiFileLike, TuiFileState} from "@taiga-ui/kit";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.less']
})
export class AppComponent implements OnInit{
  ngOnInit(): void {
    setTimeout(()=>this.documentState = 'normal', 10000);
  }

  title: string = 'Easyntation';

  documentState: TuiFileState = 'loading';
  isBtnClicked = false;

  sendDocument = false;

  loadingFile: TuiFileLike | null = {
    name: 'Your document',
  };

  removeLoading(): void {
    this.loadingFile = null;
  }

  prepaireDoc() {
    this.isBtnClicked = false;
    this.sendDocument = true;
  }

  handleConvertToPptxButton(){
    this.isBtnClicked = true;
    setTimeout(()=>this.prepaireDoc(), 20000);
  }

  ngModelChange(file: TuiFileLike | undefined) {
    console.log('test')
    this.loadingFile = file!;

    setTimeout(()=>this.documentState = 'normal', 10000);
  }
}
