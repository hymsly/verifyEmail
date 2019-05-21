// app.component.ts

import { Component, OnInit } from '@angular/core';
import { FileUploader, FileSelectDirective } from 'ng2-file-upload/ng2-file-upload';
import * as FileSaver from 'file-saver';
import {HttpParams} from  "@angular/common/http";
import { HttpClient,HttpHandler } from  "@angular/common/http";
import { fillProperties } from '@angular/core/src/util/property';
const URL = 'http://54.159.148.36:3000/api/upload';
//const URL = 'http://localhost:3000/api/upload';
const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'app';
  excel : any;
  public uploader: FileUploader = new FileUploader({ url: URL, itemAlias: 'photo' });
  constructor(private http: HttpClient) { }
  ngOnInit() {
    this.uploader.onAfterAddingFile = (file) => { file.withCredentials = false; };
    this.uploader.onCompleteItem = (item: any, response: any, status: any, headers: any) => {
      //console.log('ImageUpload:uploaded:', item, status, response);
      //console.log(response);
      let filename = JSON.parse(response).filename
      this.excel = this.http.get(URL + '/'+filename,{ responseType: 'blob'}).subscribe(res=>{
          const data: Blob = new Blob([res], {type: EXCEL_TYPE});
          FileSaver.saveAs(data, "resultado.xlsx");
          alert('File uploaded successfully');
      })
      
    };
  }
}