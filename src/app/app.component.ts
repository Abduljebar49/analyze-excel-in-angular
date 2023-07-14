import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'analyze-excel-in-angular';
  excelData: any[] = [];
  excelHeader = []
  fileName: string = "";
  p: number = 1;
  isFileUploading: boolean = false;

  getFileName(file: any) {
    var name = file.name.toString().split('.');
    this.fileName = name[0];
  }

  onStockStatusFileChange($event: any) {
    this.excelData = [];
    this.getFileName($event.target.files[0]);
    this.excelHeader = [];
    this.excelData = [];
    const target: DataTransfer = <DataTransfer>$event.target;
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { raw: false, type: 'binary' });
      const wsname: string = wb.SheetNames[0];

      const ws: XLSX.WorkSheet = wb.Sheets[wsname];
      this.excelData = XLSX.utils.sheet_to_json(ws, { header: 1 });
      this.excelHeader = this.excelData[0];
      this.excelData = this.excelData.slice(1);
      this.checkStockStatusCompatibility();
    };
    reader.readAsBinaryString(target.files[0]);
  }

  checkStockStatusCompatibility() {

  }

  uploadStockStatus() {
    this.isFileUploading = true;
  }
}
