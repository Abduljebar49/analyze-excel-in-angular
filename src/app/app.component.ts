import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import * as fileSaver from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'analyze-excel-in-angular';
  excelData: any[] = [];
  convertedData: any[] = [];
  excelHeader = []
  fileName: string = "";
  p: number = 1;
  isFileUploading: boolean = false;
  //               0         1       2         3              4             5'
  columnTypes = ['gender', 'yesNo', 'age', 'effectOrNo', 'agreeNotAgree', 'default', 'education', 'experience', 'size', 'defaultNew']
  dataTypeHeader = [0, 2, 6, 7, 7, 8, 3, 5, 5, 5, 9, 9, 1, 1, 3, 3, 5, 5, 5, 5, 4, 4, 4, 4];
  headerConstant = ['sex', 'age', 'education', 'experience', 'organizationAge', 'organizationSize', 'recruitmentProcessEffective', 'qualifiedApplicant', 'performanceImprove', 'employeeOrientation', 'employeeCost', 'vacancyPriority', 'trainingOrganization', 'training', 'effectivenessTraining', 'effectivenessPerformance', 'performanceFeedback', 'weaknessPerformance', 'employeePerformance', 'perfomanceDevelopment', 'encourageTeamwork', 'effortEmployeeOpinion', 'otherDepartment', 'distributionWork']
  headers = ['sex', 'age', 'education', 'experience', 'organizationAge', 'organizationSize', 'recruitmentProcessEffective', 'qualifiedApplicant', 'performanceImprove', 'employeeOrientation', 'employeeCost', 'vacancyPriority', 'trainingOrganization', 'training', 'effectivenessTraining', 'effectivenessPerformance', 'performanceFeedback', 'weaknessPerformance', 'employeePerformance', 'perfomanceDevelopment', 'encourageTeamwork', 'effortEmployeeOpinion', 'otherDepartment', 'distributionWork']
  gender = {
    'male': 1,
    'female': 2
  }
  yesNo = {
    'yes': 1,
    'no': 0
  }
  age = {
    '20-25': 0,
    '25-35': 1,
    '35-45': 2,
    '45+': 3
  }
  effectiveOrNot = {
    'very effective': 2,
    'effective': 1,
    'less effective': 0
  }

  default = {
    'to great extent': 4,
    'to some extent': 3,
    'moderate': 2,
    'to small extent': 1,
    'not at all': 0
  }

  defaultNew = {
    'to great extent': 4,
    'to some extent': 3,
    'moderate extent': 2,
    'to small extent': 1,
    'not at all': 0
  }

  agreeOrNot = {
    'strongly agree': 5,
    'somehow agree': 4,
    'agree': 3,
    'neutral': 2,
    'disagree': 1,
    'strongly disagree': 0
  }
  experience = {
    '2-5 years': 0,
    '5-10 years': 1,
    '10-15 years': 2,
    '15+ years': 3
  }

  education = {
    'high school deploma': 0,
    'graduate': 1,
    'post graduate': 2,
    'ms/phd': 3
  }

  size = {
    '5-10': 0,
    '10-50': 1,
    '50-100': 2,
    '100-200': 3,
    '200-500': 4,
    '500+': 5
  }


  convertData() {
    var newData: any[] = [];

    this.excelData.forEach((ele) => {
      var newEle = [];
      for (let i = 0; i < ele.length; i++) {
        const type = this.dataTypeHeader[i];
        newEle.push(this.getValue(type, ele[i]));
      }
      newData.push(newEle);
      this.convertedData.push(newEle)
    })

  }
  //  columnTypes = ['gender', 'yesNo', 'age', 'effectOrNo', 'agreeNotAgree', 'default', 'education', 'experience', 'size', 'defaultNew']

  getValue(type: number, value: string) {
    switch (this.columnTypes[type]) {
      case this.columnTypes[0]:
        return this.getValueFromObject(this.gender, value);
      case this.columnTypes[1]:
        return this.getValueFromObject(this.yesNo, value);
      case this.columnTypes[2]:
        return this.getValueFromObject(this.age, value);
      case this.columnTypes[3]:
        return this.getValueFromObject(this.effectiveOrNot, value);
      case this.columnTypes[4]:
        return this.getValueFromObject(this.agreeOrNot, value);
      case this.columnTypes[5]:
        return this.getValueFromObject(this.default, value);
      case this.columnTypes[6]:
        return this.getValueFromObject(this.education, value);
      case this.columnTypes[7]:
        return this.getValueFromObject(this.experience, value);
      case this.columnTypes[8]:
        return this.getValueFromObject(this.size, value);
      case this.columnTypes[9]:
        return this.getValueFromObject(this.defaultNew, value);
    }
  }

  getValueFromObject(object: any, value: string) {
    var index = "";
    if (value) {
      try {
        index = value.toLocaleLowerCase();
      } catch (e) {
        return 0
      }
    }
    else
      return 0;
    return object[index];
  }

  getFileName(file: any) {
    var name = file.name.toString().split('.');
    this.fileName = name[0];
  }

  onFileChange($event: any) {
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
      console.log(this.excelHeader);
      console.log(this.excelData);
      this.checkStockStatusCompatibility();
    };
    reader.readAsBinaryString(target.files[0]);
  }

  checkStockStatusCompatibility() {

  }

  uploadStockStatus() {
    this.isFileUploading = true;
  }
  generateJSON() {
    const result = this.changeGivenDataToJSON(this.excelData)
    this.downloadJson(result);
  }

  changeGivenDataToJSON(data:any[]) {
    const result = data.reduce((acc, cur) => {
      const values = Object.values(cur);
      acc.push(this.excelHeader.reduce((obj: any, header, i) => {
        obj[header] = values[i];
        return obj;
      }, {}));
      return acc;
    }, []);
    return result;// this.stockStatusJson = result;
  }

  downloadJson(jsonData: any): void {
    const json = JSON.stringify(jsonData);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'myJsonFile.json';
    a.click();
  }


  exportToExacel() {
    const result = this.changeGivenDataToJSON(this.convertedData)
    this.exportAsExcelFile(result, 'converted data');
  }

  public exportAsExcelFile(json: any[], excelFileName: string): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet },
      SheetNames: ['data'],
    };
    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });
    this.saveAsExcelFile(excelBuffer, excelFileName);
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], { type: EXCEL_TYPE });
    fileSaver.saveAs(
      data,
      fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION
    );
  }
}
