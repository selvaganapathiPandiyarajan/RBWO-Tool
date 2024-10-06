import { Component, OnInit, HostListener, ɵAPP_ID_RANDOM_PROVIDER } from "@angular/core";
import * as XLSX from 'xlsx';
import { NgToastService } from 'ng-angular-popup';
import { HttpClient } from '@angular/common/http';
import * as moment from 'moment';
import { Workbook, Worksheet, Row, Cell } from 'exceljs';

import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

@Component({
  selector: 'app-uploadsection',
  templateUrl: './uploadsection.component.html',
  styleUrls: ['./uploadsection.component.css']
})
export class UploadsectionComponent implements OnInit {
  name = 'Angular 5';
  width: any = 0;
  totalBuffer: any;
  totalQAvalue: any;
  stylebtn: any;
  totalFtevalue: any;
  datafromexcel = true;
  inbetweenholiday: any = [];
  enddateArr: any = [];
  endMontharr: any = [];
  generatebtn = true;
  companyholidayArr: any = [];
  endeffortArr: any = [];
  endmonthkey: any = [];
  eventArr: any = [];
  headervis = false;
  acceptanceArr: any = [];
  largeMontharr: any = [];
  acceptArr: any = [];
  uploadsection = false;
  willDownload = false;
  datebar = false;
  uniqueArray: any = [];
  datesize: any;
  effArr: any = [];
  dataString: any;
  dateString: any;
  value: number = 0;
  deliverable: any = [];
  filename: any;
  datefilename: any;
  mouseArr: any = [];
  totalmonthArr: any = [];
  size: any;
  dragAreaClass: any;
  draggedFiles: any;
  progressbar = false;
  uploadarea = false;
  disbledwn: any;
  uploaddatearea = false;
  filetype: any;
  disblebtn = true;
  disbledatebtn = true;
  totaleffort: any;
  sheet: any;
  checkyear: any = [];
  eventname: any;
  finaleffort: any;
  endholiday: any = [];
  effortVal: any = [];
  totalFteArr: any = [];
  checkVal: any = [];
  totalholiday: any = [];
  resultdata: any = [];
  idVal: any;
  bufferValue: any;
  dwnbtn = false;
  datavis = false;
  fileuploading = true;
  dateuploading = true;
  temp: any;
  holidayArr: any = [];
  dateArr: any = [];
  resultArr: any = [];
  monthData: any = [];
  monthDateData: any = [];
  totalmonthDateData: any = [];
  ftecal: any = [];
  bufferArr: any = [];
  deliverableVal: any = [];
  monthArr: any = [];
  qAValue: any;
  acceptValue: any = 0;
  errAccept = false;
  dateholiday: any = [];
  milestoneArr: any = [];
  deliverableArr: any = [];
  serialnoArr: any = [];
  serialtwoArr: any = [];
  holidayDate: any = [];
  holidayoneDate: any = [];
  holidayname: any = [];
  totaldataStringarr: any = [];
  tablevisone = false;
  tablevissix = false;
  tablevisfive = false;
  tablevisfour = false;
  tablevisthree = false;
  tablevistwo = false;
  tablevisseven = false;
  public btnStyle: any;
  public btnStyle1: any;
  public btnStyle2: any;
  public btnStyle3: any;
  public btnStyle4: any;
  public btnStyle5: any;
  public btnStyle6: any;
  public btnStyle7: any;
  checkBoxvalue = true;
  public weekStartdate = false;
  public weekEnddate = false;
  public weekAcceptdate = false;
  public showQAform = false;
  monthArrone: any = [];
  AcceptDateArr: any = [];
  effort: any;
  constructor(private toast: NgToastService, private http: HttpClient) {
  }

  ngOnInit() {
    this.getholiday();
    let startdate = new Date('2023-03-10');
    let endDate = new Date('2024-02-20');
    this.btnStyle = 'box';
    this.btnStyle1 = 'abutton';
    this.btnStyle2 = 'abutton';
    this.btnStyle3 = 'abutton';
    this.btnStyle4 = 'abutton';
    this.btnStyle5 = 'abutton';
    this.btnStyle6 = 'abutton';
    this.btnStyle7 = 'abutton';
    this.stylebtn = 'calan';
    this.serialnoArr = ['1', '2', '3', '4', '5', '6', '7', '8', '9'];
    this.holidayDate = ['01-01-2023', '26-01-2023', '07-03-2023', '15-08-2023', '09-07-2023', '02-10-2023', '13-11-2023', '26-11-2023', '25-12-2023']
    this.holidayoneDate = ['01-01-2024', '26-01-2024', '07-03-2024', '15-08-2024', '09-07-2024', '02-10-2024', '13-11-2024', '26-11-2024', '25-12-2024']
    this.holidayname = ['New Year Day', "Republic Day", "US Independence Day", "Memorial Day", "Independence Day", "Gandhi Jayanthi", "Deepavali Eve", "Day After Thanks Giving", "Christmas"]
  }
  chanebtn() {
    this.stylebtn = 'calanClick';
  }
  checkbox() {
    var checkbox = (<HTMLInputElement>document.getElementById('checkval')).checked;

    if (checkbox) {
      this.showQAform = true;

    }
    else {
      this.showQAform = false;

    }

  }
  openPDF(): void {
    let DATA: any = document.getElementById('htmlData');
    DATA.style.display = 'block';
    html2canvas(DATA).then((canvas) => {
      let fileWidth = 208; 
      let fileHeight = (canvas.height * fileWidth) / canvas.width;
      const FILEURI = canvas.toDataURL('image/png');
      let PDF = new jsPDF('p', 'mm', 'a4');
      let position = 0;

      PDF.addImage(FILEURI, 'PNG', 0, position, fileWidth, fileHeight);
      const downloadLink: any = document.createElement('a');
      downloadLink.href = PDF.output('bloburl');
      downloadLink.download = 'User_Manual.pdf';
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);

      DATA.style.display = 'none';
    });
  }

  getyear(year: any) {
    this.stylebtn = 'calan';

    for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
      if (this.holidayArr[0].Sheet1[j]) {
        let eventYear = new Date(this.holidayArr[0].Sheet1[j].Date).toLocaleString('default', { year: 'numeric' })
        if (eventYear === year) {
          this.eventArr.push(this.holidayArr[0].Sheet1[j]);
        }
      }
    }

    setTimeout(() => {
      this.exportholiday(year)
    }, 3000);
  }
  isstartDate(date: Date) {
    let dayOfWeek = date.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      this.weekStartdate = true;
    }
    else {
      this.weekStartdate = false;

    }
  }
  isendDate(date: Date) {
    let dayOfWeek = date.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      this.weekEnddate = true;
    }
    else {
      this.weekEnddate = false;

    }
  }
  isAcceptDate(date: Date) {
    const dayOfWeek = date.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      this.weekAcceptdate = true;
    }
    else {
      this.weekAcceptdate = false;

    }
  }


  dateBar(ev: any) {
    this.datebar = true;
    let interval = setInterval(() => {
      var elem = <HTMLElement>document.getElementById("mydateBar");
      this.value = this.value + Math.floor(Math.random() * 60) + 1;
      elem.style.width = this.value + "%";
      if (this.value >= 100) {
        this.value = 100;
        clearInterval(interval);
        this.datebar = false;
        this.disbledatebtn = false;

        this.ondateChange(ev);

      }
    }, 500);

  }
  uploadbar() {
    this.progressbar = true;

    let interval = setInterval(() => {
      var elem = <HTMLElement>document.getElementById("myBar");
      this.value = this.value + Math.floor(Math.random() * 60) + 1;
      elem.style.width = this.value + "%";
      if (this.value >= 100) {
        this.value = 100;
        clearInterval(interval);
        this.uploadsection = true;
        this.datafromexcel = false;
      }
    }, 500);

  }
  onFileChange(ev: any) {
    this.disbledatebtn = false;
    this.eventname = ev;
    let file = ev.target.files[0];
    this.filename = file.name;
    this.size = (file.size) / 1024
    this.filetype = file.type;
    this.checkfile(this.filename);
    this.sizefile(ev);
    this.getholiday();


  }
  async dateName(ev: any) {
    let file = ev.target.files[0];
    this.filename = file.name;

    let workbook = new Workbook();
    await workbook.xlsx.load(file);
    let sheetName: any;
    workbook.eachSheet((worksheet, sheetId) => {
      if (!worksheet.state || worksheet.state === 'visible') {
        sheetName = worksheet.name;
      }
    });
    if (sheetName != "Sheet1") {
      this.toast.error({ detail: "Error", summary: "Sheet Name should be Sheet1", position: 'br', duration: 5000 });
      this.datebar = false;
    }

    else {
      this.dateBar(ev);

    }
  }

  async namefile(ev: any) {
    let file = ev.target.files[0];
    this.filename = file.name;

    let workbook = new Workbook();
    await workbook.xlsx.load(file);
    let sheetName: any;
    workbook.eachSheet((worksheet, sheetId) => {
      if (!worksheet.state || worksheet.state === 'visible') {
        sheetName = worksheet.name;
      }
    });
    if (sheetName != "Sheet1") {
      this.toast.error({ detail: "Error", summary: "Sheet Name should be Sheet1", position: 'br', duration: 5000 });
      this.progressbar = false;
      this.uploadarea = false;
    }

    else {
      this.disblebtn = false;
      this.uploadbar();

    }
  }
  async sizefile(ev: any) {
    let file = ev.target.files[0];
    this.filename = file.name;

    const workbook = new Workbook();
    await workbook.xlsx.load(file);

    let presentSheetCount = 0;
    workbook.eachSheet((worksheet, sheetId) => {
      if (!worksheet.state || worksheet.state === 'visible') {
        presentSheetCount++;
      }
    });
    if (presentSheetCount > 1) {
      this.toast.error({ detail: "Error", summary: "Only one sheet allow", position: 'br', duration: 5000 });
      this.progressbar = false;
      this.uploadarea = false;
    }
    else {
      this.namefile(ev);
    }


  }
  async dateSize(ev: any) {
    let file = ev.target.files[0];
    const workbook = new Workbook();
    await workbook.xlsx.load(file);

    let presentSheetCount = 0;
    workbook.eachSheet((worksheet, sheetId) => {
      if (!worksheet.state || worksheet.state === 'visible') {
        presentSheetCount++;
      }
    });
    if (presentSheetCount > 1) {
      this.toast.error({ detail: "Error", summary: "Holiady file should have only one allow", position: 'br', duration: 5000 });
      this.datebar = false;
    }
    else {
      this.dateName(ev);
    }


  }

  checkValidation(arrdata: any) {
    let arrVal: any = [];
    let sumtotal: any;
    let totalArrval: any = [];
    var check = arrdata.map((b: any, index: any) => {
      this.isstartDate(b.StartDate);
      this.isendDate(b.EndDate);
      this.isAcceptDate(b.AcceptanceDate);
      let myTotalArr: any = [];
      if (!b.Phase || !b.Deliverable || !b.Activity || !b.StartDate || !b.EndDate || !b.Effort || !b.AcceptanceDate || b.StartDate > b.EndDate || b.EndDate > b.AcceptanceDate || this.weekStartdate || this.weekEnddate || this.weekAcceptdate) {
        arrVal.push(index + 1);
        for (let i = 0; i < arrVal.length; i++) {
          let sum = 1;
          sumtotal = sum + arrVal[i];
          totalArrval.push(sumtotal);
          myTotalArr = [...new Set(totalArrval)]
        }
        this.toast.error({ detail: "Error", summary: "In valid data from No:" + myTotalArr, duration: 10000 });

      }
    })
    if (totalArrval.length === 0) {
      this.toast.success({ detail: "Success", duration: 2000 });
      this.dwnbtn = true;

    }
    else {
      this.dwnbtn = false;
    }

  }

  checkdate(file: any) {


    let fileName = file.split(".")[1];
    if (fileName != 'xlsx') {
      this.toast.error({ detail: "Error", summary: "Holiady file should be Excel File", position: 'br', duration: 5000 });
      this.datebar = false;
    }

  }
  checkfile(file: any) {

    let fileName = file.split(".")[1];
    if (fileName != 'xlsx') {
      this.toast.error({ detail: "Error", summary: "Only Allow Excel File", position: 'br', duration: 5000 });
      this.progressbar = false;
      this.uploadarea = false;

    }

  }
  tableview() {
    this.datavis = true;
    this.tablevisone = false;
    this.tablevissix = false;
    this.tablevisfive = false;
    this.tablevisfour = false;
    this.tablevisthree = false;
    this.tablevistwo = false;
    this.tablevisseven = false;
    this.btnStyle = 'box';
    this.btnStyle1 = 'abutton';
    this.btnStyle2 = 'abutton';
    this.btnStyle3 = 'abutton';
    this.btnStyle4 = 'abutton';
    this.btnStyle5 = 'abutton';
    this.btnStyle6 = 'abutton';
    this.btnStyle7 = 'abutton';

  }
  tableviewone() {
    this.datavis = false;
    this.tablevisone = true;
    this.tablevissix = false;
    this.tablevisfive = false;
    this.tablevisfour = false;
    this.tablevisthree = false;
    this.tablevistwo = false;
    this.tablevisseven = false;
    this.btnStyle = 'abutton';
    this.btnStyle1 = 'box';
    this.btnStyle2 = 'abutton';
    this.btnStyle3 = 'abutton';
    this.btnStyle4 = 'abutton';
    this.btnStyle5 = 'abutton';
    this.btnStyle6 = 'abutton';
    this.btnStyle7 = 'abutton';

  }
  tableviewtwo() {
    this.datavis = false;
    this.tablevisone = false;
    this.tablevissix = false;
    this.tablevisfive = false;
    this.tablevisfour = false;
    this.tablevisthree = false;
    this.tablevistwo = true;
    this.tablevisseven = false;
    this.btnStyle = 'abutton';
    this.btnStyle1 = 'abutton';
    this.btnStyle2 = 'box';
    this.btnStyle3 = 'abutton';
    this.btnStyle4 = 'abutton';
    this.btnStyle5 = 'abutton';
    this.btnStyle6 = 'abutton';
    this.btnStyle7 = 'abutton';


  }
  tableviewthree() {
    this.datavis = false;
    this.tablevisone = false;
    this.tablevissix = false;
    this.tablevisfive = false;
    this.tablevisfour = false;
    this.tablevisthree = true;
    this.tablevistwo = false;
    this.tablevisseven = false;
    this.btnStyle = 'abutton';
    this.btnStyle1 = 'abutton';
    this.btnStyle2 = 'abutton';
    this.btnStyle3 = 'box';
    this.btnStyle4 = 'abutton';
    this.btnStyle5 = 'abutton';
    this.btnStyle6 = 'abutton';
    this.btnStyle7 = 'abutton';

  }
  tableviewfour() {
    this.datavis = false;
    this.tablevisone = false;
    this.tablevissix = false;
    this.tablevisfive = false;
    this.tablevisfour = true;
    this.tablevisthree = false;
    this.tablevistwo = false;
    this.tablevisseven = false;
    this.btnStyle = 'abutton';
    this.btnStyle1 = 'abutton';
    this.btnStyle2 = 'abutton';
    this.btnStyle3 = 'abutton';
    this.btnStyle4 = 'box';
    this.btnStyle5 = 'abutton';
    this.btnStyle6 = 'abutton';
    this.btnStyle7 = 'abutton';

  }
  tableviewfive() {
    this.datavis = false;
    this.tablevisone = false;
    this.tablevissix = false;
    this.tablevisfive = true;
    this.tablevisfour = false;
    this.tablevisthree = false;
    this.tablevistwo = false;
    this.tablevisseven = false;
    this.btnStyle = 'abutton';
    this.btnStyle1 = 'abutton';
    this.btnStyle2 = 'abutton';
    this.btnStyle3 = 'abutton';
    this.btnStyle4 = 'abutton';
    this.btnStyle5 = 'box';
    this.btnStyle6 = 'abutton';
    this.btnStyle7 = 'abutton';

  }
  tableviewsix() {
    this.datavis = false;
    this.tablevisone = false;
    this.tablevissix = true;
    this.tablevisfive = false;
    this.tablevisfour = false;
    this.tablevisthree = false;
    this.tablevistwo = false;
    this.tablevisseven = false;
    this.btnStyle = 'abutton';
    this.btnStyle1 = 'abutton';
    this.btnStyle2 = 'abutton';
    this.btnStyle3 = 'abutton';
    this.btnStyle4 = 'abutton';
    this.btnStyle5 = 'abutton';
    this.btnStyle6 = 'box';
    this.btnStyle7 = 'abutton';

  }
  tableviewseven() {
    this.datavis = false;
    this.tablevisone = false;
    this.tablevissix = false;
    this.tablevisseven = true;
    this.tablevisfive = false;
    this.tablevisfour = false;
    this.tablevisthree = false;
    this.tablevistwo = false;
    this.btnStyle = 'abutton';
    this.btnStyle1 = 'abutton';
    this.btnStyle2 = 'abutton';
    this.btnStyle3 = 'abutton';
    this.btnStyle4 = 'abutton';
    this.btnStyle5 = 'abutton';
    this.btnStyle6 = 'abutton';
    this.btnStyle7 = 'box';

  }
  backTopage() {
    location.reload();

  }
  checkAccpetRange() {
    let accept = <HTMLInputElement>document.getElementById("acceptValue");
    if (accept == null) {
      this.acceptValue = 0
    }
    else {
      this.acceptValue = accept.value;

    }
    if (this.acceptValue > 5 || this.acceptValue < 3) {
      this.errAccept = true;
    }
    else {
      this.errAccept = false;

    }
  }
  generate() {
    this.showQAform = false;
    this.checkBoxvalue = false;
    let qa = <HTMLInputElement>document.getElementById("qavalue")
    if (qa == null) {
      this.qAValue = 0
    }
    else {
      this.qAValue = qa.value;

    }
    let buffer = <HTMLInputElement>document.getElementById("buffervalue")
    if (buffer == null) {
      this.bufferValue = 0
    }
    else {
      this.bufferValue = buffer.value;
    }
    let progressBtn = <HTMLElement>document.getElementById('progressBtn');
    let classbtn = progressBtn?.className;

    if (!progressBtn?.className.includes("active")) {
      let classbtnName = classbtn + " active";
      (<HTMLElement>document.getElementById('progressBtn')).className = classbtnName;
      setTimeout(() => {
        (<HTMLElement>document.getElementById('progressBtn')).className = classbtn;
        this.generatevalue();
        this.headervis = true;
        this.generatebtn = false;
      }, 10000);

    }

  }
  downloadbar() {
    let progressBtn = <HTMLElement>document.getElementById('progressBtnone');
    let classbtn = progressBtn?.className;

    if (!progressBtn?.className.includes("active")) {
      let classbtnName = classbtn + " active";
      (<HTMLElement>document.getElementById('progressBtnone')).className = classbtnName;
      setTimeout(() => {
        (<HTMLElement>document.getElementById('progressBtnone')).className = classbtn;
        this.exportdata();
      }, 100);

    }

  }
  generatevalue() {
    let workBook: any;
    let jsonData = null;
    let jsonDataOne: any;
    const reader = new FileReader();
    const file = this.eventname.target.files[0];
    reader.onload = (event) => {
      const data = reader.result;
      workBook = XLSX.read(data, { type: 'binary', cellDates: true });
      jsonData = workBook.SheetNames.reduce((initial: any, name: any) => {
        const sheet = workBook.Sheets[name];
        initial[name] = XLSX.utils.sheet_to_json(sheet);
        jsonDataOne = initial[name].map((row: any) => {
          let startfield = 'StartDate'
          let endfield = 'EndDate'
          let accept = 'AcceptanceDate'
          let startdate = new Date(row[startfield])
          let enddate = new Date(row[endfield])
          let acceptDate = new Date(row[accept])
          let startValue = startdate.setDate(startdate.getDate() + 1);
          let endValue = enddate.setDate(enddate.getDate() + 1);
          let acceptValue = acceptDate.setDate(acceptDate.getDate() + 1);
          let newstartdate = new Date(startValue);
          let newenddate = new Date(endValue);
          let newacceptdate = new Date(acceptValue);
          return { ...row, [startfield]: newstartdate, [endfield]: newenddate, [accept]: newacceptdate };
        });
        return jsonDataOne;

      }, {});

      this.dataString = jsonDataOne;
      this.checkValidation(this.dataString);
      this.getmonthwise(this.dataString);
      this.calculatedeliverable(this.dataString);   
    }
    reader.readAsBinaryString(file);
    this.datavis = true;
  }
  ondate(ev: any) {
    let file = ev.target.files[0];
    let fileName = file.name;
    this.checkdate(fileName);
    this.dateSize(ev);

  }
  ondateChange(ev: any) {

    let hoildayTotal = localStorage.getItem('HolidayList');
    let newArr = JSON.stringify(hoildayTotal);

    if (newArr.length >= 1) {
      localStorage.removeItem('HolidayList');
      this.writejson(ev);
      this.getholiday();
    }

    else {
      this.writejson(ev);
    }




  }

  writejson(ev: any): void {
    this.dateBar(ev);
    let workBook: any;
    let jsonData = null;
    let jsonDataOne: any;
    let initial: any
    let SheetName: any = [];
    const reader = new FileReader();
    const file = ev.target.files[0];
    reader.onload = (event) => {
      const data = reader.result;
      workBook = XLSX.read(data, { type: 'binary', cellDates: true });
      SheetName = workBook.SheetNames;
      if (SheetName.length > 1) {
        this.toast.warning({ detail: "Error", summary: "Only one sheet allow", position: 'tr', duration: 5000 });
      }
      jsonData = workBook.SheetNames.reduce((initial: any, name: any) => {
        this.sheet = workBook.Sheets[name];
        initial[name] = XLSX.utils.sheet_to_json(this.sheet);
        jsonDataOne = initial[name].map((row: any) => {
          let datefield = 'Date'
          let dateVal = new Date(row[datefield])
          let dateValue = dateVal.setDate(dateVal.getDate() + 1);
          let finalDate = new Date(dateValue)
          return { ...row, [datefield]: finalDate };

        });
        return jsonDataOne;

      }, {});

      let sheetval = SheetName[0];
      this.dateString = jsonDataOne;
      let obj = {
        id: '1',
        [sheetval]: this.dateString
      }
      let hoildaylistArr: any = [];
      hoildaylistArr.push(obj);
      setTimeout(() => {
        localStorage.setItem('HolidayList', JSON.stringify(hoildaylistArr));
        let newItem = localStorage.getItem('HolidayList')
        if (newItem) {
          this.datebar = false;
          this.toast.success({ detail: "Hoilday File upload sucessfully", position: 'br', duration: 10000 });

          setTimeout(() => {

            if (newItem) {
              location.reload();
            }

          }, 2000);


        }

      }, 2000);
    }
    reader.readAsBinaryString(file);
  }
  getholiyear() {
    if (this.holidayArr !== null) {
      for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
        if (this.holidayArr[0].Sheet1[j]) {
          this.mouseArr.push(new Date(this.holidayArr[0].Sheet1[j].Date).toLocaleString('default', { year: 'numeric' }))
        }
      }
      this.uniqueArray = [...new Set(this.mouseArr)];
    }

  }
  mouseEnter() {

    if (this.holidayArr.length == 0) {
      this.toast.error({ detail: "Error", summary: "Hoilday list not availble kindly upload it  ", position: 'br', duration: 5000 });
    }
    else {
      for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
        if (this.holidayArr[0].Sheet1[j]) {
          this.mouseArr.push(new Date(this.holidayArr[0].Sheet1[j].Date).toLocaleString('default', { year: 'numeric' }))
        }
      }
      let unique = [...new Set(this.mouseArr)];
      if (unique) {
        this.toast.warning({ detail: "I have Holiday list of   " + unique, position: 'br', duration: 5000 });
      }
    }
  }
  getholiday() {
    let item = localStorage.getItem('HolidayList')
    let myArray: any = [];
    if (item !== null) {
      myArray = JSON.parse(item);
    }
    this.holidayArr = myArray;
    if (this.holidayArr.length != 0) {
      this.getholiyear();
    }
  }

  formatDate(date: any) {
    let stringdate = date.toString();
    return stringdate;
  }
  getholidayleave(start: any, end: any) {
    let startDate = new Date(start);
    let endDate = new Date(end);
    let count = 0;
    const curDate = new Date(startDate.getTime());
    let i = 0;
    for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
      if (this.holidayArr[0].Sheet1[j]) {
        this.dateArr.push(new Date(this.holidayArr[0].Sheet1[j].Date).toLocaleDateString('en-us'))
      }
    }

    while (curDate <= endDate) {
      let holiday = new Date(curDate).toLocaleDateString('en-us');

      if (this.dateArr.includes(holiday)) count++;

      curDate.setDate(curDate.getDate() + 1);

    }

    return count;

  }


  getmonthwise(data: any) {
    this.calculatemonthwise(data);
    this.calculatemonthendwise(data)
    let startArr: any = [];
    let endArr: any = [];
    let sortStartArr: any = [];
    let sortEndArr: any = [];
    data.forEach(function (e: any) {
      let startDate = e.StartDate

      let endDate = e.EndDate
      startArr.push(startDate);
      endArr.push(endDate);

    })
    let asscendingorder = startArr.sort((a: any, b: any) => {

      let start: any = new Date(a);
      let end: any = new Date(b)
      return start - end


    })
    let descendingoreder = endArr.sort((a: any, b: any) => {
      let start: any = new Date(a);
      let end: any = new Date(b)
      return end - start

    })
    this.getMonthsBetweenDates(asscendingorder[0], descendingoreder[0])

  }

  getMonthsBetweenDates(startDate: any, end: any) {
    let current = new Date(startDate);

    let endMonth = new Date(end)
    let datedata = current.toLocaleString('default', { month: 'short', year: 'numeric' });
    this.monthData.push(datedata.toString());
    while (current.getMonth() !== end.getMonth() || current.getFullYear() !== end.getFullYear()) {
      current.setMonth(current.getMonth() + 1);
      let datedata = current.toLocaleString('default', { month: 'short', year: 'numeric' });
      this.monthData.push(datedata.toString());
      endMonth.setMonth(endMonth.getMonth() + 1);

    }

  }
  countWeekdays(startDate: any, endDate: any) {
    let leaveArr: any = [];
    let monthCount: any = {};
    let start = startDate;
    let end = endDate;
    for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
      if (this.holidayArr[0].Sheet1[j]) {
        leaveArr.push(new Date(this.holidayArr[0].Sheet1[j].Date).toLocaleDateString('en-us'));
      }
    }
    var currentDate = new Date(start);
    let count: any;
    let month1: any;
    while (currentDate <= end) {
      var month = currentDate.getMonth();
      var day = currentDate.getDay();
      var currentDateString = currentDate.toLocaleDateString('en-us');
      const date1 = new Date();
      date1.setMonth(month);
      let monththth = date1.toLocaleString([], { month: 'short' })
      month1 = monththth + ' ' + date1.getFullYear();


      if (day !== 0 && day !== 6 && !leaveArr.includes(currentDateString)) {
        if (!monthCount[month]) {
          monthCount[month] = 0;
        }
        count = monthCount[month]++;
      }
      currentDate.setDate(currentDate.getDate() + 1);

    }
    return count



  }

  calculatemonthwise(data: any) {
    this.checkmilestone(data);
    let newArr: any = [];
    let workingWeekDays: any = {};
    data.map((e: any) => {
      let StartDate = new Date(e.StartDate);
      let startformat = moment(StartDate).format('YYYY-MM-DD');
      let EndDate = new Date(e.EndDate);
      let endformat = moment(EndDate).format('YYYY-MM-DD');
      let deliverable = e.Deliverable;
      for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
        if (this.holidayArr[0].Sheet1[j]) {
          this.totalholiday.push(moment(this.holidayArr[0].Sheet1[j].Date).format('YYYY-MM-DD'));
        }
      }

      let currentDate = moment(startformat);
      while (currentDate.isBefore(endformat)) {
        let currentMonth = currentDate.format('MMM YYYY');
        if (!this.totalholiday.includes(currentDate.format('YYYY-MM-DD')) && currentDate.day() !== 0 && currentDate.day() !== 6) {
          if (!workingWeekDays[currentMonth]) {
            workingWeekDays[currentMonth] = 0;
          }
          workingWeekDays[currentMonth]++;

        }
        currentDate.add(1, 'days');

      }

    })

    newArr.push(workingWeekDays);
    setTimeout(() => {
      const keys = Object.keys(newArr[0]).map((key) => key);

      this.monthData.map((monn: any) => {
        if (keys.includes(monn)) {

          this.effArr[monn] = newArr[0][monn]
        }
        else {
          this.effArr[monn] = 0
        }

      })
      this.effortVal = Object.values(this.effArr).map((value) => value);
      let sum = 0;
      for (let i = 0; i < this.effortVal.length; i++) {
        if (this.qAValue) {
          this.ftecal.push(this.qAValue);

        }
        else {
          let j = 0;
          this.ftecal.push(j);

        }
        if (this.bufferValue) {
          this.bufferArr.push(this.bufferValue)
        }
        else {
          let j = 0;
          this.bufferArr.push(j);

        }
        sum = (this.effortVal[i] * 8) + sum;

      }

      this.finaleffort = sum;
      this.totoalBuffer();
      this.totalQaValue();
      this.totalFte();
    }, 2000);


  }

  totalFte() {
    let sum: any;
    for (let i = 0; i < this.effortVal.length; i++) {
      if (this.bufferArr[i] === 0 && this.ftecal[i] === 0) {
        let effort = Number((this.effortVal[i] * 8) / 168);
        sum = effort
        this.totalFteArr.push(sum)

      }
      else if (this.bufferArr[i] === 0) {
        let effort = Number((this.effortVal[i] * 8) / 168)
        let QA = Number(this.ftecal[i] / 100)
        sum = effort + QA
        this.totalFteArr.push(sum)

      }
      else if (this.ftecal[i] === 0) {
        let effort = Number((this.effortVal[i] * 8) / 168)
        let buffer = Number(this.bufferArr[i] / 100)
        sum = effort + buffer
        this.totalFteArr.push(sum)
      }
      else {
        let effort = Number((this.effortVal[i] * 8) / 168)
        let buffer = Number(this.bufferArr[i] / 100)
        let QA = Number(this.ftecal[i] / 100)
        sum = effort + buffer + QA


        this.totalFteArr.push(sum)
      }
      this.totalfinalFte()
    }
  }
  totalfinalFte() {
    let sum = 0;
    for (let i = 0; i < this.totalFteArr.length; i++) {

      sum = this.totalFteArr[i] + sum;

    }
    this.totalFtevalue = sum;

  }
  totoalBuffer() {
    let sum = 0;
    for (let i = 0; i < this.bufferArr.length; i++) {
      if (this.bufferArr[i] === 0) {
        sum = 0;
      }
      else {
        sum = Number(this.bufferArr[i] / 100) + sum;
      }
    }
    this.totalBuffer = sum;
  }
  totalQaValue() {
    let sum = 0;
    for (let i = 0; i < this.ftecal.length; i++) {
      if (this.ftecal[i] === 0) {
        sum = 0;
      }
      else {
        sum = Number(this.ftecal[i] / 100) + sum;
      }
    }
    this.totalQAvalue = sum;
  }
  calaculateeffort() {
    this.monthArrone = [];
    this.totaldataStringarr = this.dataString
    this.totaldataStringarr.map((e: any, ind: any) => {
      let StartDate = new Date(e.StartDate);
      let startformat = moment(StartDate).format('YYYY-MM-DD');
      let EndDate = new Date(e.EndDate);
      let endformat = moment(EndDate).format('YYYY-MM-DD');


      let totalval = this.calculatemonthwise1(StartDate, EndDate);
      let effortVal = this.getSum(totalval)
      let newrr: any = [];
      this.monthData.map((mon: any, index: any) => {

        this.totaldataStringarr[ind][mon] = totalval[index];
        this.totaldataStringarr[ind]['Effort'] = effortVal;

      })



    })

  }
  calculatemonthwise1(startDate: any, endDate: any) {
    this.monthArr = [];

    let StartDate = new Date(startDate);
    let startformat = moment(StartDate).format('YYYY-MM-DD');
    let EndDate = new Date(endDate);
    let endformat = moment(EndDate).format('YYYY-MM-DD');




    let startMonth = this.getMonthDetails(startDate.getMonth());
    let endMonth = this.getMonthDetails(endDate.getMonth());




    let startMonthwithYear = startMonth + " " + startDate.getFullYear();
    let endMonthwithYear = endMonth + " " + endDate.getFullYear();
    let datecalculs = this.datecalculator(startDate, endDate);
    let monthKeys = Object.keys(datecalculs);
    let startRemaingDays = datecalculs[startMonthwithYear];
    let endRemaingDays = datecalculs[endMonthwithYear];


    if (startDate.getMonth() === endDate.getMonth() && startDate.getFullYear() === endDate.getFullYear()) {
      startRemaingDays = this.getDaysBetween(startDate.getFullYear(), startDate.getMonth(), startDate) + this.getDaysBetween(endDate.getFullYear(), endDate.getMonth(), endDate);
    }
    let starandEndArr: any = [];



    starandEndArr[startMonthwithYear] = startRemaingDays;
    starandEndArr[endMonthwithYear] = endRemaingDays;
    let startAndEndMonths: any = [];
    startAndEndMonths.push(startMonthwithYear);
    startAndEndMonths.push(endMonthwithYear);


    if (this.monthData.length) {
      this.monthData.map((mon: any) => {
        let month = mon.split(" ");
        let monthnumber = this.getShortMonthNumber(month[0]);
        if (startAndEndMonths.includes(mon)) {
          let days = starandEndArr[mon];
          this.monthArr.push(days * 8)
        }
        else {
          if (monthKeys.includes(mon)) {

            const date = new Date(month[1], monthnumber, 1);
            let daysInMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();

            var start = new Date(date.getFullYear(), date.getMonth(), 1);
            var end = new Date(date.getFullYear(), date.getMonth() + 1, 0);

            let datecalculation = this.datecalculator(start, end);
            let monWithYear = this.getMonthDetails(monthnumber) + " " + date.getFullYear();
            this.monthArr.push(datecalculation[monWithYear] * 8);

          }
          else {
            this.monthArr.push(0);

          }
        }

      })
    }

    this.effort = this.getSum(this.monthArr);
    return this.monthArr
  }
  getSum(numbers: number[]): number {
    return numbers.reduce((sum, number) => sum + number, 0);
  }
  getMonthDetails(monthNumber: any) {
    let date = new Date();
    date.setMonth(monthNumber);
    let month = date.toLocaleString('default', { month: 'short' });
    return month;
  }

  getDaysBetween(year: any, month: any, date: any) {
    let nextMonthStart = new Date(year, month + 1, 1);
    let remainingDays = (nextMonthStart.getTime() - date.getTime()) / (1000 * 60 * 60 * 24);

    return Math.round(remainingDays);

  }

  getShortMonthNumber(shortMonth: string): number {
    const shortMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return shortMonths.indexOf(shortMonth);
  }

  mycalculation(start: any, end: any) {
    let resultarr: any = [];
    const result: any = {};
    let current = new Date(start);
    const endDate = new Date(end);
    for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
      if (this.holidayArr[0].Sheet1[j]) {
        this.inbetweenholiday.push(moment(this.holidayArr[0].Sheet1[j].Date).format('YYYY-MM-DD'));
      }
    }
    while (current < endDate) {
      if (current.getDay() !== 0 && current.getDay() !== 6 && !this.inbetweenholiday.includes(current.toLocaleDateString())) {
        const month = current.toLocaleString('default', { month: 'short' });
        const year = current.getFullYear();
        if (!result[`${month} ${year}`]) {
          result[`${month} ${year}`] = 0;
        }
        result[`${month} ${year}`]++;
      }
      current.setDate(current.getDate() + 1);
    }
    resultarr.push(result);
  }

  datecalculator(start: any, end: any) {
    let StartDate = new Date(start);
    let startformat = moment(start).format('YYYY-MM-DD');
    let EndDate = new Date(end);
    let endformat = moment(end).format('YYYY-MM-DD');
    let currentDate = moment(startformat);
    let workingWeekDays: any = {};
    for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
      if (this.holidayArr[0].Sheet1[j]) {
        this.dateholiday.push(moment(this.holidayArr[0].Sheet1[j].Date).format('YYYY-MM-DD'));
      }
    }
    while (currentDate.isBefore(endformat)) {
      let currentMonth = currentDate.format('MMM YYYY');
      if (!this.totalholiday.includes(currentDate.format('YYYY-MM-DD')) && currentDate.day() !== 0 && currentDate.day() !== 6) {
        if (!workingWeekDays[currentMonth]) {
          workingWeekDays[currentMonth] = 0;
        }
        workingWeekDays[currentMonth]++;

      }
      currentDate.add(1, 'days');
    }
    return workingWeekDays;
  }

  calculatedeliverable(data: any) {
    let result: any = {};
    data.map((e: any) => {
      let start = new Date(e.StartDate);
      let end = new Date(e.EndDate);

      if (e.Deliverable === 'Y') {
        while (start <= end) {
          let month = start.toLocaleString('default', { month: 'short' });
          let year = start.getFullYear();
          let key = `${month} ${year}`;

          if (result[key]) {
            result[key]++;
          } else {
            result[key] = 1;
          }

          start.setMonth(start.getMonth() + 1);
        }
        if (start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear()) {
          let month = end.toLocaleString('default', { month: 'short' });
          let year = end.getFullYear();
          let key = `${month} ${year}`;
          if (result[key]) {
            result[key]++
          } else {
            result[key] = 1;
          }
        }
      }
    });
    this.deliverable.push(result);
    setTimeout(() => {
      const keys = Object.keys(this.deliverable[0]).map((key) => key);

      this.monthData.map((monn: any) => {
        if (keys.includes(monn)) {

          this.deliverableArr[monn] = this.deliverable[0][monn]
        }
        else {
          this.deliverableArr[monn] = 0
        }

      })
      this.deliverableVal = Object.values(this.deliverableArr).map((value) => value);

    }, 2000);
  }
  checkmilestone(data: any) {
    this.milestoneArr = data.sort((x: { EndDate: Date; }, y: { EndDate: Date; }) => <any>new Date(x.EndDate) - <any>new Date(y.EndDate))

    this.milestoneArr.map((e: any) => {
      this.enddateArr.push(e.EndDate);
    })

    let monthYearWiseLargest: any = {}
    this.enddateArr.forEach((date: any) => {
      const dateObject = new Date(date);
      const monthYear = dateObject.toLocaleDateString(undefined, { month: 'short', year: 'numeric' });
      if (!monthYearWiseLargest[monthYear] || new Date(monthYearWiseLargest[monthYear]) < dateObject) {
        monthYearWiseLargest[monthYear] = date;
      }
    });
    this.largeMontharr = Object.values(monthYearWiseLargest);
  }

  checkmilesttwo(data: any) {
    this.milestoneArr = data.sort((x: { EndDate: Date; }, y: { EndDate: Date; }) => <any>new Date(x.EndDate) - <any>new Date(y.EndDate))

    this.milestoneArr.map((e: any) => {
      this.enddateArr.push(e.EndDate);
    })

    let monthWiseLargest: any = {}
    this.enddateArr.forEach((date: any) => {
      const dateObject = new Date(date);
      const month = dateObject.getMonth();
      const monthYear = dateObject.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
      if (!monthWiseLargest[month] || new Date(monthWiseLargest[month]) < dateObject) {
        monthWiseLargest[month] = monthYear;
      }
    });
    this.largeMontharr = Object.values(monthWiseLargest);
  }
  calculatemilestone(date: any) {
    let enddate = date;
    let indexVal = this.largeMontharr.indexOf(enddate);
    return indexVal;


  }
  calculatemonthendwise(data: any) {
    let result: any = [];
    for (let j = 0; j <= this.holidayArr[0].Sheet1.length; j++) {
      if (this.holidayArr[0].Sheet1[j]) {
        this.endholiday.push(moment(this.holidayArr[0].Sheet1[j].Date).format('YYYY-MM-DD'));
      }
    }
    let totalDaysByMonth: any = {};

    data.forEach((item: any) => {
      let start = new Date(item.StartDate);
      let end = new Date(item.EndDate);
      let days = 0;
      while (start < end) {
        if (start.getDay() !== 0 && start.getDay() !== 6 && !this.endholiday.includes(start.toLocaleDateString())) {
          days++;
        }
        start.setDate(start.getDate() + 1);
      }
      let month = end.toLocaleString('default', { month: 'short' });
      let year = end.getFullYear();
      let key = `${month} ${year}`;
      if (!totalDaysByMonth[key]) totalDaysByMonth[key] = days;
      else totalDaysByMonth[key] += days;
    });
    for (let [month, count] of Object.entries(totalDaysByMonth)) {
      const monthYear = `${month}`
      result.push({ [monthYear]: count })
    }
    result.sort((a: any, b: any) => {
      var monthA = new Date(Object.keys(a)[0]).getMonth();
      var monthB = new Date(Object.keys(b)[0]).getMonth();
      return monthA - monthB
    });
    setTimeout(() => {

      this.endeffortArr = [];

      for (let i = 0; i < result.length; i++) {
        let objValues = Object.values(result[i]);
        this.endeffortArr.push(...objValues);
      }
    }, 200)


  }
  formatdate(month: any) {
    let dateString = month;
    let finalDate = dateString.replace(' ', '-');
    return finalDate;
  }

  refersh() {
    location.reload();
  }
  exportholiday(year: any) {
    const element = <HTMLElement>document.getElementById("holidaytableone");

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(element, {
      raw: true,
      dateNF: "MMM yyyy"
    });
    this.stylebtn = 'calan';
    let title: any = year + ".xlsx";
    let sheetname: any = year
    XLSX.utils.book_append_sheet(wb, ws, sheetname);
    XLSX.writeFile(wb, title);

  }
  exportdata() {
    this.calaculateeffort();

    const workBook = new Workbook();
    const workSheetzero = workBook.addWorksheet('RBWO');
    const workSheet = workBook.addWorksheet('Timeline');
    const worksheetFour = workBook.addWorksheet('PTE');
    let worksheetThree = workBook.addWorksheet('PPM');
    let workSheetone = workBook.addWorksheet('Milestones & Deliverables');
    let workSheetTwo = workBook.addWorksheet('Effort Spend');
    const worksheetFive = workBook.addWorksheet('Pricing and Payment Schedule');


    let headerNames = this.monthData;
    let headingArr: any = ['S.No', 'Phase', 'Deliverable', 'Activity', 'Start Date', 'End Date', 'Effort', 'Acceptance Date']
    let concatenatedArray = [...headingArr, ...headerNames];
    let header = workSheet.addRow(concatenatedArray);
    let headerone = workSheetzero.addRow(concatenatedArray);
    for (let j = 1; j <= concatenatedArray.length; j++) {
      let col = header.getCell(j);
      col.font = {
        bold: true,
        size: 11,
        name
          : 'Arial'
      };
      col.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    for (let j = 1; j <= concatenatedArray.length; j++) {
      let col = headerone.getCell(j);
      col.font = {
        bold: true,
        size: 11,
        name: 'Arial'
      };
      col.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }

    let options: any = {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric'
    }

    let dataArr: any = [];


    this.totaldataStringarr.forEach((data: any, index: any) => {
      let start = new Date(data.StartDate);
      let end = new Date(data.EndDate);
      let accept = new Date(data.AcceptanceDate);

      concatenatedArray.map((conArr) => {
        let varr = data[conArr];
        dataArr.push(varr);
      })

      let arr = this.datavalues(data, headerNames);



      const row = workSheet.addRow([index + 1, data.Phase, data.Deliverable, data.Activity, moment(start).format('MM/DD/YYYY'), moment(end).format('MM/DD/YYYY'), data.Effort, moment(accept).format('MM/DD/YYYY'), ...arr]);

      for (let i = 1; i <= headingArr.length; i++) {
        let col = row.getCell(i);
        col.font = {
          size: 11,
          name: 'Arial'
        };
        col.border = {
          top: { style: 'thin', color: { argb: '000000' } },
          left: { style: 'thin', color: { argb: '000000' } },
          bottom: { style: 'thin', color: { argb: '000000' } },
          right: { style: 'thin', color: { argb: '000000' } }
        };
      }
      for (let j = 1; j <= 2; j++) {
        let column = workSheet.getColumn(j);

        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          let length = cell.value ? cell.value.toString().length : 0;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength * 1.2;
      }
      for (let j = 1; j <= 3; j++) {
        let column = workSheet.getColumn(j);

        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          let length = cell.value ? cell.value.toString().length : 0;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength * 1.2;
      }
      let colum = workSheet.getColumn(4);
      colum.alignment = { wrapText: true };
      colum.width = 50;
      for (let k = 5; k <= headingArr.length; k++) {
        let column = workSheet.getColumn(k);

        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          let length = cell.value ? cell.value.toString().length : 0;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength * 1.2;
      }
      let color = this.getRandomLightColor();
      for (let i = 9; i <= concatenatedArray.length; i++) {
        const col = row.getCell(i);
        col.font = {
          size: 11,
          name: 'Arial'
        };
        col.fill = {

          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: data[headerNames[i - 9]] !== 0 ? color : ''
          }

        };
        let column = workSheet.getColumn(i);
        column.width = 12;
        col.border = {
          top: { style: 'thin', color: { argb: '000000' } },
          left: (data.Phase == 'Development' && data[headerNames[i - 9]] != 0) || (data.Phase == 'Testing' && data[headerNames[i - 9]] != 0) ? {} : { style: 'thin', color: { argb: '000000' } },
          bottom: { style: 'thin', color: { argb: '000000' } },
          right: { style: 'thin', color: { argb: '000000' } }
        };
        const val = row.getCell(i).value = '';
      }
    });
    this.totaldataStringarr.forEach((data: any, index: any) => {
      let start = new Date(data.StartDate);
      let end = new Date(data.EndDate);
      let accept = new Date(data.AcceptanceDate);

      concatenatedArray.map((conArr) => {
        let varr = data[conArr];
        dataArr.push(varr);
      })
      let objj = {

      }
      let arr = this.datavalues(data, headerNames);



      const row = workSheetzero.addRow([index + 1, data.Phase, data.Deliverable, data.Activity, moment(start).format('MM/DD/YYYY'), moment(end).format('MM/DD/YYYY'), data.Effort, moment(accept).format('MM/DD/YYYY'), ...arr]);

      for (let i = 1; i <= headingArr.length; i++) {
        let col = row.getCell(i);
        col.border = {
          top: { style: 'thin', color: { argb: '000000' } },
          left: { style: 'thin', color: { argb: '000000' } },
          bottom: { style: 'thin', color: { argb: '000000' } },
          right: { style: 'thin', color: { argb: '000000' } }
        };
        col.font = {
          size: 11,
          name: 'Arial'
        };
      }
      for (let j = 1; j <= 3; j++) {
        let column = workSheetzero.getColumn(j);

        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          let length = cell.value ? cell.value.toString().length : 0;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength * 1.2;
      }
      let colum = workSheetzero.getColumn(4);
      colum.alignment = { wrapText: true };
      colum.width = 50;
      for (let k = 5; k <= headingArr.length; k++) {
        let column = workSheetzero.getColumn(k);

        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          let length = cell.value ? cell.value.toString().length : 0;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength * 1.2;
      }

      for (let i = 9; i <= concatenatedArray.length; i++) {
        const col = row.getCell(i);

        col.border = {
          top: { style: 'thin', color: { argb: '000000' } },
          left: (data.Phase == 'Development' && data[headerNames[i - 9]] != 0) || (data.Phase == 'Testing' && data[headerNames[i - 9]] != 0) ? {} : { style: 'thin', color: { argb: '000000' } },
          bottom: { style: 'thin', color: { argb: '000000' } },
          right: { style: 'thin', color: { argb: '000000' } }
        };
        col.font = {
          size: 11,
          name: 'Arial'
        };
        let column = workSheetzero.getColumn(i);
        column.width = 12;
      }
    });

    let headerArr = ['ETS #', 'Deliverable/Milestone', 'Deliverable Due Date', 'Acceptance Due Date'];
    let headerTwo = workSheetone.addRow(headerArr);
    for (let j = 1; j <= headerArr.length; j++) {
      let col = headerTwo.getCell(j);
      col.font = {
        bold: true,
        size: 11,
        name: 'Arial'
      };
      col.border = {
        top: { style: 'thin' },
        left: {},
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

    }
    this.milestoneArr.forEach((data: any, index: any) => {
      let end = new Date(data.EndDate);
      let accept = new Date(data.AcceptanceDate);
      if (this.largeMontharr.includes(data.EndDate)) {
        const row = workSheetone.addRow([index + 1, data.Activity, moment(end).format('MM/DD/YYYY'), this.getaccptancedate(end, accept)]);
        if (this.milestoneArr[index + 1]?.EndDate !== end) {
          const rowone = workSheetone.addRow([null, 'Milestone ' + '   ' + (this.calculatemilestone(data.EndDate) + 1) + '-' + this.getactivity(data.EndDate), null, this.getTotalAccept(this.getaccptancedate(end, accept))]);
          workSheetone.mergeCells(`B${rowone.number}:C${rowone.number}`);
          for (let j = 1; j <= headerArr.length; j++) {
            let col = rowone.getCell(j);
            let columns = workSheetone.getColumn(2);
            columns.alignment = { wrapText: true };
            columns.width = 50;
            col.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '#D3D3D3' }
            };
            col.font = {
              size: 11,
              name: 'Arial',
              bold: true,
            };
            col.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
          }

          for (let j = 1; j <= headerArr.length; j++) {
            let coln = row.getCell(j);
            let column = workSheetone.getColumn(4);
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
              let length = cell.value ? cell.value.toString().length : 0;
              if (length > maxLength) {
                maxLength = length;
              }
            });

            column.width = maxLength * 1.2
            let colum = workSheetone.getColumn(3);
            colum.alignment = { wrapText: true };


            colum.width = 20;


            coln.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };

            coln.font = {
              size: 11,
              name: 'Arial'
            };

          }

        }
      } else {
        const row = workSheetone.addRow([index + 1, data.Activity, moment(end).format('MM/DD/YYYY'), this.getaccptancedate(end, accept)]);
        for (let j = 1; j <= headerArr.length; j++) {
          let col = row.getCell(j);

          col.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
          col.font = {
            size: 11,
            name: 'Arial'
          };
        }
      }

    })

    workSheetTwo.columns = [
      { header: '# Milestone', key: 'milestone', width: 15 },
      { header: 'Date', key: 'date', width: 13 },
      { header: 'Effortsspend', key: 'effort', width: 18 },
      { header: 'Effort%/Month', key: 'effortDays', width: 18 }
    ];

    for (let i = 0; i < this.largeMontharr.length; i++) {
      let end = new Date(this.largeMontharr[i]);

      let headerThree = workSheetTwo.addRow({
        milestone: '# Milestone ' + (i + 1),
        date: moment(end).format('MM/DD/YYYY'),
        effort: Number(this.endeffortArr[i] * 8),
        effortDays: Number((((this.endeffortArr[i]) / 168) * 100).toFixed(3))
      });
      const headerRow = workSheetTwo.getRow(1);
      let headerrow = workSheetTwo.getRow(i)
      headerrow.font = {
        size: 11,
        name: 'Arial'
      }
      headerRow.font = {
        bold: true, size: 11,
        name: 'Arial'
      };
      headerRow.eachCell(cell => {
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
      workSheetTwo.eachRow((headerThree, rowNumber) => {
        if (rowNumber !== 1) {
          headerThree.eachCell(cell => {
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
            cell.font = {
              size: 11,
              name: 'Arial'
            }
          });
        }
      });

    }
    const headerRow = ['Months', ...headerNames, 'Total'];
    const headerRowWithBorders = worksheetThree.addRow(headerRow);
    headerRowWithBorders.font = {
      bold: true, size: 11,
      name: 'Arial'
    };
    headerRowWithBorders.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };

    });

    const effortRow = ['Effort for Month', ...this.effortVal.map((effort: any) => Number(effort * 8)), this.finaleffort];
    const effortRowWithBorders = worksheetThree.addRow(effortRow);
    effortRowWithBorders.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.font = {
        size: 11,
        name: 'Arial'
      }
    });

    const fteForMonth = ['FTE for Month', ...this.effortVal.map((effort: any) => Number(((effort * 8) / 168).toFixed(3))), Number(((this.finaleffort) / 168).toFixed(3))];
    const fteForMonthWithBorders = worksheetThree.addRow(fteForMonth);
    fteForMonthWithBorders.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.font = {
        size: 11,
        name: 'Arial'
      }
    });
    const qaInFte = ['QA % (0) in FTE', ...this.ftecal, Number((this.totalQAvalue).toFixed(3))];
    const qaInFteWithBorders = worksheetThree.addRow(qaInFte);
    qaInFteWithBorders.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.font = {
        size: 11,
        name: 'Arial'
      }
    });
    const bufferInFte = ['Buffer % (0) in FTE', ...this.bufferArr, Number(this.totalBuffer.toFixed(3))];
    const bufferInFteWithBorders = worksheetThree.addRow(bufferInFte);
    bufferInFteWithBorders.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
      cell.font = {
        size: 11,
        name: 'Arial'
      }

    });

    const finalFte = ['Final FTE', ...this.totalFteArr.map((effort: any) => Number((effort).toFixed(3))), Number(this.totalFtevalue.toFixed(3))];
    const finalFteWithBorders = worksheetThree.addRow(finalFte);
    finalFteWithBorders.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
      cell.font = {
        size: 11,
        name: 'Arial'
      }
    });


    worksheetThree.getColumn(1).width = 20;
    for (let i = 2; i <= headerRow.length; i++) {
      worksheetThree.getColumn(i).width = 10;
    }

    const headerRowFour = worksheetFour.addRow(['Months', ...headerNames.map((month: any) => (month))]);
    headerRowFour.font = {
      bold: true, size: 11,
      name: 'Arial'
    };
    headerRowFour.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }

    });
    worksheetFour.getColumn(1).width = 28;
    for (let i = 2; i <= headerRow.length; i++) {
      worksheetFour.getColumn(i).width = 12;
    }

    let deliverable = worksheetFour.addRow(['No of Deliverables', ...this.deliverableVal.map((deli: any) => Number(deli))]);
    deliverable.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
      cell.font = {
        size: 11,
        name: 'Arial'
      }
    });
    let Effortsheet = worksheetFour.addRow(['Efforts for the Month', ...this.effortVal.map((effort: any) => Number(effort * 8))]);
    Effortsheet.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
      cell.font = {
        size: 11,
        name: 'Arial'
      }
    });
    let effortspend = worksheetFour.addRow(['Efforts spend for deliverables', ...this.effortVal.map((effort: any) => Number(effort * 8))]);
    effortspend.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
      cell.font = {
        size: 11,
        name: 'Arial'
      }
    });
    worksheetFive.columns = [
      { header: '#', key: 'milestone', width: 5 },
      { header: 'MILESTONE / DELIVERABLE', key: 'des' },
      { header: 'Milestone Date', key: 'date', width: 17 },
      {
        header: 'PRICE', key: 'price', width: 10,
      }
    ];
    let column = worksheetFive.getColumn(2);
    column.alignment = { wrapText: true };
    column.width = 50;
    const maxCols = worksheetFive.columns.length;
    for (let i = 0; i < this.largeMontharr.length; i++) {
      let end = new Date(this.largeMontharr[i]);

      let headerThree = worksheetFive.addRow({
        milestone: (i + 1),
        des: 'Milestone ' + '' + (i + 1) + '-' + this.getactivity(end),
        date: this.AcceptDateArr[i],
        effort: Number(this.endeffortArr[i] * 8),

      });

      const headerRow = worksheetFive.getRow(1);
      headerThree.getCell(maxCols).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      headerRow.eachCell(cell => {
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
      worksheetFive.eachRow((headerThree, rowNumber) => {
        if (rowNumber !== 1) {
          headerThree.eachCell((cell, colNumber) => {
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
            cell.font = {
              size: 11,
              name: 'Arial'
            };
          });
        }
      });





      worksheetFive.eachRow((headerThree, rowNumber) => {
        if (rowNumber === 1) {
          headerThree.eachCell(cell => {
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
            cell.font = {
              bold: true,
              size: 11,
              name: 'Arial',
              color: { argb: 'FFFFFFFF' }
            }
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '#6990F2' }
            }
          });
        }
      });
    }

    let blob: any;
    workBook.xlsx.writeBuffer().then(data => {
      blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      saveAs(blob, 'Project_output.xlsx');
    })

  }

  getHolidaytemplate() {
    let workBook = new Workbook();
    let holidayWork = workBook.addWorksheet('Sheet1');
    let headerArr = ['S.No', 'Date', 'Name'];
    let headerTwo = holidayWork.addRow(headerArr);
    for (let j = 1; j <= headerArr.length; j++) {
      let col = headerTwo.getCell(j);
      col.font = {
        bold: true,
        size: 11,
        name: 'Arial'
      };
      col.border = {
        top: { style: 'thin' },
        left: {},
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

    }

    let holidayArr: any = [
      {
        "Date": "2023-01-01T18:29:50.000Z",
        "Name": "New Year Day"
      },
      {
        "Date": "2023-01-26T18:30:00.000Z",
        "Name": "Republic Day"
      },
      {
        "Date": "2023-01-05T18:29:50.000Z",
        "Name": "May Day"
      },
      {
        "Date": "2023-03-07T18:29:50.000Z",
        "Name": "US Independence Day"
      },
      {
        "Date": "2023-05-25T18:30:00.000Z",
        "Name": "Memorial Day"
      },
      {
        "Date": "2023-08-15T18:30:00.000Z",
        "Name": "Independence Day"
      },
      {
        "Date": "2023-07-09T18:29:50.000Z",
        "Name": "Labour Day"
      },
      {
        "Date": "2023-10-02T18:29:50.000Z",
        "Name": "Gandhi Jayanthi"
      },
      {
        "Date": "2023-11-13T18:30:00.000Z",
        "Name": "Deepavali Eve"
      },
      {
        "Date": "2023-11-26T18:30:00.000Z",
        "Name": "Thanks Giving Day"
      },
      {
        "Date": "2023-11-27T18:30:00.000Z",
        "Name": "Day After Thanks Giving"
      },
      {
        "Date": "2023-12-25T18:30:00.000Z",
        "Name": "Christmas"
      },
      {
        "Date": "2024-01-01T18:29:50.000Z",
        "Name": "New Year Day"
      },
      {
        "Date": "2024-01-26T18:29:50.000Z",
        "Name": "Republic Day"
      }

    ]
    holidayArr.forEach((data: any, index: any) => {
      let date = new Date(data.Date);
      const row = holidayWork.addRow([index + 1, moment(date).format('MM/DD/YYYY'), data.Name]);
      for (let j = 1; j <= headerArr.length; j++) {
        let coln = row.getCell(j)
        coln.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        coln.font = {
          size: 11,
          name: 'Arial'
        };

      }

    })

    let blob: any;
    workBook.xlsx.writeBuffer().then(data => {
      blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      saveAs(blob, 'Companyholiday.xlsx');
    })
  }

  getProjectTemplate() {
    let workBook = new Workbook();
    let projectWork = workBook.addWorksheet('Sheet1');
    let headerArr = ['S.No', 'Phase', 'Deliverable', 'Activity', 'StartDate', 'EndDate', 'Effort', 'AcceptanceDate'];
    let headerTwo = projectWork.addRow(headerArr);
    for (let j = 1; j <= headerArr.length; j++) {
      let col = headerTwo.getCell(j);
      col.font = {
        bold: true,
        size: 11,
        name: 'Arial'
      };
      col.border = {
        top: { style: 'thin' },
        left: {},
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

    }

    let ProjectArr: any = [
      {
        "Phase": "Development",
        "Deliverable": "Y",
        "Activity": "sign up page creation",
        "StartDate": "2023/01/05",
        "EndDate": "01/30/2023",
        "Effort": "48",
        "AcceptanceDate": "5/12/2023",


      },
      {
        "Phase": "Testing",
        "Deliverable": "Y",
        "Activity": "sign up page testing",
        "StartDate": "4/24/2023",
        "EndDate": "5/11/2023",
        "Effort": "48",
        "AcceptanceDate": "10/12/2023",
      },

    ]
    ProjectArr.forEach((data: any, index: any) => {
      let start = new Date(data.StartDate);
      let end = new Date(data.EndDate);
      let accept = new Date(data.AcceptanceDate);

      const row = projectWork.addRow([index + 1, data.Phase, data.Deliverable, data.Activity, moment(start).format('MM/DD/YYYY'), moment(end).format('MM/DD/YYYY'), data.Effort, moment(accept).format('MM/DD/YYYY')]);
      for (let j = 1; j <= headerArr.length; j++) {
        let coln = row.getCell(j)
        coln.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        coln.font = {
          size: 11,
          name: 'Arial'
        };

      }

    })

    let blob: any;
    workBook.xlsx.writeBuffer().then(data => {
      blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      saveAs(blob, 'Project_activity.xlsx');
    })
  }




  datavalues(data: any, concatenatedArray: any) {
    let arr: any = [];
    concatenatedArray.map((con: any) => {

      arr.push(data[con]);

    })

    return arr;
  }

  getRandomLightColor() {
    let letters = '0123456789ABCDEF';
    let color = '#';
    for (let i = 0; i < 3; i++) {
      let value = Math.floor(Math.random() * 96) + 160;
      let hex = value.toString(16).padStart(2, '0');
      color += hex;
    }
    return color;
  }
  getactivity(date: any) {
    let Enddate = new Date(date);
    let month = Enddate.getMonth();
    let year = Enddate.getFullYear();

    let filteredArr = this.milestoneArr.filter((item: any) =>
      new Date(item.EndDate).getFullYear() === year &&
      new Date(item.EndDate).getMonth() === month
    );
    let valueArr: any = [];
    filteredArr.map((e: any, index: number) => {
      valueArr.push(e.Activity);

    })


    const outPut = this.removeCommonWords(valueArr);

    return outPut;
  }

  removeCommonWords(strArr: any) {
    if (strArr.length == 1) {
      return strArr; 
    }

    const wordArrs = strArr.map((str: any) => str.split(' '));

    const common = wordArrs[0].filter((word: any) => wordArrs.every((arr: any) => arr.includes(word)));

    const resultArrs = wordArrs.map((arr: any) => arr.filter((word: any) => !common.includes(word)));

    const result = resultArrs.map((arr: any) => arr.join(' '));
    const prefix = "Codebase and Unit Test reports of ";
    const suffix = " section related screens";
    const totalResult = prefix + result.join(' ') + suffix;

    return totalResult;
  }
  getaccptancedate(end: any, accept: any) {

    const date = new Date(end);
    let acceptDate = new Date(accept)
    const daysToAdd = this.acceptValue;
    let i = 0;
    console.log(daysToAdd, "this.acceptValue");
    if (this.acceptValue !== 0) {
      while (i < daysToAdd) {
        date.setDate(date.getDate() + 1);
        if (date.getDay() !== 0 && date.getDay() !== 6) { 
          i++;
        }
      }

      // convert Date object to string

      return moment(date).format('MM/DD/YYYY')
    }
    else {
      return moment(acceptDate).format('MM/DD/YYYY')
    }
  }
  getTotalAccept(end: any) {
    const date = new Date(end);

    const dayOfWeek = date.getDay();

    const nextDay = new Date(date);
    nextDay.setDate(date.getDate() + 1);

    const isLastDayOfMonth = nextDay.getMonth() !== date.getMonth();
    if (isLastDayOfMonth || dayOfWeek === 5 || dayOfWeek === 6) {
      this.AcceptDateArr.push(date.toLocaleDateString());
      return date.toLocaleDateString();
    } else {
      let nextWeekday = nextDay;
      while (nextWeekday.getDay() === 0 || nextWeekday.getDay() === 6) {
        nextWeekday.setDate(nextWeekday.getDate() + 1);
      }
      this.AcceptDateArr.push(nextWeekday.toLocaleDateString());
      return nextWeekday.toLocaleDateString(); 
    }
  }

}

