const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("node:path");
const XLSX = require("xlsx");
const writeXlsxFile = require("write-excel-file/node");
const fs = require('fs');
if (require("electron-squirrel-startup")) {
  app.quit();
}

const createWindow = () => {
  const mainWindow = new BrowserWindow({
    title: 'Nimar Motors Khargone',
    // width: 1290,
    // height: 1080,
    icon: path.join(__dirname, './assets/NimarMotor.png'),
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });
  ipcMain.on('reset-application', () => {
    mainWindow.reload();
  });
  mainWindow.once('ready-to-show', () => {
    mainWindow.maximize()
  })
  mainWindow.loadFile(path.join(__dirname, "index.html"));
  // mainWindow.webContents.openDevTools();
};


let data1 = [];
let data2 = [];
let dataForExcelObj = [];
let dat1;
let dat2;
let cMap;


//changes start
let interestPercent = 0;
let noDueDays = 0;
let EndDate = 0;
let copyInterestPercent = 0;
let copyNoDueDays = 0;
let copyEndDate = 0;
let customerIdVal = '';
function calculateDaysBetween(startDate, EndDate) {
  const start = new Date(startDate);
  const end = new Date(EndDate);
  const diffTime = Math.abs(end - start);
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

// let EndDate = "2025-03-31T18:30:00.000Z";
function interestAmount(dueAmt, dueDays) {
  return (dueDays * (((dueAmt) * interestPercent) / 365));
}

function applyPaymentsAndCalculateInterest(datt1, datt2) {

  // cMap = datt1.reduce((acc, purchase) => {
  //   if(EndDate >  purchase['Date']){
  //     if (!acc[purchase['Customer Code']]) acc[purchase['Customer Code']] = [];
  //     acc[purchase['Customer Code']].push({ ...purchase, RemainingChallanAmount: purchase['Total Amount'], LastPaymentDate: 0, interest: 0 });
  //     return acc;
  //   }
  // }, {});

  if (customerIdVal !== '') {
    cMap = datt1.reduce((acc, purchase) => {
      const parsedPurchaseDate = XLSX.SSF.parse_date_code(purchase.Date);
      const jsPurchaseDate = new Date(parsedPurchaseDate.y, parsedPurchaseDate.m - 1, parsedPurchaseDate.d);
      if (jsPurchaseDate < new Date(EndDate)) {
        if (purchase['Customer Name'].includes(customerIdVal) || purchase['Customer Code'].includes(customerIdVal)) {
          acc[customerIdVal] = acc[customerIdVal] || [];
          acc[customerIdVal].push({
            ...purchase,
            RemainingChallanAmount: purchase['Total Amount'],
            LastPaymentDate: 0,
            interest: 0
          });
        }
      }
      return acc;
    }, {});
  } else {
    cMap = datt1.reduce((acc, purchase) => {
      const parsedPurchaseDate = XLSX.SSF.parse_date_code(purchase.Date);
      const jsPurchaseDate = new Date(parsedPurchaseDate.y, parsedPurchaseDate.m - 1, parsedPurchaseDate.d);
      if (jsPurchaseDate < new Date(EndDate)) {
        if (!acc[purchase['Customer Code']]) acc[purchase['Customer Code']] = [];
        acc[purchase['Customer Code']].push({
          ...purchase,
          RemainingChallanAmount: purchase['TOTAL (Final Challan Price)'],
          LastPaymentDate: 0,
          interest: 0
        });
      }
      return acc;
    }, {});
  }




  datt2.forEach(payment => {
    if (cMap[payment['Customer Code']]) {
      let RemainingChallanAmountPayment = payment['Total Amount'];
      for (const purchase of cMap[payment['Customer Code']]) {
        if (RemainingChallanAmountPayment === 0) break;
        if (purchase.RemainingChallanAmount > 0) {
          const parsedDate1 = XLSX.SSF.parse_date_code(purchase.Date);
          const jsDate1 = new Date(parsedDate1.y, parsedDate1.m - 1, parsedDate1.d, parsedDate1.H, parsedDate1.M, parsedDate1.S);
          const parsedDate2 = XLSX.SSF.parse_date_code(payment.Date);
          const jsDate2 = new Date(parsedDate2.y, parsedDate2.m - 1, parsedDate2.d, parsedDate2.H, parsedDate2.M, parsedDate2.S);
          let dueDays = 0;
          const daysPastDue = calculateDaysBetween(jsDate1, jsDate2);
          purchase["Customer Code"]
          if (purchase.LastPaymentDate !== 0) {
            const parsedDate3 = XLSX.SSF.parse_date_code(purchase.LastPaymentDate);
            const jsDate3 = new Date(parsedDate3.y, parsedDate3.m - 1, parsedDate3.d, parsedDate3.H, parsedDate3.M, parsedDate3.S);
            dueDays = calculateDaysBetween(jsDate3, jsDate2);
          }

          if (parseInt(payment.Date) < parseInt(purchase.Date)) {

            const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
            purchase.RemainingChallanAmount -= deduction;
            RemainingChallanAmountPayment -= deduction;
            purchase.LastPaymentDate = payment.Date;
          } else {
            //normal deduction with no interest charges
            if (daysPastDue <= parseInt(noDueDays)) {
              const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
              purchase.RemainingChallanAmount -= deduction;
              RemainingChallanAmountPayment -= deduction;
              purchase.LastPaymentDate = payment.Date;
            } else //deduction with interest charges
            {
              if (daysPastDue <= (parseInt(noDueDays) * 2)) {
                if (purchase.LastPaymentDate === 0) {
                  const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                  purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));
                  purchase.RemainingChallanAmount -= deduction;
                  RemainingChallanAmountPayment -= deduction;
                  purchase.LastPaymentDate = payment.Date;
                } else {
                  if (parseInt(purchase.Date) > parseInt(purchase.LastPaymentDate)) {
                    const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                    purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));
                    purchase.RemainingChallanAmount -= deduction;
                    RemainingChallanAmountPayment -= deduction;
                    purchase.LastPaymentDate = payment.Date;
                  } else {
                    const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                    purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));
                    purchase.RemainingChallanAmount -= deduction;
                    RemainingChallanAmountPayment -= deduction;
                    purchase.LastPaymentDate = payment.Date;
                  }
                }
              } else {
                // calculate interest from last payment date;
                //deduct amount and set lastpayment date
                if (purchase.LastPaymentDate === 0) {
                  const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                  purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));
                  purchase.RemainingChallanAmount -= deduction;
                  RemainingChallanAmountPayment -= deduction;
                  purchase.LastPaymentDate = payment.Date;
                } else {
                  if (parseInt(purchase.LastPaymentDate) < parseInt(purchase.Date)) {
                    const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                    purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));
                    purchase.RemainingChallanAmount -= deduction;
                    RemainingChallanAmountPayment -= deduction;
                    purchase.LastPaymentDate = payment.Date;
                  } else {
                    if (parseInt(purchase.LastPaymentDate) < parseInt(purchase.Date) + parseInt(noDueDays)) {
                      const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                      purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));
                      purchase.RemainingChallanAmount -= deduction;
                      RemainingChallanAmountPayment -= deduction;
                      purchase.LastPaymentDate = payment.Date;
                    } else {
                      const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                      purchase.interest += interestAmount(purchase.RemainingChallanAmount, dueDays);
                      purchase.RemainingChallanAmount -= deduction;
                      RemainingChallanAmountPayment -= deduction;
                      purchase.LastPaymentDate = payment.Date;
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  })
  console.log("cMap::", cMap)


  //calculation for End Date
  let ids = Object.keys(cMap);
  ids.forEach(id => {
    cMap[id].forEach(obj => {
      if (obj.RemainingChallanAmount > 0) {
        console.log("Customer Code:::", obj["Customer Code"]);
        console.log("Date:::", obj["Date"]);

        const parsedDate1 = XLSX.SSF.parse_date_code(obj.Date);
        console.log("parsedDate1:::", parsedDate1);

        const jsDate1 = new Date(parsedDate1.y, parsedDate1.m - 1, parsedDate1.d, parsedDate1.H, parsedDate1.M, parsedDate1.S);
        console.log("jsDate1:::", jsDate1);

        let dueDays = 0;
        let daysPastDue = calculateDaysBetween(jsDate1, EndDate);
        console.log("daysPastDue:::", daysPastDue);

        if (obj.LastPaymentDate !== 0) {
          const parsedDate3 = XLSX.SSF.parse_date_code(obj.LastPaymentDate);
          const jsDate3 = new Date(parsedDate3.y, parsedDate3.m - 1, parsedDate3.d, parsedDate3.H, parsedDate3.M, parsedDate3.S);
          dueDays = calculateDaysBetween(jsDate3, EndDate);
        }
        if (obj.LastPaymentDate === 0) {
          obj.interest += interestAmount(obj.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));

        } else {
          if (parseInt(obj.Date) > parseInt(obj.LastPaymentDate)) {
            obj.interest += interestAmount(obj.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));

          } else {
            if (parseInt(obj.Date) + parseInt(noDueDays) > parseInt(obj.LastPaymentDate)) {
              obj.interest += interestAmount(obj.RemainingChallanAmount, daysPastDue - parseInt(noDueDays));

            } else {
              obj.interest += interestAmount(obj.RemainingChallanAmount, dueDays);

            }
          }
        }
      }
    })
  })
  return cMap;
}

ipcMain.on("file-selected1", (event, path) => {
  const workbook1 = XLSX.readFile(path);
  const sheetName1 = workbook1.SheetNames[0];
  const sheet1 = workbook1.Sheets[sheetName1];
  data1 = XLSX.utils.sheet_to_json(sheet1);
  dat1 = data1;
  dat1.forEach(obj => {
    delete obj.Sno;
  })
});

ipcMain.on("file-selected2", (event, path) => {
  const workbook2 = XLSX.readFile(path);
  const sheetName2 = workbook2.SheetNames[0];
  const sheet2 = workbook2.Sheets[sheetName2];
  data2 = XLSX.utils.sheet_to_json(sheet2);
  dat2 = data2;
  const dataForExcel = applyPaymentsAndCalculateInterest(dat1, dat2);
  console.log(dataForExcelObj);
  let ids = Object.keys(dataForExcel);

  ids.forEach(id => {

    dataForExcel[id].forEach((row) => {

      let newObj = {};
      const parsedDate1 = XLSX.SSF.parse_date_code(row.Date);
      const jsDate1 = new Date(parsedDate1.y, parsedDate1.m - 1, parsedDate1.d);

      let jsDate2 = "-";
      if (row.LastPaymentDate != 0) {
        const parsedDate2 = XLSX.SSF.parse_date_code(row.LastPaymentDate);
        jsDate2 = new Date(parsedDate2.y, parsedDate2.m - 1, parsedDate2.d);
      }
      newObj = {
        "Party Id": id,
        "Challan No": row["Challan No."],
        "Party Name": row["Customer Name"],
        "Challan Date": jsDate1,
        "TOTAL (Final Challan Price)": row["TOTAL (Final Challan Price)"],
        "Payment Date": jsDate2,
        "Amount Left": Math.round(row.RemainingChallanAmount),
        "Interest Amount (13.5% per annum)": Math.round(row.interest),
      }
      dataForExcelObj.push(newObj);
    })
    console.log("dataForExcelObj::::", JSON.stringify(dataForExcelObj));
  })
  console.log("event")

});
ipcMain.on('form-submitted', (event) => {
  event.reply("dataForExcelObj", dataForExcelObj);



  const nowDate = new Date();
  const month = nowDate.getMonth() + 1;
  const date = nowDate.getDate();
  const year = nowDate.getFullYear();
  const time = nowDate.toLocaleTimeString().replace(/:/g, '-');

  const newWorkbook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.json_to_sheet(dataForExcelObj);
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Sheet1");

  const fileName = `calculatedInterestAmount_${customerIdVal ? `(${customerIdVal})` : ""}_${date}-${month}-${year}_${time}.xlsx`;
  const folderPath = "./DataSheets";
  if (!fs.existsSync(folderPath)) {
    fs.mkdirSync(folderPath);
    console.log(`Directory ${folderPath} created.`);
  } else {
    console.log(`Directory ${folderPath} already exists.`);
  }
  XLSX.writeFile(newWorkbook, `./DataSheets/${fileName}`);

  // Clear the data arrays
  data1 = [];
  data2 = [];
  dataForExcelObj = [];
  dat1 = [];
  dat2 = [];
  cMap = [];
  customerIdVal = '';
});


ipcMain.on("days", (event, data) => {
  noDueDays = data;

  console.log("parseInt(noDueDays)::", parseInt(noDueDays))
});

ipcMain.on("date", (event, data) => {
  const date = new Date(data);
  const formattedDate = date;
  EndDate = formattedDate;
  console.log("EndDate::", EndDate);
});

ipcMain.on("percent", (event, data) => {
  interestPercent = data / 100;
  console.log("interestPercent::", interestPercent)
});

ipcMain.on("customerIdVal", (event, data) => {
  customerIdVal = data;
  console.log("customerIdVal::", customerIdVal);
});


app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
