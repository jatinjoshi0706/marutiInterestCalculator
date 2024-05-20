const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("node:path");
const XLSX = require("xlsx");
const writeXlsxFile = require("write-excel-file/node");

if (require("electron-squirrel-startup")) {
  app.quit();
}

const createWindow = () => {
  const mainWindow = new BrowserWindow({
    title: 'Nimar Motors Khargone',
    width: 800,
    height: 600,
    icon: path.join(__dirname, 'NimarMotor.png'),
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));
  // mainWindow.webContents.openDevTools();
};

let data1 = [];
let data2 = [];
let dataForExcelObj = [];
let dat1;
let dat2;
let cMap;

function calculateDaysBetween(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const diffTime = Math.abs(end - start);
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

let EndDate = "2025-03-31T18:30:00.000Z";
function interestAmount(dueAmt, dueDays) {
  return (dueDays * (((dueAmt) * 0.135) / 365));
}

function applyPaymentsAndCalculateInterest(datt1, datt2) {

  cMap = datt1.reduce((acc, purchase) => {
    if (!acc[purchase['Customer Code']]) acc[purchase['Customer Code']] = [];
    acc[purchase['Customer Code']].push({ ...purchase, RemainingChallanAmount: purchase['Total Amount'], LastPaymentDate: 0, interest: 0 });
    return acc;
  }, {});

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

            if (daysPastDue <= 10) {
              const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
              purchase.RemainingChallanAmount -= deduction;
              RemainingChallanAmountPayment -= deduction;
              purchase.LastPaymentDate = payment.Date;
              // normal deduct and payment date set
            } else {
              if (daysPastDue <= 20) {
                if (purchase.LastPaymentDate === 0) {

                  const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                  purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - 10);
                  purchase.RemainingChallanAmount -= deduction;
                  RemainingChallanAmountPayment -= deduction;
                  purchase.LastPaymentDate = payment.Date;
                } else {

                  if (parseInt(purchase.Date) > parseInt(purchase.LastPaymentDate)) {
                    const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                    purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - 10);
                    purchase.RemainingChallanAmount -= deduction;
                    RemainingChallanAmountPayment -= deduction;
                    purchase.LastPaymentDate = payment.Date;
                  } else {
                    const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                    purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - 10);
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
                  purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - 10);
                  purchase.RemainingChallanAmount -= deduction;
                  RemainingChallanAmountPayment -= deduction;
                  purchase.LastPaymentDate = payment.Date;
                } else {
                  if (parseInt(purchase.LastPaymentDate) < parseInt(purchase.Date)) {
                    const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                    purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - 10);
                    purchase.RemainingChallanAmount -= deduction;
                    RemainingChallanAmountPayment -= deduction;
                    purchase.LastPaymentDate = payment.Date;
                  } else {
                    if (parseInt(purchase.LastPaymentDate) < parseInt(purchase.Date) + 10) {
                      const deduction = Math.min(purchase.RemainingChallanAmount, RemainingChallanAmountPayment);
                      purchase.interest += interestAmount(purchase.RemainingChallanAmount, daysPastDue - 10);
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

  let ids = Object.keys(cMap);
  ids.forEach(id => {
    cMap[id].forEach(obj => {
      if (obj.RemainingChallanAmount > 0) {
        const parsedDate1 = XLSX.SSF.parse_date_code(obj.Date);
        const jsDate1 = new Date(parsedDate1.y, parsedDate1.m - 1, parsedDate1.d, parsedDate1.H, parsedDate1.M, parsedDate1.S);
        let dueDays = 0;
        const daysPastDue = calculateDaysBetween(jsDate1, EndDate);
        if (obj.LastPaymentDate !== 0) {
          const parsedDate3 = XLSX.SSF.parse_date_code(obj.LastPaymentDate);
          const jsDate3 = new Date(parsedDate3.y, parsedDate3.m - 1, parsedDate3.d, parsedDate3.H, parsedDate3.M, parsedDate3.S);
          dueDays = calculateDaysBetween(jsDate3, EndDate);
        }
        if (obj.LastPaymentDate === 0) {
          obj.interest += interestAmount(obj.RemainingChallanAmount, daysPastDue - 10);
        } else {
          if (parseInt(obj.Date) > parseInt(obj.LastPaymentDate)) {
            obj.interest += interestAmount(obj.RemainingChallanAmount, daysPastDue - 10);
          } else {
            if (parseInt(obj.Date) + 10 > parseInt(obj.LastPaymentDate)) {
              obj.interest += interestAmount(obj.RemainingChallanAmount, daysPastDue - 10);
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
        "Party Name": row["Customer Name"],
        "Challan Date": jsDate1,
        "Total Challan Amount": row["Total Amount"],
        "Payment Date": jsDate2,
        "Amount Left": row.RemainingChallanAmount,
        "Interest Amount (13.5% per annum)": row.interest,
      }
      dataForExcelObj.push(newObj);
    })
    console.log(JSON.stringify(dataForExcelObj));
  })
  console.log("event")
  event.reply("dataForExcelObj", dataForExcelObj);
  console.log("jnsbjkx")
  const newWorkbook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.json_to_sheet(dataForExcelObj);
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Sheet1");
  XLSX.writeFile(newWorkbook, "finalDataSheet.xlsx");
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