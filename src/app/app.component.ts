import { Component, ViewEncapsulation } from '@angular/core';

import * as GC from "@grapecity/spread-sheets";
import * as GcDesigner from '@grapecity/spread-sheets-designer';
import * as ExcelIO from '@grapecity/spread-excelio';
import "@grapecity/spread-sheets-print";
import "@grapecity/spread-sheets-shapes";
import "@grapecity/spread-sheets-pivot-addon";
import "@grapecity/spread-sheets-tablesheet";
import "@grapecity/spread-sheets-io";
import '@grapecity/spread-sheets-designer-resources-en';
import '@grapecity/spread-sheets-designer';

import { reportData } from './report';

GC.Spread.Common.CultureManager.culture("en-us");
// const sjsLicense = '3.7.89.78,E756192611822541#B16U4YlS7VkZ6cTdMdVMkJkZMdUZQVXc7skdw84azxkNGNGO9QTTYx4aNlUWQt4VNNEbsFXRygFOYZDMmBlbTVmT9Z5TzZ7YadWSzhUZ8V5MplDUpNHa4onTohUaxI7K7h5bCZkUotSakR7LmFFZrIkVlZ4LvhFa4pGZXB7dhhnRO94aTVVSvMUW9d4LiJGaYBVcttyMuRlVZZTO7UWMshVW7FlQwAjcPxkT4gFZSNjbSdja0Fje7smZFhmTz3iatl5UjZVaZticIFjeuF5YJx4ahRVQM3UUTx4NBZ5STF6YN94LrVlS7AXV9AleM34TY9GTwA5ZY3yalhWcOhUS0lHTMdTZJZGcQpHaiojITJCLiQkMzEEO5AjNiojIIJCLzMTNwkzN8AzM0IicfJye=#Qf35VfigUSKJkI0IyQiwiI8EjL6ByUKBCZhVmcwNlI0IiTis7W0ICZyBlIsISOwAjMwEDI8ITMwUjMwIjI0ICdyNkIsICNwMDM6IDMyIiOiAHeFJCLigzNukDOucjLzIiOiMXbEJCLi8CRUxEIuQlVQByUO34UgYCIPFkUgE4ROFkUu8kI0ISYONkIsUWdyRnOiwmdFJCLiEDN5IjM8ETM6ITOxYTN7IiOiQWSiwSfdJSZsJWYUR7b6lGUislOicGbmJCLlNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TPnFUcJR4QxtUWzNUU5Y6TllnTzt6YYhXcix6KMZ7VWNVe8gFd8Imb8cXdjxWTzN6Z8hXV9EzTQhmWT3WTYBXYGxENvg4NC9WTF5WeuI';
// const  sjsDesignerLicense = '3.7.89.78,E219113261896968#B1JRbQEdERYFzc6RTeZV7NulFM8AVcoRFRWVUZxNlVuFTMsFVbNtmaxkndINnWhpHWQtWcvMFMjdWW7p4UiJzRxNnQ5UFT0lmRxUzYKd7Z4Z5VmFXZLRXTnFUT0J5VSJHcMRjZEx6U0RldxNnUX5mcNFDO4QDbM3yLS3CbLJjV8NXR0JUVKNEVTVVS9gTcCV6bQFFUXRlTStmMuVGZzQHMD54dkxEc9Nzbz5EMXl7ZXZnUtB5SNVDbxFnbttGeP9WSo36LvRHR5IVb7VTRsN4bzNGR8gVSX9WexcEaYtSWY96ZtBjQ9E4VJt4NLpXUFN5SvgTY4hWaElGbJpmI0IyUiwiIFNUQ6QEMzYjI0ICSiwSM9kTO9AzMyMTM0IicfJye#4Xfd5nI9k5MzIiOiMkIsICOx8idg86bkRWQtIXZudWazVGRtMlSkFWZyB7UiojIOJyebpjIkJHUiwiI8EDMyATMggjMxATNyAjMiojI4J7QiwiI4AzMwYjMwIjI0ICc8VkIsICO78SO88yNuMjI0IyctRkIsIiLERFTg8CVWBFIT94TTBiJg2UQSBSQH9UQS9iTiojIh94QiwSZ5JHd0ICb6VkIsICO6kjN9gTM6IzMxETOxIjI0ICZJJye0ICRiwiI34zZ49kZa3GTXFzV5U6LmRkUzFVRI36L7klY9cjTTBjQ6gVV9EkYDJnWQ9WZCB5d8R6NQFkMFhFVxV6RBpmaNtyVsJ7dqdjSFJHWDRVcjp7VMNzdQ3kerYlQ2BcM';
const sjsLicense = 'Grapecity India,E541325983913486#B1KFVYqR7KUh4MY3yKHdzTtJUb9sydrNmShJUW4IETNx4LMhHRUhFSzx6Qvp4RKF6dzZTdmFjUuljarNXayskQ5hjRhZVWix4aZ3WMillTohUdU96UphlT5lzYwMzS0d7UDhmawoWMvdFS4gUZrUHUNB7bvBXUwomY5skQwJERat6T4lGeChUZrVXblNXTJBXMkVTVxN4bZRXYFVVauZTdNZTT4FXUX3GWzVXVMFjRzJkVzonV8IHaTtUaWpnMwUjdEd5VkZUNERENwpHbmplNTR6N7MWNv4mNYZFVlNlQ7hGbGJ5VxImQv54LTVTOvBHMrJzNWRXdm3iViojITJCLiYURzU4QGFTNiojIIJCLzkzN5QjNzMTO0IicfJye35XX3JCSJpkQiojIDJCLigTMuYHITpEIkFWZyB7UiojIOJyebpjIkJHUiwiI6AzN5gDMgQjM9ATNyAjMiojI4J7QiwiI5EDMxUjMwIjI0ICc8VkIsISYpRmbJBSe4l6YlBXYydkI0ISYONkIsUWdyRnOiwmdFJCLiYDO4MTM9MDO9UjMzEDN5IiOiQWSiwSfdJSZsJWYUR7b6lGUislOicGbmJCLlNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TP7pGdhp7RN3URpB7cGd6ZUh6TM3yTzIFauNEdI5mdGNTOsN6N52SYLVGetN5ZBNkMxEDc8EjYtF6MDVHRDhHbzNzNGlUNxpGUxZHOk3ScydlUORENuJ5QKlWcvVxW';
const  sjsDesignerLicense = 'Grapecity India,E677557551491623#B1UcvBHUNtWW6EnT65UWvZlQQdHM9NUdnpVMxVUcwITa8RWWZlnVvcWayJWQrVGbvYUbV56Y9dXUSVUb4dFUZhnSxpURYB7Rp5EbwVVerJEZjtUe936LFlnezI6VHtyVntiUBFzdYJlYMt6YSVVSx9kVQljQGJVT8dGcIRmd9k6NCFGUORHOaRlbUtCU4E4YWhlexgFZDJlRpt6QsdkWrIDRZNWYnRHUFBHS8IGdH5GMVJlZttES4xmZxsGdlZXaZJ6LvknS4lFSGlTWzEzapVjcNRke75UNIx4alBDOxh5YpxEbvlVaz5WNXtiQiojITJCLiIUM6gDR8UUNiojIIJCLyAzNyIDN4kjN0IicfJye35XX3JSOZNzMiojIDJCLigTMuYHIu3GZkFULyVmbnl6clRULTpEZhVmcwNlI0IiTis7W0ICZyBlIsISOwgTN8ADI4ITOwUjMwIjI0ICdyNkIsISNxATM5IDMyIiOiAHeFJCLiEWak9WSgkHdpNWZwFmcHJiOiEmTDJCLlVnc4pjIsZXRiwiIzIjNxkDNxUTN7UTN7cjNiojIklkI1pjIEJCLi4TPBVDeQRDd4EWewE4b5R7S8RGe9BHOmtyK8g6M8pVNRJEaTRHdTZmR5hES8UHVzkTRz4Gd9EGauNUdFRmdCF4cmR6VzV5UD3iZxc6ThllTSFXRQZDO0FWMNZVWrs6LGlkaZ9kbthzZaRPOhN';
GC.Spread.Sheets.LicenseKey = sjsLicense;
(ExcelIO as any).LicenseKey = sjsLicense;
(GC.Spread.Sheets as any).Designer.LicenseKey = sjsDesignerLicense;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  encapsulation: ViewEncapsulation.None
})
export class AppComponent {
  title = 'pivot-failure';
  report = reportData;
  workbook!: GC.Spread.Sheets.Workbook;
    props = {
    styleInfo: 'width: 100%; height: 100vh;',
    config: null
  }

  // Initializes the designer workbook, sets up the Primary Sales Report, and creates the Pivot Table.
  afterDesignerInit(e: {designer : GcDesigner.Spread.Sheets.Designer.Designer}) {
    const workbook = e.designer.getWorkbook() as GC.Spread.Sheets.Workbook;
    this.workbook = workbook;
    workbook.suspendEvent();
    workbook.suspendPaint();
    workbook.clearSheets();
    workbook.options.autoFitType = GC.Spread.Sheets.AutoFitType.cellWithHeader;
    workbook.setSheetCount(1);

    const dashboardSheet = workbook.getSheet(0), tabSheetName = "PrimarySalesReport";
    dashboardSheet.name('Pivotable Report');
    const dataManager = workbook.dataManager(), report = this.report;
    const tablePrimarySales = dataManager.addTable('tablePrimarySales', {
      remote: {
        read: function():Promise<any> {
          return Promise.resolve(report);
        }
      }
    });

    const tableSheet = workbook.addSheetTab(1, tabSheetName, GC.Spread.Sheets.SheetType.tableSheet);
    tableSheet.options.allowAddNew = false;
    workbook.setActiveSheetTab(tabSheetName);
    const columns = [
      { value: 'fiscalPeriod', width: 120, caption: 'Fiscal Period'},
      { value: 'quarter', width: 80, caption: 'Quarter'},
      { value: 'month', width: 150, caption: 'Month'},
      { value: 'channelName', width: 150, caption: 'Channel Name'},
      { value: 'subChannelName', width: 150, caption: 'Sub Channel Name'},
      { value: 'depotName', width: 200, caption: 'Depot Name'},
      { value: 'businessStateName', width: 200, caption: 'Business State Name'},
      { value: 'stateName', width: 200, caption: 'State Name'},
      { value: 'districtName', width: 200, caption: 'District Name'},
      { value: 'talukName', width: 200, caption: 'Taluk'},
      { value: 'placeName', width: 200, caption: 'Place Name'},
      { value: 'pincode', width: 250, caption:'Distributor Pincode'},
      { value: 'nameOfFirm', width: 200, caption: 'Name of Firm'},
      { value: 'rmsCode', width: 100, caption: 'RMS Code'},
      { value: 'shipToPt', width: 200, caption: 'SAP Customer Code'},
      { value: 'billDoc', width: 120, caption: 'Invoice Number'},
      { value: 'cancelInvoiceNo', width: 200, caption: 'Cancel Invoice Number'},
      { value: 'invoiceType', width: 120, caption: 'Invoice Type'},
      { value: 'productCode', width: 150, caption: 'Product Code'},
      { value: 'sapMaterialCode', width: 150, caption: 'SAP Material Code'},
      { value: 'productName', width: 200, caption: 'Product Name'},
      { value: 'segmentName', width: 150, caption: 'Category'},
      { value: 'subSegmentName', width: 180, caption: 'Sub Category'},
      { value: 'brandName', width: 150, caption: 'Brand Name'},
      { value: 'subBrandName', width: 150, caption: 'Variant'},
      { value: 'baseUom', width: 100, caption: 'UOM'},
      { value: 'netValue', width: 100, caption: 'Net Value'},
      { value: 'mrp', width: 70, caption: 'MRP'},
      { value: 'mrpSlab', width: 120, caption: 'MRP Slab'},
      { value: 'basicPrice', width: 120, caption: 'Basic Price'},
      { value: 'qty', width: 100, caption: 'Bill Quantity'},
      { value: 'gto', width: 80, caption: 'GTO'},
      { value: 'cdTdSrd', width: 200, caption: 'CDTDSRD'},
      { value: 'cd', width: 200, caption: 'CD'},
      { value: 'td', width: 200, caption: 'TD'},
      { value: 'tto', width: 80, caption: 'TTO'},
      { value: 'rdd', width: 200, caption: 'RDD'},
      { value: 'spd', width: 200, caption: 'SPD'},
      { value: 'srd', width: 200, caption: 'SRD'},
      { value: 'taxAmount', width: 200, caption: 'Tax Amount'},
      { value: 'sgst', width: 200, caption: 'SGST'},
      { value: 'cgst', width: 200, caption: 'CGST'},
      { value: 'igst', width: 200, caption: 'IGST'},
      { value: 'pricingProcedure', width: 200, caption: 'Pricing Procedure Name'},
      { value: 'customerReference', width: 200, caption: 'Customer Reference'},
      { value: 'sapCode', width: 100, caption: 'SAP Code'},
      { value: 'meName', width: 200, caption: 'ME Name'},
      { value: 'rsmName', width: 200, caption: 'RSM Name'},
      { value: 'asmName', width: 200, caption: 'ASM Name'},
      { value: 'aseName', width: 200, caption: 'ASE Name'},
      { value: 'soName', width: 200, caption: 'SO Name'},
      { value: 'srName', width: 200, caption: 'SR Name'},
      { value: 'replaceProductCode', width: 200, caption: 'Replaced Product'},
      { value: 'consumerPromo', width: 200, caption: 'Consumer Promo'},
      { value: 'productType', width: 200, caption: 'Product Type'},
      { value: 'barcode', width: 200, caption: 'Barcode'},
      { value: 'hsnNumber', width: 200, caption: 'HSN Number'},
      { value: 'productCategory', width: 150, caption: 'Custom Category'},
    ];
    tablePrimarySales.fetch().then(() => {
      const view = tablePrimarySales.addView('viewPSR');
      tableSheet.setDataView(view);
      this.initializePivotTable(dashboardSheet);
    });
    
    tableSheet.applyTableTheme(GC.Spread.Sheets.Tables.TableThemes.light13);
    // tableSheet.visible(false);
    workbook.resumeEvent();
    workbook.resumePaint();
  }

  // * Initializes the pivot table on the specified sheet with the provided configuration.
  initializePivotTable(sheet: GC.Spread.Sheets.Worksheet) {

    const pivotOptions = {
      bandRows: true,
      bandColumns: true,
      mergeItem: false,
      subtotalsPosition: GC.Spread.Pivot.SubtotalsPosition.none,
      insertBlankLineAfterEachItem: false
    };

    const dashboard = sheet.pivotTables.add('Pivotable Report', 'PrimarySalesReport', 1, 1, GC.Spread.Pivot.PivotTableLayoutType.tabular,
      GC.Spread.Pivot.PivotTableThemes.medium9, pivotOptions
    );

    const filterField = GC.Spread.Pivot.PivotTableFieldType.filterField, 
    rowField = GC.Spread.Pivot.PivotTableFieldType.rowField, 
    columnField = GC.Spread.Pivot.PivotTableFieldType.columnField, 
    valueField = GC.Spread.Pivot.PivotTableFieldType.valueField, 
    sumTypeSubTotal = GC.Pivot.SubtotalType.sum;

    dashboard.suspendLayout();
    dashboard.options.showRowHeader = true;
    dashboard.options.showColumnHeader = true;
    dashboard.add('Channel Name', 'Channel', filterField);
    dashboard.add('Political State', 'State', filterField);
    dashboard.add('Depot Name', 'Depot',filterField);
    dashboard.add('ME Name', 'ME',rowField);
    dashboard.add('ASE Name', 'ASE', rowField);
    dashboard.add('Segment', 'Segment', rowField);
    dashboard.add('Sub Segment', 'Sub Segment',rowField);
    dashboard.add('Fiscal Period','Fiscal Period',  columnField);
    dashboard.add('Quarter', 'Quarter', columnField);
    dashboard.add('Month', 'Month', columnField);
    dashboard.sort("Month", {
      customSortCallback: this.monthSort
    });

    dashboard.add('Quantity', 'Quantity', valueField, sumTypeSubTotal);
    dashboard.add('GTO', 'GTO', valueField, sumTypeSubTotal);
    dashboard.dataPosition(GC.Pivot.DataPosition.col, 3);

    sheet.bind(GC.Spread.Sheets.Events.PivotTableChanged, (sender: string, args: any) => {
      if(args.fieldName === "Month" && args.newIndex !== null){
        dashboard.sort("Month", {
          customSortCallback: this.monthSort
        });
      }
    });

    dashboard.resumeLayout();
    dashboard.autoFitColumn();
  }
  // Custom sorting function for sorting month names in chronological order.
  monthSort = (fieldItemNameArray: string[]) => {
    return fieldItemNameArray.sort((a: string, b: string) => {
      const monthsOrder: { [key: string]: number } = {
        "April": 1,
        "May": 2,
        "June": 3,
        "July": 4,
        "August": 5,
        "September": 6,
        "October": 7,
        "November": 8,
        "December": 9,
        "January": 10,
        "February": 11,
        "March": 12,
      };
  
      if (!isNaN(parseInt(a)) && !isNaN(parseInt(b))) {
        return parseInt(a) - parseInt(b); 
      } else if (!isNaN(parseInt(a))) {
        return -1; 
      } else if (!isNaN(parseInt(b))) {
        return 1;
      } else if (monthsOrder[a] && monthsOrder[b]) {
        return monthsOrder[a] - monthsOrder[b]; 
      } else if (monthsOrder[a]) {
        return -1;
      } else if (monthsOrder[b]) {
        return 1; 
      } else {
        return a.localeCompare(b); 
      }
    });
  }
}
