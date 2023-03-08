import { Component } from "@angular/core";

/* global console, Excel */

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  welcomeMessage = "Welcome";
  private isExcelFileCompressed: boolean = null;
  private base64String: string = null;

  async run() {
    let file;
    this.isExcelFileCompressed = false;
    file = await new Promise<number[]>(async (resolve, reject) => {
      for (let index = 0; index < 3; index++) {
        if (!this.isExcelFileCompressed) {
          try {
            file = await this._getExcelFile();
          } catch (error) {
            console.error("Debug info:", error.debugInfo);
          }
        } else {
          break;
        }
      }
      resolve(file);
    });
    if (this.isExcelFileCompressed) {
      this.base64String = this.getCurrentWorkbook(file);
      this.createNewFile().then((res) => {
        this.deletePivotTable("PivotTable1", "Sheet3");
      });
    }
  }
  private _getExcelFile() {
    return new Promise<number[]>((resolve, reject) => {
      Office.onReady(() => {
        Office.context.document.getFileAsync(Office.FileType.Compressed, async (file) => {
          if (file.status === Office.AsyncResultStatus.Failed) {
            Promise.all([])
              .then((datas) => {
                resolve(datas);
              })
              .finally(() => file.value.closeAsync());
          } else {
            let sliceRequest: Promise<number[]>[] = [];
            this.isExcelFileCompressed = true;
            for (let i = 0; i < file.value.sliceCount; i++) {
              sliceRequest.push(
                new Promise<number[]>((resolve, reject) => {
                  file.value.getSliceAsync(i, (slice) => {
                    if (slice.status === Office.AsyncResultStatus.Failed) {
                      reject(file.error);
                    } else {
                      resolve(slice.value.data);
                    }
                  });
                })
              );
            }
            Promise.all(sliceRequest)
              .then((datas) => {
                resolve(
                  datas.reduce((p, c) => {
                    return p.concat(c);
                  }, [])
                );
              })
              .finally(() => file.value.closeAsync());
          }
        });
      });
    });
  }
  private getCurrentWorkbook(bytes: number[]) {
    var base64String = Array.prototype.map.call(bytes, (c) => String.fromCharCode(c)).join(String());
    return btoa(base64String);
  }
  async createNewFile() {
    return new Promise((resolve, reject) => {
      if (this.isExcelFileCompressed) {
        const base64String: string = this.base64String;
        Excel.run(async function (context) {
          await Excel.createWorkbook(base64String).then((res) => {
            console.log("then Log:", res);
          });
          return context.sync().then(() => {
            resolve("create done");
          });
        }).catch((error) => {
          reject(error);
          console.error("Debug info:", error.debugInfo);
        });
      }
    });
  }
  async deletePivotTable(pivotName: string, workbookId: string) {
    return new Promise(async (resolve, reject) => {
      try {
        await Excel.run(async (context) => {
          let sheet = context.workbook.worksheets.getItem(workbookId);
          sheet.load("name");
          await context.sync();
          console.log(sheet.name);
          let pivot = sheet.pivotTables.getItem(pivotName);
          pivot.load("name");
          await context.sync();
          console.log(pivot.name);
          pivot.delete();
          await context.sync();
        }).catch((error) => {
          console.error("Debug info:", error.debugInfo);
        });
      } catch (e) {
        reject("error");
        console.error("Delete Pivot table error", e);
        throw e;
      }
      resolve("done");
    });
  }
}
