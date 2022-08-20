const xlsx = require("xlsx");
const fs = require("fs");
const dayjs = require("dayjs");
const _ = require("lodash");
class excelHandle {
  readFile() {
    return new Promise((resolve, reject) => {
      fs.readdir(`${__dirname}/excel/`, (err, files) => {
        if (err) reject("文件读取出错:" + err);
        resolve(files);
      });
    });
  }
  readExcel() {
    this.readFile().then((files) => {
      const workBook = xlsx.readFile(`${__dirname}/excel/${files}`);
      this.dataHandle(workBook.Sheets["考勤"]);
    });
  }
  dataHandle(data) {
    const dataArr = xlsx.utils.sheet_to_json(data);
    console.log(dataArr);
    // // 取出月份
    const { monthObj, newMonthArr, chartMonth } = this.useMonth([
      dataArr[0],
      dataArr[1],
    ]);
    // console.log(monthObj, newMonthArr, chartMonth);
    const newExcelData = dataArr
      .filter(
        (item) =>
          item["姓名"] && !Object.values(item).some((value) => value === "淘汰")
      )
      .map((item, index) => {
        const newDataObj = {};
        // console.log(item)
        newDataObj["month"] = newMonthArr.forEach((item) => {
          for (let i = 1; i < 5; i++) {
            // console.log(i)
          }
        });
      });

    // console.log(monthObj,newMonthArr);
    // for(let i in monthObj){
    //   monthObj[i] = 0
    // }
    // const monthMap = new Map(Object.entries(monthObj))

    // // 取出员工信息
    // const newExcelData = dataArr
    //   .filter(
    //     (item) =>
    //       item["姓名"] && !Object.values(item).some((value) => value === "淘汰")
    //   )
    //   .map((item) => {
    //     console.log(item)
    //     item["month"] = []
    //     for(let i = 1; i<=chartMonth; i++){
    //       if(i === 1) item['month'][Object.keys(monthObj)[0]] = item[Object.keys(monthObj)[0]]
    //      console.log(item)
    //       item["month"][`__EMPTY_${i}`] = item[`__EMPTY_${i}`] === 1 ? 1 : 0
    //     }
    //     console.log(item)
    //     const newData = {};
    //     newData["姓名"] = item["姓名"];
    //     newData["直播游戏"] = item["所属项目"];
    //     newData["基本工资"] = item["工资单价/天"];
    //     newData["主播分成比例"] = item["主播分成比例"];
    //     // newData["核算周期"] =
    //     return newData;
    //   });
    // TODO：生成新的excel表
    // this.createNewExcel(newExcelData);
  }
  useMonth(month) {
    const weekArr = Object.keys(month[0]);
    // 截取表中月份(目的：确定当月是否有那么多天数)
    const chartMonth = dayjs(
      weekArr[0].replace(/年/, "-").replace(/月/, "")
    ).daysInMonth();
    const monthObj = _.pick(month[1], weekArr);
    if (chartMonth !== Object.values(monthObj).length)
      throw new Error("月份不匹配");
    const newMonthArr = [];
    let key = 0;
    let arr = [];
    let i = 1;
    const monthLength = Object.values(month[0]).length;
    let thatMonth = weekArr[0].replace(/年/, "-").replace(/月/, "");
    Object.values(month[0]).forEach((item, index) => {
      let tem = {};
      tem[`${thatMonth}-${i}`] = 0;
      i++;
      arr.push(tem);
      if (item == "日") {
        newMonthArr.push(arr);
        key++;
        arr = [];
      }
      if (index == monthLength - 1 && item != "日") {
        newMonthArr.push(arr);
      }
    });

    return { monthObj, newMonthArr, chartMonth };
  }
  // 生成处理好的excel表
  createNewExcel(data) {
    const workBook = xlsx.utils.book_new();
    const newSheet = xlsx.utils.json_to_sheet(data);
    xlsx.utils.book_append_sheet(workBook, newSheet, "工资结算");
    xlsx.writeFile(workBook, "工资结算.xlsx");
  }
}
const eHandle = new excelHandle();
eHandle.readExcel();
