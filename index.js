const xlsx = require("xlsx");
const fs = require("fs");
const dayjs = require("dayjs");
const _ = require("lodash");
class excelHandle {
  week;
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
    this.week = dataArr[0];
    // 该hooks用于确定当月是否有那么多天数和格式化月份
    const { chartMonth, MonthFormat } = this.getMonth(dataArr);
    // 取出员工信息
    const newExcelData = dataArr
      .filter(
        (item) =>
          item["姓名"] && !Object.values(item).some((value) => value === "淘汰")
      )
      .map((item) => {
        const newData = {}; // 新表格的数据
        // 对每个角色进行处理
        this.handleMonth({ item, MonthFormat, chartMonth });
        newData["姓名"] = item["姓名"];
        newData["直播游戏"] = item["所属项目"];
        newData["基本工资"] = item["工资单价/天"];
        newData["主播分成比例"] = item["主播分成比例"];
        // newData["核算周期"] =
        return newData;
      });
    // TODO：生成新的excel表
    // this.createNewExcel(newExcelData);
  }
  handleMonth(time) {
    const { item, MonthFormat, chartMonth } = time;
    const Attendance = [];
    const day = {};
    for (let i = 1; i <= chartMonth; i++) {
      day[`${MonthFormat}-${i}`] = item[`__EMPTY_${i}`] === 1 ? 1 : 0;
      console.log(Object.values(this.week));
      // console.log(this.week[`__EMPTY_${i}`]);
      Attendance.push(day);
    }
    // return Attendance;
  }
  getMonth(month) {
    const weekArr = Object.keys(month[0]);
    let MonthFormat = weekArr[0].replace(/年/, "-").replace(/月/, "");
    // 截取表中月份(目的：确定当月是否有那么多天数)
    const chartMonth = dayjs(MonthFormat).daysInMonth();
    const monthObj = _.pick(month[1], weekArr);
    if (chartMonth !== Object.values(monthObj).length)
      throw new Error("月份不匹配");
    // const newMonthArr = [];
    // let key = 0;
    // let arr = [];
    // let i = 1;
    // const monthLength = Object.values(month[0]).length;

    // Object.values(month[0]).forEach((item, index) => {
    //   let tem = {};
    //   tem[`${MonthFormat}-${i}`] = 0;
    //   i++;
    //   arr.push(tem);
    //   if (item == "日") {
    //     newMonthArr.push(arr);
    //     key++;
    //     arr = [];
    //   }
    //   if (index == monthLength - 1 && item != "日") {
    //     newMonthArr.push(arr);
    //   }
    // });
    return { chartMonth, MonthFormat };
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
