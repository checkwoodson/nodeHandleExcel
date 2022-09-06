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
  async readExcel() {
    await this.readFile().then((files) => {
      const workBook = files.map((file) => {
        return xlsx.readFile(`${__dirname}/excel/${file}`);
      });
      workBook.forEach((item) => {
        item.Sheets['考勤'] && this.dataHandle(item.Sheets['考勤'])
      });
      const game = workBook
        .map((item) => {
          return item.Sheets["工资明细表（对公）"];
        })
        .filter((item) => typeof item !== "undefined");
      game.length > 0  && this.gameConsume(game);
    });
  }
  gameConsume(value) {
    // 获取文件中的行和列
    const reg = new RegExp("小计");
    const allGame = {};
    value.forEach((item) => {
      const { r: totalRow } = xlsx.utils.decode_range(item["!ref"]).e;
      for (let r = 0; r < totalRow; r++) {
        const everyCol = item[xlsx.utils.encode_cell({ r, c: 0 })];
        if (reg.test(everyCol?.v)) {
          const gameName = everyCol.v.split("-")[0];
          const liveValue = item[xlsx.utils.encode_cell({ r, c: 12 })].v;
          if (typeof allGame[gameName] === "undefined") {
            allGame[gameName] = liveValue;
          } else {
            allGame[gameName] += liveValue;
          }
        }
      }
    });
    this.gameStatistics(allGame);
  }
  // 生成以直播游戏来统计并生成文件
  gameStatistics(value) {
    const gameAll = []
    for(let key in value){
      const game ={
        '直播游戏':key,
        '总价': value[key]
      }
      gameAll.push(game)
    }
    const workBook = xlsx.utils.book_new();
    const ws = xlsx.utils.json_to_sheet(gameAll);
    xlsx.utils.book_append_sheet(workBook, ws, "直播游戏总计");
    xlsx.writeFile(workBook, '直播游戏消费.xlsx')
  }
  dataHandle(data) {
    const dataArr = xlsx.utils.sheet_to_json(data);
    this.weeks = dataArr[0];

    // 该hooks用于确定当月是否有那么多天数和格式化月份
    const { chartMonth, MonthFormat } = this.getMonth(dataArr);
    // 取出员工信息
    const newExcelData = dataArr
      .filter(
        (item) =>
          item["姓名"] && !Object.values(item).some((value) => value === "淘汰")
      )
      .map((item) => {
        // 基础数据
        const baseData = {
          姓名: item["姓名"],
          直播游戏: item["所属项目"],
          基本工资: item["工资单价/天"] ?? "",
          主播分成比例: item["主播分成比例"],
        };
        // 对每个角色进行处理
        const nameMonth = this.handleMonth({ item, MonthFormat, chartMonth });
        const weeklyData = this.weeklyAttendance(nameMonth);
        const anchor = [];
        for (let p in weeklyData) {
          const newData = Object.assign({}, baseData, weeklyData[p], {
            应发工资: baseData["基本工资"] * weeklyData[p]["实出勤天数"] ?? 0,
          });
          anchor.push(newData);
        }
        return anchor;
      });
    // TODO：生成新的excel表
    this.createNewExcel(newExcelData);
  }
  // 周考勤数据处理
  weeklyAttendance(nameMonth) {
    const monthAttendance = [];
    for (let week = 0; week < nameMonth.length; week++) {
      const data = this.attendanceStatistics(nameMonth, week);
      monthAttendance.push(data);
    }
    return monthAttendance;
  }
  // 统计考勤次数
  attendanceStatistics(nameMonth, week) {
    const cycle = {};
    const weekLastDay = Object.keys(nameMonth[week]).length;
    cycle["核算工资周期"] = `${Object.keys(nameMonth[week][0])}/${Object.keys(
      nameMonth[week][weekLastDay - 1]
    )}`;
    cycle["应出勤天数"] = nameMonth[week].length;
    let count = 0;
    for (let i = 0; i < nameMonth[week].length; i++) {
      for (let j in nameMonth[week][i]) {
        if (nameMonth[week][i][j] === 1) {
          count++;
        }
      }
    }
    cycle["实出勤天数"] = count;
    return cycle;
  }
  handleMonth(time) {
    const { item, MonthFormat, chartMonth } = time;
    const Attendance = []; // 考勤统计
    let key = 0;
    let arr = [];
    let i = 0;
    Object.values(this.weeks).forEach((week, index) => {
      let day = {};
      if (i === 0) {
        ++i;
        day[`${MonthFormat}-${i}`] =
          item[MonthFormat.replace(/-/, "年") + "月"] === 1 ? 1 : 0;
      } else {
        day[`${MonthFormat}-${i}`] = item[`__EMPTY_${i - 1}`] === 1 ? 1 : 0;
      }
      i++;
      arr.push(day);
      if (week === "日") {
        Attendance.push(arr);
        key++;
        arr = [];
      }
      if (index === chartMonth - 1 && week !== "日") {
        Attendance.push(arr);
      }
    });
    return Attendance;
  }
  getMonth(month) {
    const weekArr = Object.keys(month[0]);
    let MonthFormat = weekArr[0].replace(/年/, "-").replace(/月/, "");
    // 截取表中月份(目的：确定当月是否有那么多天数)
    const chartMonth = dayjs(MonthFormat).daysInMonth();
    const monthObj = _.pick(month[1], weekArr);
    if (chartMonth !== Object.values(monthObj).length)
      throw new Error("月份不匹配");
    return { chartMonth, MonthFormat };
  }
  // 生成处理好的excel表
  createNewExcel(data) {
    if (!data[0]) return new ErrorEvent("excel表格数据无效！！");
    const workBook = xlsx.utils.book_new();
    let weekData = [];
    for (let i = 0; i < data[0].length; i++) {
      let temp = [];
      data.forEach((item) => temp.push(item[i]));
      weekData.push(temp);
      const sheetName = `第${i + 1}周`;
      const sheet = xlsx.utils.json_to_sheet(weekData[i]);
      xlsx.utils.book_append_sheet(workBook, sheet, sheetName);
    }
    xlsx.writeFile(workBook, "工资结算.xlsx");
  }
}
const eHandle = new excelHandle();
eHandle.readExcel();
