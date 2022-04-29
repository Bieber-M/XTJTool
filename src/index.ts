/**
 * 工具的使用
 *
 * 将满足要求格式的json文件，放到json文件夹
 *
 * 选择操作类型，1: translate; 2: generate
 *
 * generate操作，会根据json文件，生成对应的excel文件
 *
 * 库
 * https://github.com/SheetJS/sheetjs#installation
 * https://oss.sheetjs.com/sheetjs/tests/
 *
 * 参考：
 * https://www.cnblogs.com/liuxianan/p/js-excel.html
 */
import * as fs from "fs";
import * as path from "path";
import XLSX from "xlsx";

const rootDir = path.resolve(process.cwd());

let inputDirPath = `${rootDir}/`;
let outputDirPath = `${rootDir}/`;

// 生成的工作簿名
const wbName = "wb";
let wbIndex = 0;
let isMergeFile = false;

interface Config {
  path?: string;
  output?: string;
  isMerge?: boolean;
}

export const setConfig = (args: Config) => {
  const { path: dirPath, output, isMerge = false } = args;

  if (dirPath) {
    inputDirPath = outputDirPath = path.resolve(dirPath);
  }

  if (output) {
    outputDirPath = path.resolve(output);
  }

  isMergeFile = isMerge;
};

/**
 * 将json格式的数据格式化为符合xlsx表格格式的数据
 * @param data
 * @returns
 */
const formatJsonWithSheet = (data: Record<string, string>) => {
  const temp = [];
  // 键值对
  for (const i in data) {
    temp.push({
      col1: i,
      col2: data[i],
    });
  }

  return temp;
};

/**
 *
 * @param isMergeFile boolean 是否将多个json文件合并为一个json
 */
const importJson = async () => {
  const fileNameList = await fs.readdirSync(inputDirPath);

  let jsons: any;
  if (!isMergeFile) {
    let tempFileNames: string[] = [];
    jsons = [];

    fileNameList.forEach((fileName: string) => {
      // 排除隐藏文件和非json文件的影响
      if (/^\./.test(fileName) || !fileName.includes(".json")) {
        return;
      }
      tempFileNames.push(fileName);
      jsons.push(require(`${inputDirPath}/${fileName}`));
    });
  } else {
    jsons = {};
    fileNameList.forEach((fileName: string) => {
      // 排除隐藏文件的影响
      if (/^\./.test(fileName) || !fileName.includes(".json")) {
        return;
      }
      const jsonFile = require(`${inputDirPath}/${fileName}`);
      jsons = { ...jsonFile };
    });
  }

  return jsons;
};

/**
 *
 * @description 将json文件转换为xlsx文件
 */
const generateXlsWithJson = async (
  jsonData: Record<string, string>,
  sheetName = "sheet1"
) => {
  const formatJsonData = formatJsonWithSheet(jsonData);
  // 创建工作薄
  const wb = XLSX.utils.book_new();
  // 创建工作表
  const ws = XLSX.utils.json_to_sheet(formatJsonData, {
    header: ["col1", "col2"],
    skipHeader: true, // 跳过上面的标题行
  });
  // 将工作表添加近工作薄; 第三个参数为表名字
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  // 将工作簿转换为对应类型的数据
  const result = XLSX.write(wb, {
    bookType: "xlsx", // 输出的文件类型
    type: "buffer", // 输出的数据类型
    compression: true, // 开启zip压缩
  });
  // 生成xlsx
  fs.writeFileSync(`${outputDirPath}/${wbName}-${wbIndex}.xlsx`, result);
};

// json to excel
export const transferJson = () => {
  importJson().then((jsonData) => {
    // console.log('type: ', Object)
    if (jsonData instanceof Array) {
      jsonData.forEach((jd) => {
        generateXlsWithJson(jd);
        wbIndex++;
      });
      wbIndex = 0;
    } else {
      generateXlsWithJson(jsonData);
    }
  });
};

/**
 *
 * @param isMerge
 * @returns
 */
const importXlsx = async (isMerge = true) => {
  const fileNameList = await fs.readdirSync(inputDirPath);

  const wbData: XLSX.WorkBook[] = [];

  fileNameList.forEach((fileName: string) => {
    // 排除隐藏文件的影响
    if (/^\./.test(fileName) || !fileName.includes(".xlsx")) {
      return;
    }
    wbData.push(
      XLSX.read(fs.readFileSync(`${inputDirPath}/${fileName}.xlsx`, "binary"), {
        type: "binary",
      })
    );
  });

  if (isMergeFile) {
    const newWb = XLSX.utils.book_new();
    wbData.forEach((wb) => {
      wb.SheetNames.forEach((sn) => {
        XLSX.utils.book_append_sheet(newWb, wb.Sheets[sn]);
      });
    });

    return newWb;
  } else {
    return wbData;
  }
};

/**
 * 将xlsx文件转换为json文件
 * @param jsonData
 * @param fileName 输出的xlsx名
 */
const generateJsonWithXlsx = (
  jsonData: Record<string, string>,
  fileName = "json"
) => {
  //将 JSON 对象转换为 String
  var jsonStr = JSON.stringify(jsonData);

  //将json字符串读取到缓冲区
  const buf = Buffer.from(jsonStr);

  fs.writeFileSync(`${outputDirPath}/${fileName}.json`, buf);
};

// excel to json
export const transferXlsx = () => {
  importXlsx().then((data) => {
    const generateJson = (wb: XLSX.WorkBook) => {
      const jsonDataArray: any[] = [];

      wb.SheetNames.forEach((sn) => {
        const ws = wb.Sheets[sn];
        const jsonArray = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          raw: true,
        });

        // sheet表的左侧有数据的第一列为json键名，第二列为键值
        let jsonData: any = {};
        jsonArray.forEach((ja: any) => {
          if (ja && ja.length) {
            jsonData[ja[0]] = ja[1];
          }
        });
        jsonDataArray.push({
          sheetName: sn,
          sheetData: jsonData,
        });
      });

      jsonDataArray.forEach((jd) => {
        generateJsonWithXlsx(jd.sheetData, jd.sheetName);
      });

      return jsonDataArray;
    };

    if (data instanceof Array) {
      data.forEach((wb) => {
        generateJson(wb);
      });
    } else {
      generateJson(data);
    }
  });
};
