"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.transferXlsx = exports.transferJson = exports.setConfig = void 0;
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
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
const xlsx_1 = __importDefault(require("xlsx"));
const rootDir = path.resolve(process.cwd());
let inputDirPath = `${rootDir}/`;
let outputDirPath = `${rootDir}/`;
// 生成的工作簿名
const wbName = "wb";
let wbIndex = 0;
let isMergeFile = false;
const setConfig = (args) => {
    const { path: dirPath, output, isMerge = false } = args;
    if (dirPath) {
        inputDirPath = outputDirPath = path.resolve(dirPath);
    }
    if (output) {
        outputDirPath = path.resolve(output);
    }
    isMergeFile = isMerge;
};
exports.setConfig = setConfig;
/**
 * 将json格式的数据格式化为符合xlsx表格格式的数据
 * @param data
 * @returns
 */
const formatJsonWithSheet = (data) => {
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
const importJson = () => __awaiter(void 0, void 0, void 0, function* () {
    const fileNameList = yield fs.readdirSync(inputDirPath);
    let jsons;
    if (!isMergeFile) {
        let tempFileNames = [];
        jsons = [];
        fileNameList.forEach((fileName) => {
            // 排除隐藏文件和非json文件的影响
            if (/^\./.test(fileName) || !fileName.includes(".json")) {
                return;
            }
            tempFileNames.push(fileName);
            jsons.push(require(`${inputDirPath}/${fileName}`));
        });
    }
    else {
        jsons = {};
        fileNameList.forEach((fileName) => {
            // 排除隐藏文件的影响
            if (/^\./.test(fileName) || !fileName.includes(".json")) {
                return;
            }
            const jsonFile = require(`${inputDirPath}/${fileName}`);
            jsons = Object.assign({}, jsonFile);
        });
    }
    return jsons;
});
/**
 *
 * @description 将json文件转换为xlsx文件
 */
const generateXlsWithJson = (jsonData, sheetName = "sheet1") => __awaiter(void 0, void 0, void 0, function* () {
    const formatJsonData = formatJsonWithSheet(jsonData);
    // 创建工作薄
    const wb = xlsx_1.default.utils.book_new();
    // 创建工作表
    const ws = xlsx_1.default.utils.json_to_sheet(formatJsonData, {
        header: ["col1", "col2"],
        skipHeader: true, // 跳过上面的标题行
    });
    // 将工作表添加近工作薄; 第三个参数为表名字
    xlsx_1.default.utils.book_append_sheet(wb, ws, sheetName);
    // 将工作簿转换为对应类型的数据
    const result = xlsx_1.default.write(wb, {
        bookType: "xlsx",
        type: "buffer",
        compression: true, // 开启zip压缩
    });
    // 生成xlsx
    fs.writeFileSync(`${outputDirPath}/${wbName}-${wbIndex}.xlsx`, result);
});
// json to excel
const transferJson = () => {
    importJson().then((jsonData) => {
        // console.log('type: ', Object)
        if (jsonData instanceof Array) {
            jsonData.forEach((jd) => {
                generateXlsWithJson(jd);
                wbIndex++;
            });
            wbIndex = 0;
        }
        else {
            generateXlsWithJson(jsonData);
        }
    });
};
exports.transferJson = transferJson;
/**
 *
 * @param isMerge
 * @returns
 */
const importXlsx = (isMerge = true) => __awaiter(void 0, void 0, void 0, function* () {
    const fileNameList = yield fs.readdirSync(inputDirPath);
    const wbData = [];
    fileNameList.forEach((fileName) => {
        // 排除隐藏文件的影响
        if (/^\./.test(fileName) || !fileName.includes(".xlsx")) {
            return;
        }
        wbData.push(xlsx_1.default.read(fs.readFileSync(`${inputDirPath}/${fileName}.xlsx`, "binary"), {
            type: "binary",
        }));
    });
    if (isMergeFile) {
        const newWb = xlsx_1.default.utils.book_new();
        wbData.forEach((wb) => {
            wb.SheetNames.forEach((sn) => {
                xlsx_1.default.utils.book_append_sheet(newWb, wb.Sheets[sn]);
            });
        });
        return newWb;
    }
    else {
        return wbData;
    }
});
/**
 * 将xlsx文件转换为json文件
 * @param jsonData
 * @param fileName 输出的xlsx名
 */
const generateJsonWithXlsx = (jsonData, fileName = "json") => {
    //将 JSON 对象转换为 String
    var jsonStr = JSON.stringify(jsonData);
    //将json字符串读取到缓冲区
    const buf = Buffer.from(jsonStr);
    fs.writeFileSync(`${outputDirPath}/${fileName}.json`, buf);
};
// excel to json
const transferXlsx = () => {
    importXlsx().then((data) => {
        const generateJson = (wb) => {
            const jsonDataArray = [];
            wb.SheetNames.forEach((sn) => {
                const ws = wb.Sheets[sn];
                const jsonArray = xlsx_1.default.utils.sheet_to_json(ws, {
                    header: 1,
                    raw: true,
                });
                // sheet表的左侧有数据的第一列为json键名，第二列为键值
                let jsonData = {};
                jsonArray.forEach((ja) => {
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
        }
        else {
            generateJson(data);
        }
    });
};
exports.transferXlsx = transferXlsx;
