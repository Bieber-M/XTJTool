#!/usr/bin/env node
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const commander_1 = require("commander");
const index_1 = require("../src/index");
const program = new commander_1.Command();
program
    .version("1.0.0")
    .option("-x, --xlsx [boolean]", "将json文件转换为xlsx文件") // [boolean]，用中括号表示为可选参数；改用<boolean>, 则为必选
    .option("-j, --json [boolean]", "将xlsx文件转换为json文件")
    .option("-p, --path [string]", "文件夹路径; 默认为根目录")
    .option("-o, --output [string]", "文件夹路径; 默认为根目录")
    .option("-m, --merge [boolean]", "是否合并多份文件为一份")
    .parse(process.argv);
const options = program.opts();
(0, index_1.setConfig)({
    path: options.path,
    output: options.output,
    isMerge: options.merge,
});
if (options.xlsx || options.xlsx === "true") {
    (0, index_1.transferJson)();
}
else if (options.json || options.json === "true") {
    (0, index_1.transferXlsx)();
}
