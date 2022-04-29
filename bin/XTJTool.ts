#!/usr/bin/env node

import { Command } from "commander";
import { setConfig, transferJson, transferXlsx } from "../src/index";

const program = new Command();

program
  .version("1.0.0")
  .option("-x, --xlsx [boolean]", "将json文件转换为xlsx文件") // [boolean]，用中括号表示为可选参数；改用<boolean>, 则为必选
  .option("-j, --json [boolean]", "将xlsx文件转换为json文件")
  .option("-p, --path [string]", "文件夹路径; 默认为根目录")
  .option("-o, --output [string]", "文件夹路径; 默认为根目录")
  .option("-m, --merge [boolean]", "是否合并多份文件为一份")
  .parse(process.argv);

const options = program.opts();

setConfig({
  path: options.path,
  output: options.output,
  isMerge: options.merge,
});

if (options.xlsx || options.xlsx === "true") {
  transferJson();
} else if (options.json || options.json === "true") {
  transferXlsx();
}
