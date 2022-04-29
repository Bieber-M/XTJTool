### 功能

1.将满足对应格式的 json 文件转换为 xlsx 文件；

2.将满足对应格式的 xlsx 文件转换为 json 文件；

```json
// demo.json
{
  "key1": "value1",
  "key2": "value2",
  "key3": "value3"
}
```

上面的 json 会被转换成下面的 xlsx 表
表头|表头
---|:--
key1|value1
key2|value2

同样，满足如此格式的 xlsx 表文件也可转换为 json 文件

### 用法

```js
// 安装
yarn add xtjtool -D

// 常用命令
XTJ -h
```
