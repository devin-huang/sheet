# 自定义 excel 格式 导入导出

![](https://raw.githubusercontent.com/devin-huang/sheet/main/demo.png)

- JS 控制格式（单元格合并，样式）

## 重点注意：

- `./cptable` 是xslx-style的依赖引起报错，所以必须要在webpack中; `fs: 'empty'` 是让项目支持引入fs

  ```
    防止将特定打包，而是在运行时从外部获取
    externals: {
      './cptable': 'var cptable',
    },
    node: {
      fs: 'empty',
    }
  ```

## base

- 依赖 [`xlsx`](https://www.npmjs.com/package/xlsx) [`xlsx-style`](https://www.npmjs.com/package/xlsx-style) [`file-saver`](https://www.npmjs.com/package/file-saver) 实现

- 支持导出格式： `csv` `xls` `xlsx` `json` `html`

- xlsx-style [单元格样式配置](https://www.jianshu.com/p/869375439fee)

### Compiles and hot-reloads for development

```
npm run serve
```

### Compiles and minifies for production

```
npm run build
```

### Lints and fixes files

```
npm run lint
```

### Customize configuration

See [Configuration Reference](https://cli.vuejs.org/config/).
