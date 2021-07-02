import XLSX from "xlsx";
import XLSXStyle from "xlsx-style";
import { saveAs } from "file-saver";
import path from "path";
import {
  pickOne,
  assignAll,
  stringToNum,
  s2ab,
  isPlainObject,
  setMergeData,
} from "./tools";

const FILE_NAME = "表格.xlsx";
const COL_PARAMS = ["hidden", "wpx", "width", "wch", "MDW"];
const STYLE_PARAMS = ["fill", "font", "alignment", "border"];
class Sheet {
  constructor() {
    this.wbOut = null;
    this.ws = null;
    this.wb = null;
    this.header = [];
    this.bookType = null;
    this.data = [];
  }

  /**
   * 导入
   * @param {Object} selector 上传DOM对象
   * @param {Object} displayEle 显示EXCEL区域DOM
   */
  import = (selector, displayEle) => {
    event.preventDefault();
    var [files] = document.querySelector(selector).files;

    if (files) {
      let reader = new FileReader();
      const XLS_TYPE = "application/vnd.ms-excel";
      const XLSX_TYPE =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      const { size, type } = files;

      if (type.includes(XLS_TYPE) || type.includes(XLSX_TYPE)) {
        if (size < 1024 * 1000 * 5) {
          reader.readAsArrayBuffer(files);
          reader.onload = function(event) {
            const data = event.target.result;
            const { SheetNames, Sheets } = XLSX.read(data, { type: "array" });
            const [wsName] = SheetNames;
            const ws = Sheets[wsName];
            // console.log(XLSX.utils.sheet_to_html(ws));
            // 渲染
            document.querySelector(
              displayEle
            ).innerHTML = XLSX.utils.sheet_to_html(ws);
          };
        } else {
          alert("必须小于5MB");
        }
      } else {
        alert("格式必须为xlsx, xls");
      }
    } else {
      alert("请选择上传文件");
    }
  };

  /**
   * 导出成Excel
   * @param {Array} data 数据
   * @param {Object} columns 每列参数说明。可直接写对应的中文名称，也可以写这一列的样式。比如列宽，背景色
   * @param {String} filename 文件名。可根据文件名后缀动态判断文件格式。支持xlsx, xlsm, xlsb（这三个支持自定义样式）, html, csv等
   * @param {Object} styleConf 指定单元格的样式
   * */
  export = (data, columns, filename = FILE_NAME, styleConf) => {
    // 获取自定义表头属性名
    let keys = Object.keys(columns);

    // 将数组转为对象
    const colNames = this.getObjectByColAttr(keys, columns);

    // 新建表
    this.wb = XLSX.utils.book_new();

    // 插入表头
    this.ws = this.createSheetHeader([colNames], 1);

    // 过滤数据，只显示表头包含的数据
    this.data = data.map((item) => pickOne(keys, item));

    // 插入标题 (属性值：name是根据数据)
    this.insertContentByJson([{ name: "***经营情况" }]);

    // 插入表头
    this.insertContentByJson(this.data, 2);

    // 生成表
    this.creatBook();

    // 设置导出类型
    this.setFileType(filename);

    // 自定义渲染样式并导出
    this.renderStyleAndExport(columns, filename, styleConf);
  };

  renderStyleAndExport = (columns, filename, styleConf) => {
    let keys = Object.keys(columns);
    // 选择以哪个字段合并
    const getMergeData = setMergeData("idCard");

    if (["xlsx", "xlsm", "xlsb"].includes(this.bookType)) {
      // 设置合并数量
      const setMergeColmuns = this.mergeEachByCols(30);

      // 获取合并的键值对应关系
      const mergeData = getMergeData(this.data);

      // 获取将键值对的值
      const mergeNumArr = Object.values(mergeData);

      // 设置table中需要合并的数据
      const getMergeColumns = () => {
        let curPositionX = 2;
        let curPositionY = 0;
        let allData = [];

        mergeNumArr.forEach((quantity) => {
          quantity = Number(quantity);
          if (quantity > 1) {
            curPositionY = curPositionX + (quantity - 1);
            allData = [
              ...allData,
              ...setMergeColmuns(curPositionX, curPositionY),
            ];
          }
          // 合并单元格后再赋值给开始行
          curPositionX = curPositionY + 1;
        });

        //  (属性值：name是根据数据)
        this.insertContentByJson(
          [
            { name: "取数来源" },
            { name: "**基本信息，来源于合同查询，（**有同步后台）" },
            { name: "实收金额：柜组销售月报" },
            {
              name: "创利额：BI品类经营结果对比分析",
            },
          ],
          curPositionY + 5
        );
        // 5：空五行显示Note
        const note = [1, 2, 3, 4].map((_, index) => {
          const sRow = curPositionY + 5 + Number(index);
          return this.mergeCell({ eRow: sRow, sRow, eCol: 5 })[0];
        });

        return { allData, note };
      };

      // 表头合并单元格
      const headerMerge = this.mergeCell({ eCol: 40 });

      const { allData, note } = getMergeColumns();
      const [first] = note;
      const {
        s: { r: rowByNote },
      } = first;

      // 合并所有单元格
      this.ws["!merges"] = [...headerMerge, ...allData, ...note];

      // 样式对象
      const styleObj = {
        1: {
          express: (key) => key.replace(/[^0-9]/gi, "") === "1",
          value: {
            fill: {
              fgColor: { rgb: "87CEEB" },
            },
            font: {
              name: "宋体",
              sz: 12,
              bold: true,
            },
            // border: {
            //   bottom: {
            //     style: "thin",
            //     color: "FF000000",
            //   },
            // },
            alignment: {
              wrapText: true,
              // horizontal: "center",
              vertical: "center",
            },
          },
        },
        [rowByNote + 1]: {
          express: (key) =>
            key.replace(/[^0-9]/gi, "") === String(rowByNote + 1),
          value: {
            font: {
              bold: true,
            },
          },
        },
        [rowByNote]: {
          express: (key) =>
            Number(key.replace(/[^0-9]/gi, "")) < Number(rowByNote),
          value: (key) => {
            let str = key.replace(/[^A-Za-z]+$/gi, "");
            let colIndex = stringToNum(str) - 1;
            if (keys[colIndex]) {
              let a = {};
              const style = {
                alignment: {
                  wrapText: true,
                  horizontal: "center",
                  vertical: "center",
                },
              };
              if (isPlainObject(columns[keys[colIndex]])) {
                a = pickOne(STYLE_PARAMS, columns[keys[colIndex]]);
              }

              this.ws[key].s = assignAll(this.ws[key].s, a, style);
            }
          },
        },
      };

      // 设置每行的样式
      this.setStyleEachRow(styleObj);

      // 设置列宽度
      this.setColumnWid(columns);

      // 设置其他个别单元格的样式
      this.mergeOtherStyle(styleConf);

      // 导出
      this.saveData(filename);
    } else {
      this.wbOut = XLSX.write(this.wb, {
        bookType: this.bookType,
        bookSST: false,
        type: "binary",
      });
    }
  };

  // 将数组转为对象
  getObjectByColAttr = (columns, column) => {
    return columns.reduce((prev, key) => {
      if (isPlainObject(column[key])) {
        return { ...prev, ...{ [key]: column[key].name } };
      } else {
        return { ...prev, ...{ [key]: column[key] } };
      }
    }, {});
  };

  // 设置列宽
  setColumnWid = (columns) => {
    const colsP = [];
    Object.values(columns).forEach((item) => {
      colsP.push(pickOne(COL_PARAMS, item));
    });
    this.ws["!cols"] = colsP;
  };

  // 合并其他样式参数
  mergeOtherStyle = (styleConf) => {
    if (styleConf) {
      for (const key in styleConf) {
        if (Object.prototype.hasOwnProperty.call(this.ws, key)) {
          this.ws[key].s = styleConf[key];
        }
      }
    }
    this.wbOut = XLSXStyle.write(this.wb, {
      bookType: this.bookType,
      bookSST: false,
      type: "binary",
    });
  };

  /**
   * 格式：
   * {
   *  1: { express, value }
   * }
   * 遍历设置表格样式
   */
  setStyleEachRow = (styleObj) => {
    for (const o in styleObj) {
      for (const key in this.ws) {
        if (Object.prototype.hasOwnProperty.call(styleObj[o], "express")) {
          // 对象值为函数时处理
          if (
            styleObj[o].value &&
            typeof styleObj[o].value === "function" &&
            styleObj[o].express(key)
          ) {
            styleObj[o].value.call(this, key);
          } else {
            // 基础配置样式
            if (
              Object.prototype.hasOwnProperty.call(this.ws[key], "v") &&
              styleObj[o].express(key)
            ) {
              this.ws[key].s = styleObj[o].value;
            }
          }
        }
      }
    }
  };

  // 保存自定义的数据（含样式）并输出
  saveData = (filename) => {
    saveAs(new Blob([s2ab(this.wbOut)], { type: "" }), filename);
  };

  // 设置每行中需要合并的列数
  mergeEachByCols = (quantity) => {
    // 需合并列数
    const mergeColumns = new Array(quantity).fill();
    return (x, y) =>
      mergeColumns.reduce(
        (prev, _, key) => [
          ...prev,
          ...this.mergeCell({ sCol: key, eCol: key, sRow: x, eRow: y }),
        ],
        []
      );
  };

  // 合并单元格
  mergeCell = ({ sRow = 0, sCol = 0, eRow = 0, eCol = 0 }) => {
    return [
      {
        s: {
          // s为开始
          c: sCol, // 开始列
          r: sRow, // 开始行
        },
        e: {
          // e结束
          c: eCol, // 结束列
          r: eRow, // 结束行
        },
      },
    ];
  };

  // 设置导出文件类型
  setFileType = (filename) => {
    let ext = path.extname(filename);
    if (ext == null) {
      filename += ".xlsx";
      this.bookType = "xlsx";
    } else {
      this.bookType = ext.substr(1).toLowerCase();
    }
  };

  // 创建表
  creatBook = (name = "sheet1") => {
    // this.wb = XLSX.utils.book_new()
    this.wb.SheetNames.push(name);
    this.wb.Sheets[name] = this.ws;
  };

  // 创建表头
  createSheetHeader = (data, index = 0) => {
    if (!Array.isArray(data)) {
      throw "数据必须为数组类型";
    }
    return XLSX.utils.json_to_sheet(data, {
      header: this.header,
      skipHeader: true,
      origin: index, // 索引从0开始
    });
  };

  // 插入内容
  insertContentByJson = (data, index = 0) => {
    if (!Array.isArray(data)) {
      throw "数据必须为数组类型";
    }
    XLSX.utils.sheet_add_json(this.ws, data, {
      header: this.header,
      skipHeader: true,
      origin: index,
    });
  };
}

export default new Sheet();
