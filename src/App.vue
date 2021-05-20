<template>
  <div id="app">
    <img
      alt="Vue logo"
      class="devin-logo"
      src="https://avatars.githubusercontent.com/u/29721413?v=4"
    />
    <div class="container">
      <div class="import">
        <h4>导入EXCEL并显示</h4>
        <form method="post" enctype="multipart/form-data">
          <input type="file" name="file" id="upload" />
          <button @click="importData">导入</button>
        </form>

        <div id="excelView"></div>
      </div>
      <div class="export">
        <h4>导出EXCEL</h4>
        <small>导出JS控制自定义格式 </small>
        <button @click="exportStudent">导出</button>
      </div>
    </div>
  </div>
</template>

<script>
import { mockData, mockHeader } from "./utils/mockData";
import { exportExcel, impoerExcel } from "./utils/xlsxUtils.js";

export default {
  name: "App",
  methods: {
    importData() {
      impoerExcel("#upload", "#excelView");
    },
    async exportStudent() {
      this.exportLoading = true;
      try {
        const res = {
          data: [...mockData],
        };
        res.data.forEach((item) => {
          item.courseStr = item.courses.map((o) => o.name).join("，");
          item.classStr = item.classes.map((o) => o.name).join("，");
        });
        await exportExcel(res.data, mockHeader, "学员信息表.xlsx", {
          A2: {
            fill: { fgColor: { rgb: "409F9F80" } },
          },
        });
      } catch (error) {
        console.log(error);
      }
      this.exportLoading = false;
    },
  },
};
</script>

<style scoped>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
.container {
  display: flex;
  justify-content: space-around;
  align-items: stretch;
}
.devin-logo {
  border-radius: 50%;
}
.import {
  background: #ffded2;
  padding: 20px;
  flex: 0.6;
}
.export {
  background: #b2ffdb;
  padding: 20px;
  flex: 0.4;
}
</style>
