console.log("🔥 JS已加载成功");

// 全局变量
let data1 = [];
let data2 = [];

// 读取Excel
function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        resolve(json);
      } catch (err) {
        reject(err);
      }
    };

    reader.readAsArrayBuffer(file);
  });
}

// 主函数（关键修复点：必须挂到 window）
window.generate = async function () {

  const file1 = document.getElementById("file1").files[0];
  const file2 = document.getElementById("file2").files[0];

  if (!file1 || !file2) {
    alert("请先上传两个文件");
    return;
  }

  console.log("📥 开始读取文件...");

  try {
    data1 = await readExcel(file1);
    data2 = await readExcel(file2);

    console.log("课程目录-1读取成功", data1);
    console.log("课程目录-2读取成功", data2);

    // ====== 核心逻辑（修复版）======
    const map = {};

    // 规则表建立映射（上级编码）
    data2.forEach(row => {
      const name = row["目录名称"] || row["名称"];
      const code = row["目录编码"];
      if (name && code) {
        map[name.trim()] = code;
      }
    });

    let result = [];

    data1.forEach((row, index) => {

      const courseName = row["课程名称"] || "";
      const level1 = row["一级分类"];
      const level2 = row["二级分类"];

      const parentCode = map[level2] || "未匹配";

      const code = "TY" + String(1000 + index + 1);

      result.push({
        "目录名称(必填)": courseName,
        "目录编码(必填)": code,
        "上级目录编码": parentCode
      });
    });

    // 输出Excel
    const ws = XLSX.utils.json_to_sheet(result);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "课程目录-3");

    XLSX.writeFile(wb, "课程目录-3.xlsx");

    alert("生成成功！");

  } catch (err) {
    console.error("❌ 错误:", err);
    alert("处理失败，请看控制台");
  }
};
