console.log("🔥 JS已加载成功");

let file1 = [];
let file2 = [];
let treeMap = {};
let result = [];

/* =========================
   Excel读取
========================= */
function readExcel(file) {
  return new Promise((resolve) => {

    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      resolve(json);
    };

    reader.readAsArrayBuffer(file);
  });
}

/* =========================
   初始化（关键）
========================= */
window.onload = function () {

  console.log("✔ 页面加载完成");

  const file1Input = document.getElementById("file1");
  const file2Input = document.getElementById("file2");
  const btn = document.getElementById("runBtn");

  if (!file1Input || !file2Input || !btn) {
    console.error("❌ HTML元素缺失（检查id）");
    return;
  }

  /* file1 */
  file1Input.addEventListener("change", async (e) => {
    file1 = await readExcel(e.target.files[0]);
    console.log("✔ file1", file1);
  });

  /* file2 */
  file2Input.addEventListener("change", async (e) => {

    file2 = await readExcel(e.target.files[0]);

    treeMap = {};

    file2.forEach(r => {
      const name = Object.values(r)[0];
      const code = Object.values(r)[1];
      const parent = Object.values(r)[2] || "";

      if (!treeMap[parent]) treeMap[parent] = [];

      treeMap[parent].push({ name, code });
    });

    console.log("✔ treeMap", treeMap);
  });

  /* 绑定按钮（关键修复点） */
  btn.addEventListener("click", generate);
};

/* =========================
   生成逻辑
========================= */
function generate() {

  console.log("🚀 点击成功触发 generate");

  if (!file1.length || !file2.length) {
    alert("请先上传两个Excel文件");
    return;
  }

  result = file1.map((row, i) => {

    const course = Object.values(row)[0];
    const l1 = Object.values(row)[1];
    const l2 = Object.values(row)[2];

    let parentCode = "未匹配";

    for (let key in treeMap) {
      for (let n of treeMap[key]) {
        if (n.name === l1 || n.name === l2) {
          parentCode = n.code;
        }
      }
    }

    const code = parentCode !== "未匹配"
      ? parentCode + String(i + 1).padStart(2, "0")
      : "未匹配";

    return {
      "目录名称(必填)": course,
      "目录编码(必填)": code,
      "上级目录编码": parentCode
    };
  });

  console.log("🎯 result", result);

  const ws = XLSX.utils.json_to_sheet(result);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "课程目录3");

  XLSX.writeFile(wb, "课程目录3.xlsx");
}
