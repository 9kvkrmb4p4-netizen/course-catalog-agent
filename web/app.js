console.log("🔥 课程目录系统已加载");

let file1 = [];
let file2 = [];
let file3 = [];
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
   file1
========================= */
document.getElementById("file1").addEventListener("change", async (e) => {
  file1 = await readExcel(e.target.files[0]);
  console.log("✔ file1", file1);
});

/* =========================
   file2（构建树结构 + 上级编码）
========================= */
document.getElementById("file2").addEventListener("change", async (e) => {

  file2 = await readExcel(e.target.files[0]);

  treeMap = {};

  file2.forEach(row => {

    const name = row["目录名称(必填)"] || row["目录名称"] || "";
    const code = row["目录编码(必填)"] || row["目录编码"] || "";
    const parent = row["上级目录编码"] || "";

    if (!treeMap[parent]) treeMap[parent] = [];

    treeMap[parent].push({
      name,
      code
    });
  });

  console.log("✔ treeMap", treeMap);
});

/* =========================
   file3（可选模板）
========================= */
document.getElementById("file3").addEventListener("change", async (e) => {
  file3 = await readExcel(e.target.files[0]);
  console.log("✔ file3", file3);
});

/* =========================
   生成按钮
========================= */
document.getElementById("runBtn").addEventListener("click", generate);

/* =========================
   核心生成逻辑（目录树编码）
========================= */
function generate() {

  if (!file1.length || !file2.length) {
    alert("请上传 file1 和 file2");
    return;
  }

  result = [];

  file1.forEach((row, index) => {

    const course = row["课程名称"] || Object.values(row)[0];
    const l1 = row["一级分类"];
    const l2 = row["二级分类"];

    const parentCode = findParentCode(l1, l2);

    const newCode = generateNextCode(parentCode, index);

    result.push({
      "目录名称(必填)": course,
      "目录编码(必填)": newCode,
      "上级目录编码": parentCode
    });
  });

  console.log("🎯 result", result);

  exportExcel(result);
}

/* =========================
   找上级编码（按file2树）
========================= */
function findParentCode(l1, l2) {

  for (let key in treeMap) {
    const nodes = treeMap[key];

    for (let n of nodes) {
      if (n.name === l1 || n.name === l2) {
        return n.code;
      }
    }
  }

  return "未匹配";
}

/* =========================
   生成连续编码
========================= */
function generateNextCode(parentCode, index) {

  if (!parentCode || parentCode === "未匹配") return "未匹配";

  const num = String(index + 1).padStart(2, "0");

  return parentCode + num;
}

/* =========================
   导出Excel
========================= */
function exportExcel(data) {

  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "课程目录-3");

  XLSX.writeFile(wb, "课程目录-3_结果.xlsx");

  document.getElementById("msg").innerText = "生成完成 ✔";
}
