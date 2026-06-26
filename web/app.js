let file1Data = [];
let ruleMap = {};

/* =========================
   读取 Excel（稳定版）
========================= */
function readExcel(file) {
  return new Promise((resolve, reject) => {

    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const json = XLSX.utils.sheet_to_json(sheet, {
          defval: "",
          raw: true
        });

        resolve(json);

      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/* =========================
   file1
========================= */
document.getElementById("file1").addEventListener("change", async (e) => {
  file1Data = await readExcel(e.target.files[0]);
  console.log("课程目录-1 ✔", file1Data);
});

/* =========================
   file2（关键修复）
========================= */
document.getElementById("file2").addEventListener("change", async (e) => {

  const file = e.target.files[0];

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const arr = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  console.log("规则表二维数组 ✔", arr);

  ruleMap = {};

  for (let i = 1; i < arr.length; i++) {
    const row = arr[i];

    const name = (row[0] || "").toString().trim();
    const code = (row[1] || "").toString().trim();

    if (name && code) {
      ruleMap[name] = code;
    }
  }

  console.log("规则Map ✔", ruleMap);
});

/* =========================
   生成逻辑
========================= */
function generate() {

  if (!file1Data.length) {
    alert("请先上传课程目录-1");
    return;
  }

  if (Object.keys(ruleMap).length === 0) {
    alert("请先上传课程目录-2（规则表）");
    return;
  }

  const result = file1Data.map(row => {

    const name = row["课程名称"] || row["课程名称 "] || Object.values(row)[0];

    const code = ruleMap[name] || "未匹配";

    return {
      "目录名称(必填)": name,
      "目录编码(必填)": code,
      "上级目录编码": code !== "未匹配" ? code.slice(0, -2) : "未匹配"
    };
  });

  console.log("最终结果 ✔", result);

  const ws = XLSX.utils.json_to_sheet(result);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "结果");

  XLSX.writeFile(wb, "课程目录生成结果.xlsx");

  document.getElementById("msg").innerText = "生成完成 ✔";
}
