let file1Data = null;
let file2Data = null;

document.addEventListener("DOMContentLoaded", () => {
  const file1 = document.getElementById("file1");
  const file2 = document.getElementById("file2");
  const btn = document.getElementById("generateBtn");
  const status = document.getElementById("status");

  // 读取文件1
  file1.addEventListener("change", async (e) => {
    file1Data = await readExcel(e.target.files[0]);
    console.log("课程目录-1读取成功", file1Data);
    status.innerText = "课程目录-1读取成功";
  });

  // 读取文件2
  file2.addEventListener("change", async (e) => {
    file2Data = await readExcel(e.target.files[0]);
    console.log("课程目录-2读取成功", file2Data);
    status.innerText = "课程目录-2读取成功";
  });

  // 点击生成
  btn.addEventListener("click", () => {
    console.log("点击生成按钮");

    if (!file1Data || !file2Data) {
      alert("请先上传两个文件");
      return;
    }

    try {
      const result = generate(file1Data, file2Data);
      downloadExcel(result);
    } catch (err) {
      console.error(err);
      alert("生成失败：" + err.message);
    }
  });
});


// ========== 读取Excel ==========
function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });

        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);

        resolve(json);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}


// ========== 核心生成逻辑（占位但不会报错） ==========
function generate(list1, list2) {
  console.log("开始生成", list1, list2);

  // TODO：这里你后面可以继续升级规则逻辑
  // 当前先保证系统“能跑通”

  return list1.map((item, index) => {
    return {
      "课程名称(必填)": item["课程名称"] || "",
      "目录编码(必填)": "TEMP" + String(index + 1).padStart(4, "0"),
      "上级目录编码": "PARENT001"
    };
  });
}


// ========== 导出Excel ==========
function downloadExcel(data) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "课程目录-3");

  XLSX.writeFile(wb, "课程目录-3_结果.xlsx");
}
