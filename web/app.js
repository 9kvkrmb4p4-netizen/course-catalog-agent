let data1 = [];
let data2 = [];
let ruleMap = [];
let result = [];

/* =========================
   文件1：课程目录-1
========================= */
document.getElementById("file1").addEventListener("change", function (e) {
    readExcel(e, 1);
});

/* =========================
   文件2：课程目录-2（规则表）
========================= */
document.getElementById("file2").addEventListener("change", function (e) {
    readExcel(e, 2);
});

/* =========================
   读取Excel
========================= */
function readExcel(e, type) {
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function (evt) {
        let wb = XLSX.read(evt.target.result, { type: "binary" });
        let sheet = wb.Sheets[wb.SheetNames[0]];
        let data = XLSX.utils.sheet_to_json(sheet);

        if (type === 1) {
            data1 = data;
            console.log("✅ 课程目录-1读取成功", data1);
        }

        if (type === 2) {
            data2 = data;
            ruleMap = buildRuleMap(data2);
            console.log("✅ 课程目录-2读取成功", ruleMap);
        }
    };

    reader.readAsBinaryString(file);
}

/* =========================
   构建规则Map（核心修复版）
   完全按字段名匹配（不会错）
========================= */
function buildRuleMap(data) {

    let map = {};

    data.forEach(row => {

        let name =
            row["目录名称(必填)"] ||
            row["目录名称"] ||
            row["名称"] ||
            row["课程名称"];

        let code =
            row["目录编码(必填)"] ||
            row["目录编码"] ||
            row["编码"];

        if (name && code) {
            map[String(name).trim()] = String(code).trim();
        }
    });

    return map;
}

/* =========================
   生成逻辑（核心）
========================= */
function generate() {

    if (!data1.length) {
        alert("请先上传课程目录-1");
        return;
    }

    if (!data2.length) {
        alert("❌ 请先上传课程目录-2（规则表）");
        return;
    }

    result = [];

    data1.forEach(row => {

        let course =
            row["课程名称"] ||
            row["课程名称(必填)"] ||
            row["名称"] ||
            Object.values(row)[0];

        let l1 =
            row["一级分类"] ||
            row["一级目录"] ||
            Object.values(row)[1];

        let l2 =
            row["二级分类"] ||
            row["二级目录"] ||
            Object.values(row)[2];

        // 精确匹配 + 兜底匹配
        let code =
            ruleMap[l1] ||
            ruleMap[l1 + "->" + l2] ||
            "未匹配";

        result.push({
            "目录名称(必填)": course || "",
            "目录编码(必填)": code,
            "上级目录编码": l2 || ""
        });
    });

    console.log("🎯 最终结果", result);

    download();
}

/* =========================
   下载Excel
========================= */
function download() {

    let ws = XLSX.utils.json_to_sheet(result);
    let wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "课程目录");

    XLSX.writeFile(wb, "课程目录-输出.xlsx");
}
