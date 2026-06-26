let data1 = [];
let data2 = [];
let ruleMap = {};
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
   Excel读取（已修复Netlify兼容问题）
========================= */
function readExcel(e, type) {

    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function (evt) {

        let data = new Uint8Array(evt.target.result);

        let wb = XLSX.read(data, { type: "array" });

        let sheet = wb.Sheets[wb.SheetNames[0]];

        let json = XLSX.utils.sheet_to_json(sheet);

        if (type === 1) {
            data1 = json;
            console.log("✅ 课程目录-1读取成功", data1);
        }

        if (type === 2) {
            data2 = json;
            ruleMap = buildRuleMap(data2);
            console.log("✅ 课程目录-2读取成功", ruleMap);
        }
    };

    reader.readAsArrayBuffer(file);
}

/* =========================
   构建规则Map（完全字段兼容）
========================= */
function buildRuleMap(data) {

    let map = {};

    data.forEach(row => {

        let name =
            row["目录名称(必填)"] ||
            row["目录名称"] ||
            row["名称"] ||
            row["课程名称"] ||
            Object.values(row)[0];

        let code =
            row["目录编码(必填)"] ||
            row["目录编码"] ||
            row["编码"] ||
            Object.values(row)[1];

        if (name && code) {
            map[String(name).trim()] = String(code).trim();
        }
    });

    return map;
}

/* =========================
   生成逻辑
========================= */
function generate() {

    if (!data1.length) {
        alert("请先上传课程目录-1");
        return;
    }

    if (!data2.length) {
        alert("❌ 课程目录-2未成功读取，请重新上传");
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

    console.log("🎯 生成结果", result);

    download();
}

/* =========================
   下载Excel
========================= */
function download() {

    let ws = XLSX.utils.json_to_sheet(result);
    let wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "课程目录输出");

    XLSX.writeFile(wb, "课程目录-输出.xlsx");
}
