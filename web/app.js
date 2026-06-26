let data1 = [];
let ruleMap = {};
let result = [];

/* =========================
   文件1：课程目录-1
========================= */
document.getElementById("file1").addEventListener("change", readFile1);

function readFile1(e) {
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function (evt) {
        let wb = XLSX.read(evt.target.result, { type: "binary" });
        let sheet = wb.Sheets[wb.SheetNames[0]];
        data1 = XLSX.utils.sheet_to_json(sheet);
    };

    reader.readAsBinaryString(file);
}

/* =========================
   文件2：课程目录-2（规则表）
========================= */
document.getElementById("file2").addEventListener("change", readFile2);

function readFile2(e) {
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function (evt) {
        let wb = XLSX.read(evt.target.result, { type: "binary" });
        let sheet = wb.Sheets[wb.SheetNames[0]];
        let data = XLSX.utils.sheet_to_json(sheet);

        ruleMap = buildRuleMap(data);

        console.log("规则加载完成：", ruleMap);
    };

    reader.readAsBinaryString(file);
}

/* =========================
   构建规则Map（核心修复）
   支持任意列名
========================= */
function buildRuleMap(data) {

    let map = {};

    data.forEach(row => {

        let keys = Object.keys(row);

        if (keys.length < 2) return;

        let name = row[keys[0]];
        let code = row[keys[1]];

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

    if (!Object.keys(ruleMap).length) {
        alert("请先上传课程目录-2（规则表）");
        return;
    }

    result = [];

    data1.forEach(row => {

        let keys = Object.keys(row);

        // 自动识别字段（容错）
        let course = row[keys[0]]; // 课程名称
        let l1 = row[keys[1]];     // 一级目录
        let l2 = row[keys[2]];     // 二级目录

        let key1 = (l1 || "").trim();
        let key2 = ((l1 || "") + "->" + (l2 || "")).trim();

        // 规则匹配（核心）
        let code =
            ruleMap[key2] ||
            ruleMap[key1] ||
            "未匹配";

        result.push({
            "目录名称(必填)": course || "",
            "目录编码(必填)": code,
            "上级目录编码": l2 || ""
        });
    });

    alert("生成完成！");
    download();
}

/* =========================
   下载Excel
========================= */
function download() {

    let ws = XLSX.utils.json_to_sheet(result);
    let wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "课程目录-2");

    XLSX.writeFile(wb, "课程目录-2-输出.xlsx");
}
