let data1 = [];
let ruleMap = {};
let result = [];

document.getElementById("file1").addEventListener("change", readFile1);
document.getElementById("file2").addEventListener("change", readFile2);

function readFile1(e){
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function(evt){
        let wb = XLSX.read(evt.target.result, {type:"binary"});
        let sheet = wb.Sheets[wb.SheetNames[0]];
        data1 = XLSX.utils.sheet_to_json(sheet);
    };

    reader.readAsBinaryString(file);
}

function readFile2(e){
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function(evt){
        let wb = XLSX.read(evt.target.result, {type:"binary"});

        // 关键：读取规则sheet（你文件里的“上级目录编码”）
        let sheet = wb.Sheets[wb.SheetNames[0]];
        let data = XLSX.utils.sheet_to_json(sheet);

        ruleMap = {};
        data.forEach(r => {
            ruleMap[r["名称"]] = r["编码"];
        });

    };

    reader.readAsBinaryString(file);
}

// 核心生成逻辑
function generate(){

    result = [];

    data1.forEach(row => {

        let l1 = row["课程目录"];
        let l2 = row["二级目录"];
        let course = row["课程名称"];

        let key1 = l1;
        let key2 = l1 + "->" + l2;

        let code = ruleMap[key2] || ruleMap[key1] || "未匹配";

        result.push({
            "课程名称": course,
            "上级目录编码": code
        });
    });

    download();
}

// 下载Excel
function download(){
    let ws = XLSX.utils.json_to_sheet(result);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "结果");

    XLSX.writeFile(wb, "课程目录-2-输出.xlsx");
}
