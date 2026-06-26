let resultData = [];
let codeMap = {};
let counterMap = {};

function generateCode(parent) {
    if (!counterMap[parent]) counterMap[parent] = 1;
    else counterMap[parent]++;

    return parent + String(counterMap[parent]).padStart(2, '0');
}

// 读取Excel
document.getElementById("file").addEventListener("change", handleFile);

document.getElementById("drop").addEventListener("click", () => {
    document.getElementById("file").click();
});

function handleFile(e) {
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function (e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: "array" });

        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        let json = XLSX.utils.sheet_to_json(sheet);

        processData(json);
    };

    reader.readAsArrayBuffer(file);
}

function processData(data) {

    resultData = [];

    data.forEach((row, index) => {

        let level1 = row["一级分类"];
        let level2 = row["二级分类"];
        let course = row["课程名称"];

        if (!codeMap[level1]) {
            codeMap[level1] = "TY" + String(Object.keys(codeMap).length + 1).padStart(2, '0');
        }

        let l1 = codeMap[level1];

        let key2 = level1 + "-" + level2;
        if (!codeMap[key2]) {
            codeMap[key2] = generateCode(l1);
        }

        let l2 = codeMap[key2];

        let courseCode = generateCode(l2);

        resultData.push({
            "课程名称": course,
            "上级目录编码": l2,
            "课程目录编码": courseCode
        });
    });

    alert("生成完成！点击下载");
}

// 下载Excel
function downloadResult() {
    let ws = XLSX.utils.json_to_sheet(resultData);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "课程目录-2");

    XLSX.writeFile(wb, "课程目录-2-输出.xlsx");
}
