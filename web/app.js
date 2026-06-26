let data1 = [];
let ruleMap = {};
let data2 = [];

/* ========== 文件1 ========== */
document.getElementById("file1").addEventListener("change", function(e){
    readExcel(e, 1);
});

/* ========== 文件2 ========== */
document.getElementById("file2").addEventListener("change", function(e){
    readExcel(e, 2);
});

function readExcel(e, type){
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function(evt){
        let wb = XLSX.read(evt.target.result, {type:"binary"});
        let sheet = wb.Sheets[wb.SheetNames[0]];
        let data = XLSX.utils.sheet_to_json(sheet);

        if(type === 1){
            data1 = data;
            console.log("file1加载成功", data1);
        }

        if(type === 2){
            data2 = data;
            ruleMap = buildRuleMap(data2);
            console.log("file2规则加载成功", ruleMap);
        }
    };

    reader.readAsBinaryString(file);
}

/* ========== 规则表 ========== */
function buildRuleMap(data){
    let map = {};

    data.forEach(row => {
        let keys = Object.keys(row);

        let name = row[keys[0]];
        let code = row[keys[1]];

        if(name && code){
            map[String(name).trim()] = String(code).trim();
        }
    });

    return map;
}

/* ========== 生成 ========== */
function generate(){

    console.log("点击生成");

    if(!data1.length){
        alert("请先上传课程目录-1");
        return;
    }

    if(!data2.length){
        alert("❌ 课程目录-2 没有读取成功（重点检查上传）");
        return;
    }

    let result = [];

    data1.forEach(row => {

        let keys = Object.keys(row);

        let course = row[keys[0]];
        let l1 = row[keys[1]];
        let l2 = row[keys[2]];

        let code =
            ruleMap[l1] ||
            ruleMap[l1 + "->" + l2] ||
            "未匹配";

        result.push({
            "目录名称(必填)": course,
            "目录编码(必填)": code,
            "上级目录编码": l2
        });
    });

    console.log("结果：", result);

    download(result);
}

/* ========== 下载 ========== */
function download(result){
    let ws = XLSX.utils.json_to_sheet(result);
    let wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "结果");

    XLSX.writeFile(wb, "课程目录输出.xlsx");
}
