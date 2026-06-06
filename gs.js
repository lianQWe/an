<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1.0">
    <title>老李的工时记录表</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 微软雅黑;
        }
        body {
            background: #fff;
            padding: 30px;
        }
        .box{
            max-width: 850px;
            margin: 0 auto;
        }
        h1{
            text-align: center;
            font-size: 22px;
            margin-bottom: 20px;
            font-weight: bold;
            outline: none;
            background:linear-gradient(90deg,#ff0000,#ff9900,#ffff00,#33cc33,#0099ff,#9933ff);
            -webkit-background-clip:text;
            color:transparent;
        }
        .month-box{
            text-align: center;
            margin-bottom: 15px;
            position: relative;
        }
        #monthBtn{
            padding: 8px 16px;
            font-size: 16px;
            background: #3498db;
            color: #fff;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .month-popup{
            display: none;
            position: absolute;
            left: 50%;
            transform: translateX(-50%);
            background: white;
            border: 1px solid #ccc;
            padding: 10px;
            grid-template-columns: repeat(4,1fr);
            gap: 8px;
            z-index: 99;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
        }
        .month-popup button{
            padding: 6px 10px;
            background: #f1f1f1;
            border: 1px solid #ccc;
            border-radius: 3px;
            cursor: pointer;
            color: #000;
        }
        .month-popup button.active{
            background: #3498db;
            color: #fff;
        }
        .btn{
            margin-bottom:15px;
            display:flex;
            gap:8px;
            flex-wrap:wrap;
        }
        button{
            padding:6px 12px;
            border:none;
            border-radius:3px;
            cursor:pointer;
            color:#fff;
        }
        .btn1{background:#666;}
        .btn2{background:#27ae60;}
        .btn3{background:#3498db;}
        .btnDel{background:#e53935;}
        .btnBack{background:#999;}
        .btnSave{background:#9c27b0;}
        .btnImport{background:#ff5722;}
        table{
            width:100%;
            border-collapse: collapse;
        }
        th{
            background:#ed7a7a;
            color:#000;
            border:1px solid #333;
            padding:10px;
            font-size:16px;
        }
        td{
            border:1px solid #333;
            padding:10px;
            text-align:center;
            font-size:15px;
        }
        tbody tr td:nth-child(1){background:#e0f7fa;}
        tbody tr td:nth-child(2){background:#ffc9d3;}
        tbody tr td:nth-child(3){background:#ffb347;}
        tbody tr td:nth-child(4){background:#89e889;}
        .total td{background:#ddd !important;}
        td[contenteditable]{outline:none;}
        input#fileIn{display:none;}
    </style>
</head>
<body>
<div class="box">
    <h1 id="tableTitle" contenteditable="true">老李的工时记录表</h1>
    
    <div class="month-box">
        <button id="monthBtn">当前：1月</button>
        <div class="month-popup" id="monthPopup">
            <button data-m="1">1月</button>
            <button data-m="2">2月</button>
            <button data-m="3">3月</button>
            <button data-m="4">4月</button>
            <button data-m="5">5月</button>
            <button data-m="6">6月</button>
            <button data-m="7">7月</button>
            <button data-m="8">8月</button>
            <button data-m="9">9月</button>
            <button data-m="10">10月</button>
            <button data-m="11">11月</button>
            <button data-m="12">12月</button>
        </div>
    </div>
    <div class="btn">
        <button class="btn1" onclick="addRow()">添加一行</button>
        <button class="btn1" onclick="delRow()">删除一行</button>
        <button class="btnDel" id="clearBtn">清空当月数据</button>
        <button class="btnBack" id="backBtn">恢复当月数据</button>
        <button class="btn2" onclick="exportExcel()">导出表格</button>
        <button class="btn3" onclick="copyShare()">分享表格</button>
        <button class="btnSave" onclick="saveBackup()">导出备份</button>
        <button class="btnImport" onclick="document.getElementById('fileIn').click()">导入备份</button>
        <input type="file" id="fileIn" accept=".json" onchange="importBackup(event)">
    </div>
    <table id="mainTable">
        <thead>
            <tr>
                <th>星期</th>
                <th>日期</th>
                <th>时间段</th>
                <th>工时(小时)</th>
            </tr>
        </thead>
        <tbody id="tb">
            <tr>
                <td contenteditable></td>
                <td contenteditable></td>
                <td contenteditable></td>
                <td contenteditable></td>
            </tr>
            <tr class="total">
                <td>合计</td>
                <td></td>
                <td></td>
                <td id="sum">0</td>
            </tr>
        </tbody>
    </table>
</div>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script>
let currentMonth = 1;
const weekNames = ['日','一','二','三','四','五','六'];
const monthBtn = document.getElementById('monthBtn');
const monthPopup = document.getElementById('monthPopup');
const monthBtns = monthPopup.querySelectorAll('button');

function loadLastMonth() {
    let m = localStorage.getItem('last_selected_month');
    if (m && !isNaN(m)) currentMonth = Number(m);
}

function saveLastMonth() {
    localStorage.setItem('last_selected_month', currentMonth);
}

monthBtn.onclick = () => { monthPopup.style.display = 'grid'; };

document.addEventListener('click', (e) => {
    if (!monthBtn.contains(e.target) && !monthPopup.contains(e.target))
        monthPopup.style.display = 'none';
});

monthBtns.forEach(btn => {
    btn.onclick = () => {
        currentMonth = Number(btn.dataset.m);
        monthBtn.textContent = `当前：${currentMonth}月`;
        monthPopup.style.display = 'none';
        monthBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        saveLastMonth();
        loadMonthData();
    };
});

function getThisYear() { return new Date().getFullYear(); }
function getMonthKey() { return `work_data_${currentMonth}`; }
function getMonthBackupKey() { return `work_backup_${currentMonth}`; }

function getWeek(year, month, day) {
    let d = new Date(year, month - 1, day);
    return '周' + weekNames[d.getDay()];
}

function autoFillDate() {
    const y = getThisYear();
    const m = currentMonth;
    const rows = document.querySelectorAll("#tb tr:not(.total)");
    rows.forEach((tr, i) => {
        let day = i + 1;
        tr.cells[0].innerText = getWeek(y, m, day);
        tr.cells[1].innerText = `${m}月${day}日`;
    });
}

function saveMonthData() {
    const title = document.getElementById("tableTitle").innerText;
    const html = document.getElementById("tb").innerHTML;
    localStorage.setItem('work_title', title);
    localStorage.setItem(getMonthKey(), html);
    localStorage.setItem(getMonthBackupKey(), html);
}

function makeAllCellsEditable() {
    document.querySelectorAll("#tb tr:not(.total) td").forEach(td => {
        td.setAttribute("contenteditable", "true");
    });
}

function loadMonthData() {
    const title = localStorage.getItem('work_title') || '工时记录表';
    const html = localStorage.getItem(getMonthKey());
    document.getElementById("tableTitle").innerText = title;
    if (html) {
        document.getElementById("tb").innerHTML = html;
        makeAllCellsEditable();
    } else {
        resetTable();
        autoFillDate();
    }
    countSum();
    monthBtns.forEach(b => b.classList.toggle('active', Number(b.dataset.m) === currentMonth));
}

function resetTable() {
    document.getElementById("tb").innerHTML = `
    <tr><td contenteditable></td><td contenteditable></td><td contenteditable></td><td contenteditable></td></tr>
    <tr class="total"><td>合计</td><td></td><td></td><td id="sum">0</td></tr>`;
}

function addRow() {
    const y = getThisYear();
    const m = currentMonth;
    const rows = document.querySelectorAll("#tb tr:not(.total)");
    const newDay = rows.length + 1;
    const weekText = getWeek(y, m, newDay);
    const tr = document.createElement("tr");
    tr.innerHTML = `
        <td contenteditable>${weekText}</td>
        <td contenteditable>${m}月${newDay}日</td>
        <td contenteditable></td>
        <td contenteditable></td>
    `;
    document.querySelector(".total").before(tr);
    saveMonthData();
}

function delRow() {
    const list = document.querySelectorAll("#tb tr:not(.total)");
    if (list.length <= 1) return;
    list[list.length - 1].remove();
    saveMonthData();
}

function parseTimeStr(t) {
    t = t.trim();
    if (!t) return 0;
    if (!t.includes(':')) return parseFloat(t) || 0;
    const [h, m] = t.split(':').map(Number);
    return h + (m || 0) / 60;
}

function getWorkTime(str) {
    if (!str) return '';
    str = str.replace(/[－~～]/g, '-');
    if (!str.includes('-')) return '';
    const [s, e] = str.split('-').map(parseTimeStr);
    let total = e - s;
    let rest = 0;
    if (e > 17.5) rest = 1.5;
    else if (e > 11.5) rest = 1.0;
    return total > 0 ? (total - rest).toFixed(1) : '';
}

function countSum() {
    let sum = 0;
    document.querySelectorAll("#tb tr:not(.total)").forEach(tr => {
        const v = parseFloat(tr.cells[3].innerText);
        if (!isNaN(v)) sum += v;
    });
    document.getElementById("sum").innerText = sum.toFixed(1);
}

function autoInsertDash(td) {
    const val = td.innerText.trim();
    if (val.includes('-')) return;
    if (/^\d{1,2}:\d{2}$/.test(val)) {
        td.innerText = val + ' - ';
        const r = document.createRange();
        const s = window.getSelection();
        r.selectNodeContents(td);
        r.collapse(false);
        s.removeAllRanges();
        s.addRange(r);
    }
    const v = getWorkTime(td.innerText);
    if (v) td.nextElementSibling.innerText = v;
}

function exportExcel() {
    const rows = [];
    rows.push(["星期", "日期", "时间段", "工时(小时)"]);
    
    const trs = document.querySelectorAll("#tb tr");
    trs.forEach(tr => {
        const arr = [];
        tr.querySelectorAll("td").forEach(td => arr.push(td.innerText.trim()));
        rows.push(arr);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "工时");
    const title = document.getElementById("tableTitle").innerText.trim() || "工时记录表";
    XLSX.writeFile(wb, `${currentMonth}月_${title}.xlsx`);
}

function copyShare() {
    const txt = `${currentMonth}月 ` + document.getElementById("tableTitle").innerText + " 工时表";
    const i = document.createElement("input");
    i.value = txt;
    document.body.appendChild(i);
    i.select();
    document.execCommand("copy");
    document.body.removeChild(i);
    alert("已复制到剪贴板");
}

function saveBackup() {
    const data = {
        month: currentMonth,
        title: document.getElementById("tableTitle").innerText,
        html: document.getElementById("tb").innerHTML
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${currentMonth}月工时备份.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    alert("当前月份备份已导出");
}

function importBackup(e) {
    const f = e.target.files[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = ev => {
        try {
            const d = JSON.parse(ev.target.result);
            localStorage.setItem(getMonthKey(), d.html || "");
            if (d.title) localStorage.setItem('work_title', d.title);
            loadMonthData();
            alert(`已导入【${currentMonth}月】，可直接编辑`);
        } catch { 
            alert("备份文件格式错误"); 
        }
    };
    r.readAsText(f);
    e.target.value = "";
}

document.getElementById("clearBtn").onclick = () => {
    if (confirm(`确定清空【${currentMonth}月】数据？`)) {
        localStorage.removeItem(getMonthKey());
        loadMonthData();
    }
};

document.getElementById("backBtn").onclick = () => {
    const bak = localStorage.getItem(getMonthBackupKey());
    if (!bak) return alert("无当月备份");
    if (confirm(`恢复【${currentMonth}月】数据？`)) {
        localStorage.setItem(getMonthKey(), bak);
        loadMonthData();
    }
};

document.querySelector("#tb").addEventListener("input", e => {
    const td = e.target;
    if (td.cellIndex === 2) {
        autoInsertDash(td);
        const v = getWorkTime(td.innerText);
        if (v) td.nextElementSibling.innerText = v;
    }
    countSum();
    saveMonthData();
});

document.getElementById("tableTitle").addEventListener("input", saveMonthData);

window.onload = () => {
    loadLastMonth();
    monthBtn.textContent = `当前：${currentMonth}月`;
    loadMonthData();
};
</script>
</body>
</html>
