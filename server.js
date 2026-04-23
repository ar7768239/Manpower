<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>BHEL | Corporate Portal</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>
<style>
/* Reset and Layout */
*{box-sizing:border-box;font-family:Arial}
body{margin:0;background:#f1f5f9; display: flex; flex-direction: column; min-height: 100vh;}

/* Header & Footer */
header{background:#0f172a;color:white;display:flex;justify-content:space-between;align-items:center;padding:14px 40px; flex-shrink: 0; z-index: 100;}
footer{background:#0f172a;color:white;text-align:center;padding:12px;font-size:13px; flex-shrink: 0; z-index: 100;}

.logo{display:flex;align-items:center;gap:10px}
.logo img{height:45px}
nav{display:flex; align-items:center;}
nav a{margin-left:22px;color:white;text-decoration:none;cursor:pointer}
nav a:hover{text-decoration:underline}

/* Hero Style Sections */
.hero-section {
    flex-grow: 1;
    display: flex;
    flex-direction: column;
    background-size: cover;
    background-position: center;
    background-repeat: no-repeat;
    position: relative;
}

#home { background-image: url('https://img.jagranjosh.com/imported/images/E/GK/BHEL-first-lignite-based-power-plant.webp'); }
#dashboard { background-image: url('https://bsmedia.business-standard.com/_media/bs/img/article/2015-03/04/full/1425485435-4013.jpg?im=FeatureCrop,size=(826,465)'); }
#adminDashboard { background-image: url('https://img.jagranjosh.com/images/2024/March/2732024/bhel-2024-compressed.webp'); }

.hero-overlay {
    flex-grow: 1;
    background: rgba(0,0,0,.6);
    display: flex; 
    flex-direction: column;
    align-items: center; 
    justify-content: center; 
    color: white; 
    text-align: center;
    padding: 20px;
}

/* Page Management */
.page{display:none; padding:40px; flex-grow: 1;}
.page.active{display:block}
.hero-page.active { display: flex; padding: 0; } 

/* Components */
.profile-icon {width: 35px; height: 35px; background: #1e90ff; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 18px; margin-left: 20px; text-transform: uppercase; border: 2px solid #fff; cursor: pointer; overflow: hidden;}
.profile-icon img { width: 100%; height: 100%; object-fit: cover; }
.upload-box{background:white; padding:40px; width:600px; margin:40px auto; text-align:center; border-radius:8px; color: #333; box-shadow: 0 4px 6px rgba(0,0,0,0.1);}
table{width:100%;border-collapse:collapse;margin-top:10px; background: white;}
th,td{border:1px solid #ccc;padding:10px;text-align:center; min-width: 100px; color: #333;}
th{background:#1e90ff;color:white}
input[type=number]{width:80px; padding: 5px;}
.btn{padding:12px 18px; background:#1e90ff; border:none; color:white; cursor:pointer; width:100%; margin-bottom:10px; border-radius:5px; font-weight: bold; transition: 0.3s;}
.btn:hover{background: #1873cc;}
.btn.secondary{background:#64748b}

.action-row { display: flex; gap: 10px; margin-bottom: 20px; }
.btn-back { width: auto; background: #475569; padding: 8px 15px; font-size: 14px; }

/* Modals */
.modal{display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,.5); justify-content:center; align-items:center; z-index: 1000;}
.modal.active{display:flex}
.preview-modal-box {background: white; width: 95%; height: 90%; border-radius: 8px; display: flex; flex-direction: column; padding: 15px; overflow: hidden;}
.preview-header {display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; background: #f8fafc; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0; color: #333;}
#selectorArea { display: flex; align-items: center; gap: 10px; }
#sheetSelector { padding: 8px; border-radius: 4px; border: 1px solid #cbd5e1; min-width: 200px; }
#previewContainer { width: 100%; flex-grow: 1; border: 1px solid #ccc; overflow: auto; background: #fff; padding: 10px; }
.modal-box{background:white;padding:25px;width:350px;border-radius:8px;display:flex;flex-direction:column;align-items:center; color: #333;}
.modal-box img.modal-logo { height: 60px; margin-bottom: 15px; }
.modal-box input, .modal-box select{width:100%;padding:8px;margin-bottom:10px}

/* Profile Modal Specifics */
#profileImagePreview { width: 80px; height: 80px; border-radius: 50%; object-fit: cover; margin-bottom: 10px; border: 2px solid #1e90ff; }

/* Drive/Project Section */
.drive-box{background:white;border-radius:8px;padding:20px;min-height:200px; color: #333;}
#fileList li{ padding:10px;border-bottom:1px solid #eee;display: flex;justify-content: space-between;align-items: center; }
.file-item-left { display: flex; align-items: center; gap: 12px; }
#fileList button{margin-left:10px;padding:5px 12px;cursor:pointer;border: 1px solid #ddd;border-radius: 4px;background: #f8fafc;}
.newMenu{display:none;position:absolute;background:white;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 10px rgba(0,0,0,0.15); z-index: 500;}
.newMenu button{display:block;width:180px;padding:10px;border:none;background:white;text-align:left;cursor:pointer}
.section-headline { border-bottom: 2px solid #1e90ff; padding-bottom: 10px; margin-bottom: 20px; color: #0f172a; }
.preview-cb { transform: scale(1.2); cursor: pointer; }
.summary-card { background: #e2e8f0; padding: 15px; border-radius: 8px; margin-bottom: 20px; border-left: 5px solid #1e90ff; color: #333;}
.footer-actions { display: flex; gap: 10px; margin-top: 10px; }
</style>
</head>
<body>

<header>
<div class="logo"><img src="https://upload.wikimedia.org/wikipedia/commons/b/b8/BHEL_logo.svg"></div>
<nav id="authNav"><a onclick="openSignIn()">Sign In</a><a onclick="openLogin()">Login</a></nav>
<nav id="dashNav" style="display:none"></nav>
</header>

<main style="flex: 1; display: flex; flex-direction: column;">
    <div id="home" class="page active hero-page hero-section">
        <div class="hero-overlay">
            <h1>Bharat Heavy Electricals Limited</h1>
            <p>Corporate Manpower Allocation Portal</p>
        </div>
    </div>

    <div id="dashboard" class="page hero-page hero-section">
        <div class="hero-overlay">
            <h2>Employee Dashboard</h2>
            <p>Welcome to the BHEL Corporate Portal. Manage your resources efficiently.</p>
        </div>
    </div>

    <div id="adminDashboard" class="page hero-page hero-section">
        <div class="hero-overlay">
            <h2 style="color: #60a5fa;">Admin Control Center</h2>
            <p>Monitor system logs and manage global project repositories.</p>
        </div>
    </div>

    <div id="uploadPage" class="page">
        <button class="btn btn-back" onclick="showPage('dashboard')">← Back to Dashboard</button>
        <div class="upload-box">
            <h2>Upload Manpower Excel</h2>
            <p style="color: #666; font-size: 14px; margin-bottom: 20px;">Please upload the standard Manpower distribution sheet.</p>
            <input type="file" id="excelFile" accept=".xlsx" style="margin-bottom: 20px;">
            <button class="btn" onclick="uploadExcel()">Process File</button>
        </div>
    </div>

    <div id="previewPage" class="page">
        <div class="action-row">
            <button class="btn btn-back" onclick="showPage('uploadPage')">← Back to Upload</button>
        </div>
        <h2 class="section-headline">Allocation Preview</h2>
        <div style="overflow-x:auto"><table id="dataTable"></table></div>
        <br>
        <div style="display: flex; gap: 15px;">
            <button class="btn" onclick="generateExcel()">Generate Updated Excel</button>
            <button class="btn secondary" onclick="openProjectPage('user', true)">Next: Select WBS Files</button>
        </div>
    </div>

    <div id="userStatus" class="page">
        <button class="btn btn-back" onclick="showPage('adminDashboard')">← Back to Dashboard</button>
        <h2 class="section-headline">User Login History</h2>
        <table id="loginTable"></table>
    </div>

    <div id="projectData" class="page">
        <button class="btn btn-back" id="projectBackBtn" onclick="showPage('dashboard')">← Back to Dashboard</button>
        <h2 class="section-headline">Corporate Project Repository</h2>
        
        <div id="adminUploadControls" style="display:none;gap:20px;align-items:center;margin-bottom:20px">
            <button onclick="toggleNewMenu()" style="background:#1e90ff;color:white;border:none;padding:10px 16px;border-radius:6px;cursor:pointer">+ New</button>
            <div id="newMenu" class="newMenu">
                <button onclick="uploadFile()">Upload File</button>
                <button onclick="uploadFolder()">Upload Folder</button>
            </div>
        </div>

        <div class="drive-box">
            <p id="emptyMsg">No project files available at this time.</p>
            <ul id="fileList" style="list-style:none;padding:0;margin:0"></ul>
        </div>
        <br>
        <button id="submitSelectionBtn" class="btn" style="display:none" onclick="submitProjectSelection()">Confirm File Selection</button>
        <input type="file" id="fileInput" style="display:none" onchange="handleUpload(this.files)">
        <input type="file" id="folderInput" webkitdirectory directory multiple style="display:none" onchange="handleUpload(this.files)">
    </div>
</main>

<div id="signInModal" class="modal">
    <div class="modal-box">
        <img src="https://upload.wikimedia.org/wikipedia/commons/b/b8/BHEL_logo.svg" class="modal-logo">
        <h3>Employee Sign In</h3>
        <input id="siUser" placeholder="Employee Username">
        <input id="siId" placeholder="Employee ID">
        <button class="btn" onclick="createAccount()">Create</button>
        <button class="btn secondary" onclick="closeModals()">Back</button>
    </div>
</div>

<div id="loginModal" class="modal">
    <div class="modal-box">
        <img src="https://upload.wikimedia.org/wikipedia/commons/b/b8/BHEL_logo.svg" class="modal-logo">
        <h3>Login</h3>
        <input id="liUser" placeholder="Username">
        <input id="liId" placeholder="Employee ID / Admin Password">
        <button class="btn" onclick="login()">Login</button>
        <button class="btn secondary" onclick="closeModals()">Back</button>
    </div>
</div>

<div id="profileModal" class="modal">
    <div class="modal-box">
        <h3>User Profile</h3>
        <img id="profileImagePreview" src="" style="display:none">
        <p style="font-size: 12px; color: #666;">Update Profile Picture</p>
        <input type="file" id="profilePicInput" accept="image/*" onchange="previewProfilePic(this)">
        
        <hr style="width: 100%; margin: 15px 0;">
        
        <p style="font-size: 12px; color: #666; text-align: left; width: 100%;">Change Password (Employee ID)</p>
        <input type="password" id="newPass" placeholder="Enter New Password">
        <button class="btn" onclick="updateProfile()">Update Profile</button>
        <button class="btn secondary" onclick="closeModals()">Close</button>
    </div>
</div>

<div id="sheetSelectionModal" class="modal">
    <div class="modal-box">
        <h3>Select Excel Sheets</h3>
        <div id="sheetSelectionList" style="width:100%; max-height: 300px; overflow-y: auto; text-align: left;"></div>
        <button class="btn" onclick="finalizeProjectSubmission()">Submit Final Data</button>
        <button class="btn secondary" onclick="closeModals()">Cancel</button>
    </div>
</div>

<div id="previewModal" class="modal">
    <div class="preview-modal-box">
        <div class="preview-header">
            <div>
                <b id="previewTitle">File Preview</b>
                <div id="selectorArea" style="display:none">
                    <label for="sheetSelector">Select Sheet: </label>
                    <select id="sheetSelector"></select>
                </div>
            </div>
            <div id="previewActionArea" class="footer-actions">
                <button class="btn secondary" id="previewCloseBtn" style="width:auto; margin:0" onclick="closeModals()">Close Preview</button>
            </div>
        </div>
        <div id="previewContainer"></div>
    </div>
</div>

<footer>© 2026 Bharat Heavy Electricals Limited</footer>

<script>
const adminUser="admin", adminPass="admin123";
let uploadedFiles = []; 
let selectionMode = false; 
let finalSelectedFiles = [];
let allocatedManpowerData = []; 
let finalDistributedData = []; 
let currentRole = 'user';
let capturedWBS = ""; 

function saveUserLogin(u,s){
    let logs=JSON.parse(localStorage.getItem("loginLogs"))||[];
    logs.push({user:u, time:new Date().toLocaleString(), status:s});
    localStorage.setItem("loginLogs",JSON.stringify(logs));
}

function openSignIn(){ closeModals(); signInModal.classList.add("active"); }
function openLogin(){ closeModals(); loginModal.classList.add("active"); }

function openProfileModal() {
    closeModals();
    const preview = document.getElementById("profileImagePreview");
    if(localStorage.profilePic) {
        preview.src = localStorage.profilePic;
        preview.style.display = "block";
    } else {
        preview.style.display = "none";
    }
    profileModal.classList.add("active");
}

function previewProfilePic(input) {
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const preview = document.getElementById("profileImagePreview");
            preview.src = e.target.result;
            preview.style.display = "block";
        };
        reader.readAsDataURL(input.files[0]);
    }
}

function updateProfile() {
    const newPassword = document.getElementById("newPass").value;
    const preview = document.getElementById("profileImagePreview");

    if(newPassword) {
        localStorage.empId = newPassword;
        alert("Password updated successfully!");
    }
    
    if(preview.src && preview.src.startsWith("data:image")) {
        localStorage.profilePic = preview.src;
    }

    closeModals();
    updateNavUI(); // Refresh the icon in nav
}

function updateNavUI() {
    const user = localStorage.empUser || "User";
    const profileIcon = document.querySelector(".profile-icon");
    if(profileIcon) {
        if(localStorage.profilePic) {
            profileIcon.innerHTML = `<img src="${localStorage.profilePic}">`;
        } else {
            profileIcon.innerHTML = user.charAt(0).toUpperCase();
        }
    }
}

function closeModals(){
    document.querySelectorAll('.modal').forEach(m=>m.classList.remove("active"));
    document.getElementById("sheetSelector").innerHTML = "";
    document.getElementById("selectorArea").style.display = "none";
    document.getElementById("previewContainer").innerHTML = "";
    document.getElementById("previewCloseBtn").innerText = "Close Preview";
    document.getElementById("previewCloseBtn").onclick = closeModals;
    const dBtn = document.getElementById("dynamicDownloadBtn");
    if(dBtn) dBtn.remove();
}

function clearAuthInputs(){
    document.getElementById("siUser").value = "";
    document.getElementById("siId").value = "";
    document.getElementById("liUser").value = "";
    document.getElementById("liId").value = "";
}

function createAccount(){
    localStorage.empUser=siUser.value; localStorage.empId=siId.value;
    localStorage.removeItem("profilePic"); // Reset profile pic for new account
    alert("Account Created"); clearAuthInputs(); closeModals();
}

function login(){
    let u=liUser.value, p=liId.value;
    if(u===adminUser && p===adminPass){
        currentRole = 'admin';
        authNav.style.display="none";
        dashNav.innerHTML=`<a onclick="showPage('adminDashboard')">Dashboard</a><a onclick="openUserStatus()">User Status</a><a onclick="openProjectPage('admin')">Project Data</a><a onclick="logout()">Logout</a><div class="profile-icon" title="ADMIN" onclick="openProfileModal()">${u.charAt(0)}</div>`;
        dashNav.style.display="flex"; clearAuthInputs(); closeModals(); showPage("adminDashboard"); return;
    }
    if(u===localStorage.empUser && p===localStorage.empId){
        currentRole = 'user';
        saveUserLogin(u,"Success"); authNav.style.display="none";
        
        let iconContent = localStorage.profilePic ? `<img src="${localStorage.profilePic}">` : u.charAt(0).toUpperCase();
        
        dashNav.innerHTML=`<a onclick="showPage('dashboard')">Dashboard</a><a onclick="openManpower()">Manpower Allocation</a><a onclick="openProjectPage('user')">Project Data</a><a onclick="logout()">Logout</a><div class="profile-icon" title="${u}" onclick="openProfileModal()">${iconContent}</div>`;
        dashNav.style.display="flex"; clearAuthInputs(); closeModals(); showPage("dashboard");
    } else { saveUserLogin(u,"Failed"); alert("Invalid Credentials"); }
}

function openUserStatus(){
    showPage("userStatus");
    let logs=JSON.parse(localStorage.getItem("loginLogs"))||[];
    let h=`<tr><th>Username</th><th>Date & Time</th><th>Status</th></tr>`;
    logs.forEach(l=>{ h+=`<tr><td>${l.user}</td><td>${l.time}</td><td>${l.status}</td></tr>`});
    loginTable.innerHTML=h;
}

function logout(){ dashNav.style.display="none"; authNav.style.display="block"; showPage("home"); }

function showPage(id){ 
    document.querySelectorAll(".page").forEach(p=>p.classList.remove("active")); 
    document.getElementById(id).classList.add("active"); 
}

function openManpower(){ showPage("uploadPage"); }

function openProjectPage(role, isSelection = false){
    selectionMode = isSelection;
    const controls = document.getElementById("adminUploadControls");
    const backBtn = document.getElementById("projectBackBtn");
    
    if(isSelection) {
        backBtn.onclick = () => showPage('previewPage');
        backBtn.innerText = "← Back to Allocation";
    } else {
        backBtn.onclick = () => showPage(role === 'admin' ? 'adminDashboard' : 'dashboard');
        backBtn.innerText = "← Back to Dashboard";
    }

    controls.style.display = (role === 'admin') ? "flex" : "none";
    document.getElementById("submitSelectionBtn").style.display = selectionMode ? "block" : "none";
    renderGlobalFileList(role);
    showPage('projectData');
}

/* MANPOWER ALLOCATION LOGIC */
let excelData=[];
function uploadExcel(){
    if(excelFile.files.length===0){alert("Select file"); return;}
    const reader=new FileReader();
    reader.onload=e=>{
        const data = new Uint8Array(e.target.result);
        const wb=XLSX.read(data,{type:"array"});
        excelData=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        excelData.forEach(r=>{ r.updated=r.Manpower || 0; r.selected=false; });
        renderTable(); showPage("previewPage");
    };
    reader.readAsArrayBuffer(excelFile.files[0]);
}

function renderTable(){
    let h=`<tr><th>Select</th><th>Area</th><th>Current Manpower</th><th>Allocation %</th><th>Updated Manpower</th></tr>`;
    excelData.forEach((r,i)=>{ h+=`<tr><td><input type="checkbox" onchange="excelData[${i}].selected=this.checked"></td><td>${r.Area || 'N/A'}</td><td>${r.Manpower || 0}</td><td><input type="number" value="100" onchange="calc(${i},this.value)"></td><td id="u${i}">${r.updated}</td></tr>`;});
    dataTable.innerHTML=h;
}
function calc(i,v){ excelData[i].updated=Math.round((excelData[i].Manpower || 0)*v/100); document.getElementById("u"+i).innerText=excelData[i].updated; }

function generateExcel(){
    allocatedManpowerData = excelData.filter(r=>r.selected).map(r=>({Area:r.Area, Manpower:r.updated}));
    if(allocatedManpowerData.length === 0) { alert("No rows selected in Manpower Allocation"); return; }
    
    let horizontalRow = { "SL no.": 1, "WBS ": (capturedWBS || "N/A").toUpperCase() }; 
    allocatedManpowerData.forEach(item => {
        horizontalRow[item.Area] = item.Manpower;
    });
    horizontalRow["TOTAL"] = allocatedManpowerData.reduce((sum, item) => sum + item.Manpower, 0);

    const ws=XLSX.utils.json_to_sheet([horizontalRow]);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Updated"); 
    XLSX.writeFile(wb,"BHEL_Updated_Manpower.xlsx");
}

/* SHARED PROJECT DATA LOGIC */
function toggleNewMenu(){ let m=document.getElementById("newMenu"); m.style.display=m.style.display==="block"?"none":"block"; }
function uploadFile(){ document.getElementById("fileInput").click(); }
function uploadFolder(){ document.getElementById("folderInput").click(); }

function handleUpload(files){
    for(let i=0; i<files.length; i++){ let f = files[i]; f.tempSelected = false; uploadedFiles.push(f); }
    renderGlobalFileList('admin'); document.getElementById("newMenu").style.display="none";
}

function renderGlobalFileList(role){
    let list=document.getElementById("fileList"), empty=document.getElementById("emptyMsg");
    list.innerHTML = ""; if(uploadedFiles.length > 0) empty.style.display="none"; else empty.style.display="block";
    uploadedFiles.forEach((file, index) => {
        let li=document.createElement("li");
        let leftSide = document.createElement("div"); leftSide.className = "file-item-left";
        if(selectionMode && role === 'user') {
            let cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = file.tempSelected;
            cb.onchange = function() { file.tempSelected = this.checked; }; leftSide.appendChild(cb);
        }
        let nameSpan = document.createElement("span"); nameSpan.innerText = `📄 ${file.name}`;
        leftSide.appendChild(nameSpan); li.appendChild(leftSide);
        let div = document.createElement("div"); let viewBtn=document.createElement("button"); viewBtn.innerText="View";
        viewBtn.onclick=function(){
            document.getElementById("previewTitle").innerText = file.name;
            const reader = new FileReader();
            if(file.name.endsWith('.xlsx') || file.name.endsWith('.xls')){
                reader.onload = function(e){
                    const data = new Uint8Array(e.target.result); const workbook = XLSX.read(data, {type: 'array'});
                    const selector = document.getElementById("sheetSelector"); const selectorArea = document.getElementById("selectorArea");
                    selector.innerHTML = ""; selectorArea.style.display = "block"; 
                    workbook.SheetNames.forEach((sheetName) => { const opt = document.createElement("option"); opt.value = sheetName; opt.innerText = sheetName; selector.appendChild(opt); });
                    selector.onchange = function() { document.getElementById("previewContainer").innerHTML = XLSX.utils.sheet_to_html(workbook.Sheets[this.value]); };
                    document.getElementById("previewContainer").innerHTML = XLSX.utils.sheet_to_html(workbook.Sheets[workbook.SheetNames[0]]);
                    document.getElementById("previewModal").classList.add("active");
                };
                reader.readAsArrayBuffer(file);
            } else {
                reader.onload = function(e){
                    document.getElementById("selectorArea").style.display = "none";
                    document.getElementById("previewContainer").innerHTML = `<iframe src="${e.target.result}" style="width:100%; height:100%; border:none;"></iframe>`;
                    document.getElementById("previewModal").classList.add("active");
                };
                reader.readAsDataURL(file);
            }
        };
        div.append(viewBtn);
        if(role === 'admin'){
            let delBtn=document.createElement("button"); delBtn.innerText="Delete";
            delBtn.onclick=function(){ uploadedFiles.splice(index, 1); renderGlobalFileList('admin'); }; div.append(delBtn);
        }
        li.appendChild(div); list.appendChild(li);
    });
}

async function submitProjectSelection() {
    finalSelectedFiles = uploadedFiles.filter(f => f.tempSelected);
    if(finalSelectedFiles.length === 0) { alert("Please select at least one project file."); return; }
    const container = document.getElementById("sheetSelectionList");
    container.innerHTML = ""; let excelFilesFound = false;
    for (let file of finalSelectedFiles) {
        if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            excelFilesFound = true; const data = await file.arrayBuffer(); const workbook = XLSX.read(data, { type: 'array' });
            const div = document.createElement("div"); div.style.marginBottom = "15px";
            div.innerHTML = `<label style="font-size:13px; font-weight:bold">${file.name}</label><br>`;
            const select = document.createElement("select"); select.style.marginTop = "5px";
            workbook.SheetNames.forEach(name => { const opt = document.createElement("option"); opt.value = name; opt.innerText = name; select.appendChild(opt); });
            div.appendChild(select); container.appendChild(div);
            file.selectedSheet = workbook.SheetNames[0]; select.onchange = (e) => { file.selectedSheet = e.target.value; };
        }
    }
    if (excelFilesFound) document.getElementById("sheetSelectionModal").classList.add("active");
    else finalizeProjectSubmission();
}

function downloadFinalReport() {
    if(finalDistributedData.length === 0) return;
    
    let finalRow = { "SL no.": 1, "WBS ": (capturedWBS || "N/A").toUpperCase() }; 
    finalDistributedData.forEach(item => {
        finalRow[item.Area] = item["Distributed Panels"];
    });
    finalRow["TOTAL"] = finalDistributedData.reduce((sum, item) => sum + item["Distributed Panels"], 0);

    const ws = XLSX.utils.json_to_sheet([finalRow]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Final Distribution");
    XLSX.writeFile(wb, "BHEL_Manpower_Distribution.xlsx");
}

async function finalizeProjectSubmission() {
    let combinedHtml = "";
    if(allocatedManpowerData.length === 0) {
        allocatedManpowerData = excelData.filter(r=>r.selected).map(r=>({Area:r.Area, Manpower:r.updated}));
    }
    combinedHtml += `<div class="summary-card"><h3>Step: Select WBS Rows</h3><p>Select specific rows to capture **WBS ** and distribute panels.</p></div>`;
    for (let file of finalSelectedFiles) {
        if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            const data = await file.arrayBuffer(); const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = file.selectedSheet || workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            let htmlTable = XLSX.utils.sheet_to_html(sheet);
            combinedHtml += `<h3 style="background:#f1f5f9; padding:12px; border-left:5px solid #1e90ff; margin-top:20px;">${file.name}</h3>`;
            let styledTable = htmlTable.replace('<table>', '<table class="selection-table" style="border-collapse:collapse; width:100%; margin-bottom:20px;">')
                                       .replaceAll('<td>', '<td style="border:1px solid #ccc; padding:8px; text-align:center;">')
                                       .replaceAll('<th>', '<th style="border:1px solid #ccc; padding:10px; background:#1e90ff; color:white;">');
            styledTable = styledTable.replace(/<tr>/g, '<tr><td style="border:1px solid #ccc; width:40px; text-align:center;"><input type="checkbox" class="preview-cb"></td>');
            styledTable = styledTable.replace(/<tr>/, '<tr><th style="background:#1e90ff; color:white; width:40px;">Select</th>');
            combinedHtml += `<div style="overflow-x:auto;">${styledTable}</div>`;
        }
    }
    document.getElementById("previewTitle").innerText = "WBS Row Selection";
    document.getElementById("previewContainer").innerHTML = combinedHtml;
    const closeBtn = document.getElementById("previewCloseBtn");
    closeBtn.innerText = "Calculate & Finish";
    
    closeBtn.onclick = function() {
        const tables = document.querySelectorAll('.selection-table');
        let selectedRowsHtml = "";
        let totalPanelsFromWBS = 0;
        capturedWBS = ""; 

        tables.forEach(table => {
            const rows = table.querySelectorAll('tr');
            if (rows.length < 2) return;
            const headerCells = rows[0].querySelectorAll('th');
            let panelColIdx = -1;
            let wbsColIdx = -1;
            
            headerCells.forEach((cell, idx) => {
                const text = cell.innerText.trim().toLowerCase();
                if((text.includes('panel') || text.includes('qty')) && !text.includes('id') && !text.includes('code')) panelColIdx = idx;
                if(text.includes('wbs')) wbsColIdx = idx;
            });

            for(let i=1; i < rows.length; i++) {
                const row = rows[i];
                const checkbox = row.querySelector('.preview-cb');
                if(checkbox && checkbox.checked) {
                    selectedRowsHtml += row.outerHTML;
                    const cells = row.querySelectorAll('td');
                    if(!capturedWBS) {
                        if(wbsColIdx !== -1 && cells[wbsColIdx]) {
                            capturedWBS = cells[wbsColIdx].innerText.trim().toUpperCase();
                        } else {
                            for(let j=1; j<cells.length; j++){
                                let cellText = cells[j].innerText.trim();
                                if(cellText.includes('/') && cellText.includes('-')){
                                    capturedWBS = cellText.toUpperCase();
                                    break;
                                }
                            }
                        }
                    }
                    let panelValue = 0;
                    if(panelColIdx !== -1 && cells[panelColIdx]) {
                        panelValue = parseFloat(cells[panelColIdx].innerText.replace(/[^0-9.]/g, '')) || 0;
                    } else {
                        for(let j=cells.length-1; j>=1; j--) {
                           let textVal = cells[j].innerText.trim();
                           if(!textVal.includes('/') && !textVal.includes('-')) {
                               let val = parseFloat(textVal.replace(/[^0-9.]/g, ''));
                               if(!isNaN(val) && val > 0 && val < 500) { panelValue = val; break; }
                           }
                        }
                    }
                    totalPanelsFromWBS += panelValue;
                }
            }
        });

        if(totalPanelsFromWBS === 0) { alert("Error: No panel count found in the selected rows."); return; }
        const totalManpower = allocatedManpowerData.reduce((sum, item) => sum + item.Manpower, 0);
        finalDistributedData = []; 
        let finalReport = `<h2>Final Submission Report</h2>
        <div class="summary-card">
            <p><b>Respective WBS :</b> ${capturedWBS || "N/A"}</p>
            <p><b>Total Panels (Selected Rows):</b> ${totalPanelsFromWBS}</p>
            <p><b>Total Manpower:</b> ${totalManpower}</p>
        </div>
        <table style="width:100%; border-collapse:collapse;">
            <tr><th>Area</th><th>Manpower</th><th>Distributed Panels</th></tr>`;
        
        allocatedManpowerData.forEach(area => {
            let dist = totalManpower > 0 ? Math.round((area.Manpower / totalManpower) * totalPanelsFromWBS) : 0;
            finalDistributedData.push({ "Area": area.Area, "Manpower": area.Manpower, "Distributed Panels": dist });
            finalReport += `<tr><td>${area.Area}</td><td>${area.Manpower}</td><td><b style="color:#1e90ff">${dist}</b></td></tr>`;
        });
        finalReport += `</table><br><h3>Data Preview of Selected Rows</h3><table style="width:100%; border-collapse:collapse;">${selectedRowsHtml}</table>`;
        document.getElementById("previewContainer").innerHTML = finalReport;
        
        if(!document.getElementById("dynamicDownloadBtn")){
            const dBtn = document.createElement("button");
            dBtn.id = "dynamicDownloadBtn";
            dBtn.className = "btn"; dBtn.style.width = "auto"; dBtn.style.background = "#10b981";
            dBtn.innerText = "⬇ Download Final Distribution";
            dBtn.onclick = downloadFinalReport;
            document.getElementById("previewActionArea").prepend(dBtn);
        }
        closeBtn.innerText = "Confirm & Finish";
        closeBtn.onclick = function() { alert("Data finalized successfully."); closeModals(); showPage('dashboard'); };
    };
    document.getElementById("sheetSelectionModal").classList.remove("active");
    document.getElementById("previewModal").classList.add("active");
}
</script>
</body>
</html>
