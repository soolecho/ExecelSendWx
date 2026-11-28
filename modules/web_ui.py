"""网页端单页界面（内嵌 HTML/CSS/JS，不依赖任何外部资源）。

页面结构：登录遮罩 + 单页四 Tab（状态/日志、手动发送、定时任务、监控任务）。
手机宽度下 Tab 固定在底部，电脑宽度下在顶部。
"""

PAGE_HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>表格自动发送By春风予Lu · 局域网控制</title>
<style>
:root{
  --bg:#f4f6fb; --card:#ffffff; --text:#1f2430; --muted:#7a8294;
  --line:#e3e8f2; --accent:#2f6bff; --accent-soft:#eaf0ff;
  --ok:#1e8e3e; --err:#d93025; --warn:#e67e22; --shadow:0 8px 24px rgba(31,36,48,.08);
}
@media (prefers-color-scheme: dark){
  :root{
    --bg:#12151c; --card:#1b1f28; --text:#e8ecf5; --muted:#9aa4b8;
    --line:#2a3140; --accent:#5b8cff; --accent-soft:#1e2a44;
    --shadow:0 8px 24px rgba(0,0,0,.35);
  }
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{
  background:var(--bg);color:var(--text);
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif;
  font-size:14px;line-height:1.55;
  padding-bottom:env(safe-area-inset-bottom);
}
.hidden{display:none!important}
.muted{color:var(--muted)}
.small{font-size:12px}
.err{color:var(--err);min-height:18px}
.ok{color:var(--ok)}
button{
  font:inherit;border:1px solid var(--line);background:var(--card);color:var(--text);
  border-radius:10px;padding:8px 14px;cursor:pointer;transition:.15s;
}
button:hover{border-color:var(--accent)}
button.primary{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600}
button.primary:disabled{opacity:.55;cursor:not-allowed}
button.ghost{background:transparent}
input[type=text],input[type=password],input[type=number],select,textarea{
  font:inherit;color:var(--text);background:var(--card);
  border:1px solid var(--line);border-radius:10px;padding:8px 10px;width:100%;
}
textarea{min-height:70px;resize:vertical}
label{display:block}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.row>*{min-width:0}
.grow{flex:1 1 220px}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;
  padding:14px;margin-bottom:12px;box-shadow:var(--shadow)}
.card h3{margin:0 0 10px;font-size:14px;font-weight:700}
.field{margin-bottom:10px}
.field>span{display:block;margin-bottom:4px;color:var(--muted);font-size:12px}

/* 登录 */
.login-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.login-card{width:100%;max-width:360px;background:var(--card);border:1px solid var(--line);
  border-radius:18px;padding:28px 22px;box-shadow:var(--shadow);text-align:center}
.login-card h1{font-size:20px;margin:0 0 4px}
.login-card input{margin:18px 0 10px;text-align:center;letter-spacing:4px}
.login-card button{width:100%}

/* 顶栏 + Tab */
.topbar{position:sticky;top:0;z-index:20;display:flex;align-items:center;gap:10px;
  padding:10px 14px;background:var(--card);border-bottom:1px solid var(--line)}
.brand{font-weight:700;white-space:nowrap}
#lanUrl{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tabs{display:flex;gap:6px;padding:8px 14px;background:var(--card);
  border-bottom:1px solid var(--line);overflow-x:auto}
.tabs button{white-space:nowrap;border-radius:999px;padding:6px 14px}
.tabs button.active{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600}
main{padding:14px;max-width:920px;margin:0 auto}

/* 状态卡 */
.stat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.stat{background:var(--accent-soft);border-radius:12px;padding:12px}
.stat .k{font-size:12px;color:var(--muted)}
.stat .v{font-size:18px;font-weight:700;margin-top:2px}
.logbox{height:300px;overflow:auto;background:var(--bg);border-radius:10px;
  padding:8px 10px;font-family:Consolas,"Courier New",monospace;font-size:12px;white-space:pre-wrap;word-break:break-all}
.logline{padding:1px 0}
.logline .t{color:var(--muted);margin-right:6px}

/* 列表 */
.list-item{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid var(--line)}
.list-item:last-child{border-bottom:none}
.list-item .main{flex:1;min-width:0}
.list-item .name{font-weight:600}
.list-item .meta{font-size:12px;color:var(--muted)}
.switch{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--muted);white-space:nowrap}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chip{display:inline-flex;align-items:center;gap:4px;border:1px solid var(--line);
  border-radius:999px;padding:4px 10px;font-size:12px;cursor:pointer;background:var(--card)}
.chip input{margin:0}
.names{max-height:180px;overflow:auto;border:1px solid var(--line);border-radius:10px;padding:8px}
.names label{display:flex;gap:6px;align-items:center;padding:2px 0;font-size:13px}
.progress{height:8px;background:var(--line);border-radius:999px;overflow:hidden;margin:8px 0}
.progress>i{display:block;height:100%;width:0;background:var(--accent);transition:width .3s}
.gp{padding:8px 14px 10px;background:var(--card);border-bottom:1px solid var(--line)}
.gp .gp-row{display:flex;justify-content:space-between;gap:10px;font-size:12px;
  color:var(--muted);margin-bottom:2px}
.gp .progress{margin:0}
.gp .progress>i.indet{width:100%;opacity:.5}
.badge{font-size:11px;padding:1px 8px;border-radius:999px;background:var(--accent-soft);color:var(--accent)}
.modal{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:60;
  display:flex;align-items:flex-start;justify-content:center;padding:18px;overflow:auto}
.modal .sheet{background:var(--card);border:1px solid var(--line);border-radius:16px;
  max-width:660px;width:100%;padding:16px;box-shadow:var(--shadow);margin:auto}
.modal .sheet h3{margin:0 0 12px;font-size:15px}

/* 手机：底部 Tab */
@media (max-width:720px){
  .tabs{position:fixed;left:0;right:0;bottom:0;top:auto;z-index:30;
    border-top:1px solid var(--line);border-bottom:none;justify-content:space-around;
    padding:6px 8px calc(6px + env(safe-area-inset-bottom));overflow:visible}
  .tabs button{flex:1;padding:8px 4px;font-size:12px;border-radius:10px}
  main{padding:12px 12px 88px}
  .brand{font-size:13px;overflow:hidden;text-overflow:ellipsis}
  #lanUrl{display:none}
}
</style>
</head>
<body>

<div id="login" class="login-wrap">
  <div class="login-card">
    <h1>📨 表格自动发送By春风予Lu</h1>
    <p class="muted small">局域网网页控制</p>
    <input id="pwd" type="password" placeholder="网页控制密码" autocomplete="current-password">
    <button id="loginBtn" class="primary">登录</button>
    <div id="loginErr" class="err"></div>
  </div>
</div>

<div id="app" class="hidden">
  <header class="topbar">
    <div class="brand">📨 表格自动发送By春风予Lu</div>
    <div id="lanUrl" class="muted small"></div>
    <button id="logoutBtn" class="ghost">退出</button>
  </header>

  <nav class="tabs" id="tabs">
    <button data-tab="status" class="active">状态/日志</button>
    <button data-tab="send">手动发送</button>
    <button data-tab="schedule">定时任务</button>
    <button data-tab="monitor">监控任务</button>
  </nav>

  <div id="gp" class="gp hidden">
    <div class="gp-row"><span id="gpLabel">发送中…</span><span id="gpPct"></span></div>
    <div class="progress"><i id="gpBar"></i></div>
  </div>

  <main>
    <!-- 状态 / 日志 -->
    <section id="tab-status">
      <div class="card">
        <h3>运行状态</h3>
        <div class="stat-grid">
          <div class="stat"><div class="k">发送通道</div><div class="v" id="stBusy">空闲</div></div>
          <div class="stat"><div class="k">定时执行</div><div class="v" id="stSchedule">空闲</div></div>
          <div class="stat"><div class="k">监控执行</div><div class="v" id="stMonitor">空闲</div></div>
          <div class="stat"><div class="k">服务端口</div><div class="v" id="stPort">-</div></div>
        </div>
      </div>
      <div class="card">
        <h3>运行日志 <span class="badge" id="logCount">0</span></h3>
        <div class="row" style="margin-bottom:8px">
          <button id="logRefresh" class="ghost">刷新</button>
          <button id="logClear" class="ghost">清屏</button>
          <span class="muted small">每 3 秒自动刷新</span>
        </div>
        <div id="logbox" class="logbox"></div>
      </div>
    </section>

    <!-- 手动发送 -->
    <section id="tab-send" class="hidden">
      <div class="card">
        <div class="row" style="justify-content:space-between">
          <h3 style="margin:0">电脑端已保存的配置</h3>
          <button id="profileRefresh" class="ghost">刷新</button>
        </div>
        <div class="muted small" style="margin-bottom:8px">
          来自电脑端「数据发送 → 最近配置」。可直接按原配置发送，或载入电脑端界面编辑。
        </div>
        <div id="profileList"></div>
      </div>
      <div class="card">
        <h3>1. 数据来源</h3>
        <div class="row" style="margin-bottom:8px">
          <label class="chip"><input type="radio" name="src" value="upload" checked>上传表格</label>
          <label class="chip"><input type="radio" name="src" value="cloud">金山在线表格链接</label>
        </div>
        <div id="srcUpload" class="row">
          <input type="file" id="dataFile" accept=".xlsx,.xls,.csv" class="grow">
          <button id="uploadBtn" class="primary">上传并预览</button>
        </div>
        <div id="srcCloud" class="row hidden">
          <input type="text" id="cloudLink" class="grow" placeholder="粘贴 kdocs.cn 分享链接或文档 ID">
          <button id="cloudBtn" class="primary">读取预览</button>
        </div>
        <div id="srcInfo" class="muted small" style="margin-top:8px"></div>
      </div>

      <div id="configCard" class="card hidden">
        <h3>2. 表格设置</h3>
        <div class="row">
          <label class="field grow"><span>工作表</span><select id="sheetSel"></select></label>
          <label class="field" style="width:110px"><span>表头行</span><input type="number" id="headerRow" value="1" min="1"></label>
        </div>
        <label class="field"><span>人名所在列</span><select id="nameCol"></select></label>
        <div class="field"><span>提取列（勾选要发送的数据列）</span><div id="extractCols" class="chips"></div></div>
        <div class="field">
          <span>接收人（不选=全部）</span>
          <div class="row" style="margin-bottom:6px">
            <button id="selAll" class="ghost">全选</button>
            <button id="selNone" class="ghost">清空</button>
            <span class="muted small" id="nameHint"></span>
          </div>
          <div id="namesBox" class="names"></div>
        </div>
      </div>

      <div id="sendCard" class="card hidden">
        <h3>3. 发送内容与参数</h3>
        <div class="field"><span>发送内容（按下列顺序）</span>
          <div class="chips">
            <label class="chip"><input type="checkbox" id="ordText" checked>表格文字</label>
            <label class="chip"><input type="checkbox" id="ordImage" checked>表格图片</label>
            <label class="chip"><input type="checkbox" id="ordRegion">指定区域截图</label>
            <label class="chip"><input type="checkbox" id="ordCustom">自定义消息</label>
            <label class="chip"><input type="checkbox" id="ordAttach">附加文件</label>
          </div>
        </div>
        <label class="field"><span>自定义消息</span><textarea id="customMsg" placeholder="勾选「自定义消息」后需要填写"></textarea></label>
        <div class="row">
          <label class="field grow"><span>截图区域（如 A1:F20，留空=整表）</span><input type="text" id="regionRange" placeholder="A1:F20"></label>
          <label class="field grow"><span>截图工作表</span><select id="regionSheet"></select></label>
        </div>
        <div class="row">
          <input type="file" id="attachFile" class="grow">
          <button id="attachBtn" class="ghost">上传附件</button>
          <span id="attachInfo" class="muted small"></span>
        </div>
        <div class="field" style="margin-top:10px"><span>群聊 @</span>
          <div class="row">
            <label class="chip"><input type="checkbox" id="atEnable">启用 @</label>
            <input type="text" id="atMembers" class="grow" placeholder="成员昵称，用 / 或 , 分隔">
          </div>
          <div class="row" style="margin-top:6px">
            <label class="chip"><input type="checkbox" id="atAll">@所有人</label>
            <select id="atAttach" style="max-width:200px">
              <option value="text">@ 随文字消息</option>
              <option value="first">@ 随首条消息</option>
            </select>
          </div>
        </div>
        <div class="row">
          <label class="field grow"><span>发送通道</span><select id="channel">
            <option value="desktop">桌面微信</option>
            <option value="agent">Agent WeChat API</option>
          </select></label>
          <label class="field" style="width:120px"><span>发送间隔(秒)</span><input type="number" id="sendInterval" value="0.5" step="0.1" min="0"></label>
          <label class="field" style="width:120px"><span>聊天间隔(秒)</span><input type="number" id="chatDelay" value="0.3" step="0.1" min="0"></label>
        </div>
        <button id="sendBtn" class="primary" style="width:100%;margin-top:6px">开始发送</button>
        <div id="jobBox" class="hidden" style="margin-top:12px">
          <div id="jobText" class="small"></div>
          <div class="progress"><i id="jobBar"></i></div>
        </div>
        <div id="sendErr" class="err" style="margin-top:8px"></div>
      </div>
    </section>

    <!-- 定时任务 -->
    <section id="tab-schedule" class="hidden">
      <div class="card">
        <div class="row" style="justify-content:space-between">
          <h3 style="margin:0">定时任务</h3>
          <button id="schedRefresh" class="ghost">刷新</button>
        </div>
        <div id="schedList"></div>
      </div>
    </section>

    <!-- 监控任务 -->
    <section id="tab-monitor" class="hidden">
      <div class="card">
        <div class="row" style="justify-content:space-between">
          <h3 style="margin:0">监控任务</h3>
          <button id="monRefresh" class="ghost">刷新</button>
        </div>
        <div id="monList"></div>
      </div>
    </section>
  </main>
</div>

<!-- 配置编辑弹层 -->
<div id="editor" class="modal hidden">
  <div class="sheet">
    <h3>编辑配置：<span id="edName"></span></h3>
    <label class="field"><span>数据源</span>
      <select id="edSource">
        <option value="local">本地 Excel</option>
        <option value="wps_cloud">金山在线文档</option>
      </select>
    </label>
    <label class="field" id="edLocalRow"><span>Excel 文件路径（电脑上的绝对路径）</span>
      <input type="text" id="edExcelPath" placeholder="C:\data\名单.xlsx">
    </label>
    <label class="field" id="edCloudRow"><span>云文档链接或 ID</span>
      <input type="text" id="edCloudId" placeholder="粘贴 kdocs.cn 链接或文档 ID">
    </label>
    <label class="field" id="edCloudNameRow"><span>云文档名称（可选）</span>
      <input type="text" id="edCloudName">
    </label>
    <div class="row">
      <label class="field grow"><span>工作表</span><input type="text" id="edSheet"></label>
      <label class="field" style="width:110px"><span>表头行</span><input type="number" id="edHeaderRow" min="1"></label>
    </div>
    <div class="row">
      <label class="field grow"><span>人名列</span><input type="text" id="edNameCol"></label>
      <label class="field grow"><span>微信列（可选）</span><input type="text" id="edWechatCol"></label>
    </div>
    <label class="field"><span>提取列（逗号分隔）</span><input type="text" id="edColumns"></label>
    <label class="field"><span>接收人（手动兜底，可留空）</span><input type="text" id="edRecipient"></label>
    <div class="field"><span>发送内容顺序</span>
      <div class="chips">
        <label class="chip"><input type="checkbox" id="edOrdText">表格文字</label>
        <label class="chip"><input type="checkbox" id="edOrdImage">表格图片</label>
        <label class="chip"><input type="checkbox" id="edOrdRegion">指定区域截图</label>
        <label class="chip"><input type="checkbox" id="edOrdCustom">自定义消息</label>
        <label class="chip"><input type="checkbox" id="edOrdAttach">附加文件</label>
      </div>
    </div>
    <label class="field"><span>自定义消息（支持 {{}} 占位符）</span>
      <textarea id="edCustomMsg"></textarea>
    </label>
    <label class="chip"><input type="checkbox" id="edCustomEnabled">发送自定义消息</label>
    <label class="field" style="margin-top:10px"><span>附件路径（可选）</span>
      <input type="text" id="edAttach" placeholder="C:\data\附件.pdf">
    </label>
    <div class="row">
      <label class="field grow"><span>区域截图（如 A1:F20）</span><input type="text" id="edRegion"></label>
      <label class="field grow"><span>区域工作表</span><input type="text" id="edRegionSheet"></label>
    </div>
    <div class="field"><span>群聊 @</span>
      <div class="row">
        <label class="chip"><input type="checkbox" id="edAtEnabled">启用</label>
        <input type="text" id="edAtMembers" class="grow" placeholder="成员昵称，/ 或 , 分隔">
      </div>
      <div class="row" style="margin-top:6px">
        <label class="chip"><input type="checkbox" id="edAtAll">@所有人</label>
        <select id="edAtAttach" style="max-width:180px">
          <option value="text">@ 随文字消息</option>
          <option value="first">@ 随首条消息</option>
        </select>
      </div>
    </div>
    <div class="row">
      <label class="field" style="width:120px"><span>发送间隔(秒)</span><input type="number" id="edInterval" step="0.1" min="0"></label>
      <label class="field" style="width:120px"><span>聊天间隔(秒)</span><input type="number" id="edChatDelay" step="0.1" min="0"></label>
      <label class="chip"><input type="checkbox" id="edAutoSend">加载后自动发送</label>
      <label class="chip"><input type="checkbox" id="edAgent">Agent 通道</label>
    </div>
    <div id="edErr" class="err"></div>
    <div class="row" style="justify-content:flex-end;margin-top:10px">
      <button id="edCancel" class="ghost">取消</button>
      <button id="edSave" class="primary">保存配置</button>
    </div>
  </div>
</div>

<script>
(function(){
  "use strict";
  const S = {
    uploadId: null,
    sheets: [],
    headers: [],
    names: [],
    selected: new Set(),
    extract: new Set(),
    attachmentId: null,
    jobId: null,
    logLast: 0,
    logLines: 0,
    currentTab: "status",
    timers: {}
  };
  const $ = (id) => document.getElementById(id);

  function esc(t){
    return String(t == null ? "" : t).replace(/[&<>"']/g, c => ({
      "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
    }[c]));
  }

  async function api(path, opts){
    opts = opts || {};
    const res = await fetch(path, Object.assign({credentials:"same-origin"}, opts));
    let data;
    try { data = await res.json(); } catch(e){ data = {ok:false, error:"响应解析失败"}; }
    if (res.status === 401 && !data.need_login){ data.need_login = true; }
    if (data.need_login){ showLogin(); throw new Error("need_login"); }
    if (!res.ok || data.ok === false){ throw new Error(data.error || ("HTTP " + res.status)); }
    return data;
  }
  function postJSON(path, obj){
    return api(path, {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify(obj || {})
    });
  }

  /* ---------------- 登录 ---------------- */
  function showLogin(){
    $("app").classList.add("hidden");
    $("login").classList.remove("hidden");
    stopTimers();
  }
  function showApp(){
    $("login").classList.add("hidden");
    $("app").classList.remove("hidden");
    startTimers();
    refreshProgress();
  }
  async function doLogin(){
    const pwd = $("pwd").value;
    $("loginErr").textContent = "";
    if (!pwd){ $("loginErr").textContent = "请输入密码"; return; }
    $("loginBtn").disabled = true;
    try {
      await postJSON("/api/login", {password:pwd});
      $("pwd").value = "";
      showApp();
      refreshStatus();
      refreshLogs(true);
      loadProfiles();
      loadSchedules();
      loadMonitors();
    } catch(e){
      $("loginErr").textContent = e.message === "need_login" ? "登录失败" : e.message;
    } finally {
      $("loginBtn").disabled = false;
    }
  }
  $("loginBtn").addEventListener("click", doLogin);
  $("pwd").addEventListener("keydown", e => { if (e.key === "Enter") doLogin(); });
  $("logoutBtn").addEventListener("click", async () => {
    try { await postJSON("/api/logout", {}); } catch(e){}
    showLogin();
  });

  /* ---------------- Tab 切换 ---------------- */
  function switchTab(name){
    S.currentTab = name;
    document.querySelectorAll("#tabs button").forEach(b => {
      b.classList.toggle("active", b.dataset.tab === name);
    });
    ["status","send","schedule","monitor"].forEach(k => {
      $("tab-" + k).classList.toggle("hidden", k !== name);
    });
    if (name === "status") refreshStatus();
    if (name === "send") loadProfiles();
    if (name === "schedule") loadSchedules();
    if (name === "monitor") loadMonitors();
  }
  document.querySelectorAll("#tabs button").forEach(b => {
    b.addEventListener("click", () => switchTab(b.dataset.tab));
  });

  /* ---------------- 状态 ---------------- */
  async function refreshStatus(){
    try {
      const d = await api("/api/status");
      $("stBusy").textContent = d.manual_busy ? "发送中" : "空闲";
      $("stSchedule").textContent = d.schedule_busy ? "执行中" : "空闲";
      $("stMonitor").textContent = d.monitor_busy ? "执行中" : "空闲";
      $("stPort").textContent = d.port || "-";
      if (d.lan_url) $("lanUrl").textContent = d.lan_url;
    } catch(e){}
  }

  async function refreshProgress(){
    try {
      const d = await api("/api/progress");
      const box = $("gp");
      if (!d.active){ box.classList.add("hidden"); return; }
      box.classList.remove("hidden");
      $("gpLabel").textContent = d.text || "发送中…";
      const bar = $("gpBar");
      if (d.total > 0){
        bar.classList.remove("indet");
        bar.style.width = (d.percent || 0) + "%";
        $("gpPct").textContent = (d.current || 0) + "/" + d.total +
          "（" + (d.percent || 0) + "%）";
      } else {
        bar.classList.add("indet");
        $("gpPct").textContent = "";
      }
    } catch(e){}
  }

  async function refreshLogs(reset){
    if (reset){ S.logLast = 0; S.logLines = 0; $("logbox").textContent = ""; }
    try {
      const d = await api("/api/logs?since=" + S.logLast + "&limit=200");
      if (!d.items || !d.items.length){ S.logLast = d.last || S.logLast; return; }
      const box = $("logbox");
      const frag = document.createDocumentFragment();
      d.items.forEach(it => {
        const line = document.createElement("div");
        line.className = "logline";
        line.innerHTML = '<span class="t">' + esc(it.ts || "") + '</span>' + esc(it.msg || "");
        frag.appendChild(line);
        S.logLines++;
      });
      box.appendChild(frag);
      while (box.childNodes.length > 800) box.removeChild(box.firstChild);
      box.scrollTop = box.scrollHeight;
      S.logLast = d.last;
      $("logCount").textContent = S.logLines;
    } catch(e){}
  }
  $("logRefresh").addEventListener("click", () => refreshLogs(false));
  $("logClear").addEventListener("click", () => refreshLogs(true));

  /* ---------------- 数据来源 ---------------- */
  document.querySelectorAll('input[name="src"]').forEach(r => {
    r.addEventListener("change", () => {
      const v = document.querySelector('input[name="src"]:checked').value;
      $("srcUpload").classList.toggle("hidden", v !== "upload");
      $("srcCloud").classList.toggle("hidden", v !== "cloud");
    });
  });

  $("uploadBtn").addEventListener("click", async () => {
    const file = $("dataFile").files[0];
    if (!file){ alert("请先选择表格文件"); return; }
    $("uploadBtn").disabled = true;
    $("srcInfo").textContent = "正在上传并解析…";
    try {
      const buf = await file.arrayBuffer();
      const d = await api("/api/upload?kind=data&filename=" + encodeURIComponent(file.name), {
        method:"POST",
        headers:{"Content-Type":"application/octet-stream"},
        body:buf
      });
      applyPreview(d);
    } catch(e){
      $("srcInfo").textContent = "上传失败：" + e.message;
    } finally {
      $("uploadBtn").disabled = false;
    }
  });

  $("cloudBtn").addEventListener("click", async () => {
    const link = $("cloudLink").value.trim();
    if (!link){ alert("请填写在线表格链接"); return; }
    $("cloudBtn").disabled = true;
    $("srcInfo").textContent = "正在下载金山在线表格…";
    try {
      const d = await postJSON("/api/cloud/preview", {link:link});
      applyPreview(d);
    } catch(e){
      $("srcInfo").textContent = "读取失败：" + e.message;
    } finally {
      $("cloudBtn").disabled = false;
    }
  });

  function applyPreview(d){
    S.uploadId = d.upload_id;
    S.sheets = d.sheets || [];
    S.headers = d.headers || [];
    S.names = [];
    S.selected.clear();
    $("srcInfo").textContent = "已加载：" + (d.filename || "") +
      "（共 " + (d.total_rows || 0) + " 行数据）";

    const sheetSel = $("sheetSel");
    sheetSel.innerHTML = "";
    S.sheets.forEach(s => {
      const o = document.createElement("option");
      o.value = s; o.textContent = s;
      sheetSel.appendChild(o);
    });
    if (d.sheet) sheetSel.value = d.sheet;

    const regionSel = $("regionSheet");
    regionSel.innerHTML = '<option value="">（当前数据表）</option>';
    S.sheets.forEach(s => {
      const o = document.createElement("option");
      o.value = s; o.textContent = s;
      regionSel.appendChild(o);
    });

    $("headerRow").value = d.header_row || 1;
    renderColumns();
    $("configCard").classList.remove("hidden");
    $("sendCard").classList.remove("hidden");
    loadNames();
  }

  function renderColumns(){
    const nameSel = $("nameCol");
    nameSel.innerHTML = "";
    S.headers.forEach(h => {
      const o = document.createElement("option");
      o.value = h; o.textContent = h;
      nameSel.appendChild(o);
    });
    const box = $("extractCols");
    box.innerHTML = "";
    S.extract.clear();
    S.headers.forEach(h => {
      if (h === "") return;
      S.extract.add(h);
      const label = document.createElement("label");
      label.className = "chip";
      label.innerHTML = '<input type="checkbox" checked>' + esc(h);
      label.querySelector("input").addEventListener("change", ev => {
        if (ev.target.checked) S.extract.add(h); else S.extract.delete(h);
      });
      box.appendChild(label);
    });
  }

  async function reloadPreview(){
    if (!S.uploadId) return;
    try {
      const d = await postJSON("/api/preview", {
        upload_id:S.uploadId,
        sheet:$("sheetSel").value,
        header_row:parseInt($("headerRow").value || "1", 10)
      });
      S.headers = d.headers || [];
      $("srcInfo").textContent = "已加载：" + (d.filename || "") +
        "（共 " + (d.total_rows || 0) + " 行数据）";
      renderColumns();
      loadNames();
    } catch(e){
      $("srcInfo").textContent = "切换失败：" + e.message;
    }
  }
  $("sheetSel").addEventListener("change", reloadPreview);
  $("headerRow").addEventListener("change", reloadPreview);
  $("nameCol").addEventListener("change", loadNames);

  async function loadNames(){
    if (!S.uploadId) return;
    const nameCol = $("nameCol").value;
    if (!nameCol) return;
    $("nameHint").textContent = "正在读取人员…";
    try {
      const d = await postJSON("/api/names", {
        upload_id:S.uploadId,
        sheet:$("sheetSel").value,
        header_row:parseInt($("headerRow").value || "1", 10),
        name_column:nameCol
      });
      S.names = d.names || [];
      S.selected.clear();
      const box = $("namesBox");
      box.innerHTML = "";
      S.names.forEach(n => {
        const label = document.createElement("label");
        label.innerHTML = '<input type="checkbox">' + esc(n);
        label.querySelector("input").addEventListener("change", ev => {
          if (ev.target.checked) S.selected.add(n); else S.selected.delete(n);
        });
        box.appendChild(label);
      });
      $("nameHint").textContent = "共 " + S.names.length + " 人（不勾选=全部）";
    } catch(e){
      $("nameHint").textContent = "读取人员失败：" + e.message;
    }
  }
  $("selAll").addEventListener("click", () => {
    S.selected = new Set(S.names);
    $("namesBox").querySelectorAll("input").forEach(i => i.checked = true);
  });
  $("selNone").addEventListener("click", () => {
    S.selected.clear();
    $("namesBox").querySelectorAll("input").forEach(i => i.checked = false);
  });

  /* ---------------- 附件 ---------------- */
  $("attachBtn").addEventListener("click", async () => {
    const file = $("attachFile").files[0];
    if (!file){ alert("请先选择附件文件"); return; }
    $("attachBtn").disabled = true;
    $("attachInfo").textContent = "上传中…";
    try {
      const buf = await file.arrayBuffer();
      const d = await api("/api/upload?kind=attachment&filename=" + encodeURIComponent(file.name), {
        method:"POST",
        headers:{"Content-Type":"application/octet-stream"},
        body:buf
      });
      S.attachmentId = d.upload_id;
      $("attachInfo").textContent = "已上传：" + (d.filename || "");
    } catch(e){
      $("attachInfo").textContent = "上传失败：" + e.message;
    } finally {
      $("attachBtn").disabled = false;
    }
  });

  /* ---------------- 发送 ---------------- */
  function collectOrder(){
    const order = [];
    if ($("ordText").checked) order.push("text");
    if ($("ordImage").checked) order.push("image");
    if ($("ordRegion").checked) order.push("region");
    if ($("ordCustom").checked) order.push("custom");
    if ($("ordAttach").checked) order.push("attachment");
    return order;
  }

  $("sendBtn").addEventListener("click", async () => {
    $("sendErr").textContent = "";
    if (!S.uploadId){ $("sendErr").textContent = "请先上传或读取表格"; return; }
    const order = collectOrder();
    if (!order.length){ $("sendErr").textContent = "请至少勾选一种发送内容"; return; }
    if (order.indexOf("custom") >= 0 && !$("customMsg").value.trim()){
      $("sendErr").textContent = "已勾选自定义消息，但内容为空";
      return;
    }
    const payload = {
      source:{type:"upload", upload_id:S.uploadId},
      sheet:$("sheetSel").value,
      header_row:parseInt($("headerRow").value || "1", 10),
      name_column:$("nameCol").value,
      extract_columns:Array.from(S.extract),
      recipients:Array.from(S.selected),
      custom_msg:$("customMsg").value,
      send_order:order,
      region:{
        sheet:$("regionSheet").value,
        range:$("regionRange").value.trim()
      },
      attachment_upload_id:S.attachmentId || "",
      at:{
        enabled:!!$("atEnable").checked,
        members:$("atMembers").value,
        all:!!$("atAll").checked,
        attach:$("atAttach").value
      },
      channel:$("channel").value,
      send_interval:parseFloat($("sendInterval").value || "0.5"),
      chat_delay:parseFloat($("chatDelay").value || "0.3")
    };
    $("sendBtn").disabled = true;
    $("jobBox").classList.remove("hidden");
    $("jobText").textContent = "正在启动发送…";
    $("jobBar").style.width = "0%";
    try {
      const d = await postJSON("/api/send", payload);
      S.jobId = d.job_id;
      $("jobText").textContent = "已开始发送，共 " + (d.total || 0) + " 个任务";
      pollJob();
      refreshLogs(false);
    } catch(e){
      $("sendErr").textContent = "发送失败：" + e.message;
      $("jobText").textContent = "";
      $("sendBtn").disabled = false;
    }
  });

  async function pollJob(){
    if (!S.jobId){ $("sendBtn").disabled = false; return; }
    try {
      const d = await api("/api/jobs/" + encodeURIComponent(S.jobId));
      const job = d.job || {};
      const total = job.total || 0;
      const cur = job.current || 0;
      const pct = total > 0 ? Math.min(100, Math.round(cur * 100 / total)) : 0;
      $("jobBar").style.width = pct + "%";
      $("jobText").textContent = (job.message || "发送中") +
        (total ? "（" + cur + "/" + total + "）" : "");
      if (job.status === "running"){
        setTimeout(pollJob, 1500);
      } else {
        $("sendBtn").disabled = false;
        S.jobId = null;
        refreshLogs(false);
      }
    } catch(e){
      $("sendBtn").disabled = false;
      S.jobId = null;
    }
  }

  /* ---------------- 电脑端已保存配置 ---------------- */
  async function loadProfiles(){
    const box = $("profileList");
    box.innerHTML = '<div class="muted small">加载中…</div>';
    try {
      const d = await api("/api/profiles");
      box.innerHTML = "";
      if (!d.items || !d.items.length){
        box.innerHTML = '<div class="muted small">电脑端还没有保存过配置</div>';
        return;
      }
      d.items.forEach(p => {
        const row = document.createElement("div");
        row.className = "list-item";
        const meta = [
          p.excel || "（无文件）",
          p.sheet ? ("工作表 " + p.sheet) : "",
          p.name_column ? ("人名列 " + p.name_column) : "",
          p.columns ? (p.columns + " 列") : "",
          p.recipients ? ("接收人 " + p.recipients) : "",
          p.agent_api ? "Agent 通道" : ""
        ].filter(Boolean).join(" · ");
        row.innerHTML =
          '<div class="main"><div class="name">' + esc(p.name) +
          (p.auto_send ? ' <span class="badge">加载后自动发送</span>' : '') +
          '</div><div class="meta">' + esc(meta) +
          '</div><div class="meta">' + esc(p.saved_at || "") + '</div></div>' +
          '<button class="primary">立即发送</button>' +
          '<button class="ghost">编辑</button>' +
          '<button class="ghost">载入电脑端</button>';
        const btns = row.querySelectorAll("button");
        btns[0].addEventListener("click", async ev => {
          ev.target.disabled = true;
          try {
            const r = await postJSON("/api/profiles/run", {id:p.id});
            alert(r.message || ("已开始执行：" + p.name));
            refreshLogs(false);
            switchTab("status");
          } catch(e){
            alert("执行失败：" + e.message);
          } finally {
            ev.target.disabled = false;
          }
        });
        btns[1].addEventListener("click", () => openProfileEditor(p.id));
        btns[2].addEventListener("click", async ev => {
          ev.target.disabled = true;
          try {
            const r = await postJSON("/api/profiles/edit", {id:p.id});
            alert((r.message || "已载入") + "\n请到电脑端界面继续编辑。");
          } catch(e){
            alert("载入失败：" + e.message);
          } finally {
            ev.target.disabled = false;
          }
        });
        box.appendChild(row);
      });
    } catch(e){
      box.innerHTML = '<div class="err">加载失败：' + esc(e.message) + '</div>';
    }
  }
  $("profileRefresh").addEventListener("click", loadProfiles);

  /* ---------------- 配置编辑弹层 ---------------- */
  let editingProfileId = null;
  function toggleEdSource(){
    const v = $("edSource").value;
    $("edLocalRow").classList.toggle("hidden", v !== "local");
    $("edCloudRow").classList.toggle("hidden", v !== "wps_cloud");
    $("edCloudNameRow").classList.toggle("hidden", v !== "wps_cloud");
  }
  $("edSource").addEventListener("change", toggleEdSource);

  async function openProfileEditor(id){
    editingProfileId = id;
    try {
      const d = await api("/api/profiles/detail?id=" + encodeURIComponent(id));
      const p = d.profile;
      $("edName").textContent = p.name;
      $("edSource").value = p.source || "local";
      toggleEdSource();
      $("edExcelPath").value = p.excel_path || "";
      $("edCloudId").value = p.cloud_file_id || "";
      $("edCloudName").value = p.cloud_file_name || "";
      $("edSheet").value = p.sheet || "";
      $("edHeaderRow").value = p.header_row || 1;
      $("edNameCol").value = p.name_column || "";
      $("edWechatCol").value = p.wechat_column || "";
      $("edColumns").value = (p.extract_columns || []).join(",");
      $("edRecipient").value = p.manual_recipient || "";
      const order = p.send_order || [];
      $("edOrdText").checked = order.indexOf("text") >= 0;
      $("edOrdImage").checked = order.indexOf("image") >= 0;
      $("edOrdRegion").checked = order.indexOf("region") >= 0;
      $("edOrdCustom").checked = order.indexOf("custom") >= 0;
      $("edOrdAttach").checked = order.indexOf("attachment") >= 0;
      $("edCustomMsg").value = p.custom_message || "";
      $("edCustomEnabled").checked = !!p.custom_message_enabled;
      $("edAttach").value = p.attachment || "";
      $("edRegion").value = p.region_range || "";
      $("edRegionSheet").value = p.region_sheet || "";
      $("edAtEnabled").checked = !!p.at_enabled;
      $("edAtMembers").value = p.at_members || "";
      $("edAtAll").checked = !!p.at_all;
      $("edAtAttach").value = p.at_attach || "text";
      $("edInterval").value = p.interval;
      $("edChatDelay").value = p.chat_delay;
      $("edAutoSend").checked = !!p.auto_send;
      $("edAgent").checked = !!p.agent_api;
      $("edErr").textContent = "";
      $("editor").classList.remove("hidden");
    } catch(e){
      alert("读取配置失败：" + e.message);
    }
  }
  function closeEditor(){
    $("editor").classList.add("hidden");
    editingProfileId = null;
  }
  $("edCancel").addEventListener("click", closeEditor);
  $("editor").addEventListener("click", ev => {
    if (ev.target === $("editor")) closeEditor();
  });

  $("edSave").addEventListener("click", async () => {
    if (!editingProfileId) return;
    const order = [];
    if ($("edOrdText").checked) order.push("text");
    if ($("edOrdImage").checked) order.push("image");
    if ($("edOrdRegion").checked) order.push("region");
    if ($("edOrdCustom").checked) order.push("custom");
    if ($("edOrdAttach").checked) order.push("attachment");
    const body = {
      id: editingProfileId,
      source: $("edSource").value,
      excel_path: $("edExcelPath").value,
      cloud_file_id: $("edCloudId").value,
      cloud_file_name: $("edCloudName").value,
      sheet: $("edSheet").value,
      header_row: parseInt($("edHeaderRow").value || "1", 10),
      name_column: $("edNameCol").value,
      wechat_column: $("edWechatCol").value,
      extract_columns: $("edColumns").value,
      manual_recipient: $("edRecipient").value,
      custom_message_enabled: $("edCustomEnabled").checked,
      custom_message: $("edCustomMsg").value,
      auto_send: $("edAutoSend").checked,
      interval: parseFloat($("edInterval").value || "0.5"),
      chat_delay: parseFloat($("edChatDelay").value || "0.3"),
      send_order: order,
      attachment: $("edAttach").value,
      region_range: $("edRegion").value,
      region_sheet: $("edRegionSheet").value,
      at_enabled: $("edAtEnabled").checked,
      at_members: $("edAtMembers").value,
      at_all: $("edAtAll").checked,
      at_attach: $("edAtAttach").value,
      agent_api: $("edAgent").checked,
    };
    $("edErr").textContent = "";
    $("edSave").disabled = true;
    try {
      const r = await postJSON("/api/profiles/save", body);
      alert(r.message || "配置已保存");
      closeEditor();
      loadProfiles();
    } catch(e){
      $("edErr").textContent = "保存失败：" + e.message;
    } finally {
      $("edSave").disabled = false;
    }
  });

  /* ---------------- 定时任务 ---------------- */
  async function loadSchedules(){
    try {
      const d = await api("/api/schedule");
      const box = $("schedList");
      box.innerHTML = "";
      if (!d.items || !d.items.length){
        box.innerHTML = '<div class="muted small">暂无定时任务</div>';
        return;
      }
      d.items.forEach(t => {
        const row = document.createElement("div");
        row.className = "list-item";
        row.innerHTML =
          '<div class="main"><div class="name">' + esc(t.name) + '</div>' +
          '<div class="meta">' + esc(t.rule) + ' · ' + esc((t.times || []).join("、") || "无时间点") +
          ' · ' + esc(t.target) + '</div></div>' +
          '<label class="switch"><input type="checkbox"' + (t.enabled ? " checked" : "") + '>启用</label>' +
          '<button class="ghost">立即执行</button>';
        row.querySelector(".switch input").addEventListener("change", async ev => {
          try {
            await postJSON("/api/schedule/" + encodeURIComponent(t.id) + "/enabled",
              {enabled:ev.target.checked});
          } catch(e){
            ev.target.checked = !ev.target.checked;
            alert("操作失败：" + e.message);
          }
        });
        row.querySelector("button").addEventListener("click", async ev => {
          ev.target.disabled = true;
          try {
            const r = await postJSON("/api/schedule/" + encodeURIComponent(t.id) + "/run", {});
            alert(r.message || "已开始执行");
          } catch(e){
            alert("执行失败：" + e.message);
          } finally {
            ev.target.disabled = false;
          }
        });
        box.appendChild(row);
      });
    } catch(e){}
  }
  $("schedRefresh").addEventListener("click", loadSchedules);

  /* ---------------- 监控任务 ---------------- */
  async function loadMonitors(){
    try {
      const d = await api("/api/monitor");
      const box = $("monList");
      box.innerHTML = "";
      if (!d.items || !d.items.length){
        box.innerHTML = '<div class="muted small">暂无监控任务</div>';
        return;
      }
      d.items.forEach(t => {
        const row = document.createElement("div");
        row.className = "list-item";
        row.innerHTML =
          '<div class="main"><div class="name">' + esc(t.name) + '</div>' +
          '<div class="meta">' + esc(t.rule) + ' · 每 ' + esc(t.interval) + ' 分钟 · ' +
          esc(t.target) + '</div></div>' +
          '<label class="switch"><input type="checkbox"' + (t.enabled ? " checked" : "") + '>启用</label>' +
          '<button class="ghost">立即执行</button>';
        row.querySelector(".switch input").addEventListener("change", async ev => {
          try {
            await postJSON("/api/monitor/" + encodeURIComponent(t.id) + "/enabled",
              {enabled:ev.target.checked});
          } catch(e){
            ev.target.checked = !ev.target.checked;
            alert("操作失败：" + e.message);
          }
        });
        row.querySelector("button").addEventListener("click", async ev => {
          ev.target.disabled = true;
          try {
            const r = await postJSON("/api/monitor/" + encodeURIComponent(t.id) + "/run", {});
            alert(r.message || "已开始执行");
          } catch(e){
            alert("执行失败：" + e.message);
          } finally {
            ev.target.disabled = false;
          }
        });
        box.appendChild(row);
      });
    } catch(e){}
  }
  $("monRefresh").addEventListener("click", loadMonitors);

  /* ---------------- 定时刷新 ---------------- */
  function startTimers(){
    stopTimers();
    S.timers.status = setInterval(() => {
      if (S.currentTab === "status") refreshStatus();
    }, 3000);
    S.timers.logs = setInterval(() => {
      if (S.currentTab === "status") refreshLogs(false);
    }, 3000);
    S.timers.progress = setInterval(refreshProgress, 1500);
  }
  function stopTimers(){
    Object.keys(S.timers).forEach(k => clearInterval(S.timers[k]));
    S.timers = {};
  }

  /* ---------------- 启动 ---------------- */
  (async function boot(){
    try {
      await api("/api/service");
      const st = await api("/api/status");
      showApp();
      refreshStatus();
      refreshLogs(true);
      loadProfiles();
      loadSchedules();
      loadMonitors();
      if (st.lan_url) $("lanUrl").textContent = st.lan_url;
    } catch(e){
      if (e.message === "need_login") showLogin();
      else showLogin();
    }
  })();
})();
</script>
</body>
</html>
"""
