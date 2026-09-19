const express = require('express');
const bodyParser = require('body-parser');
const { Pool } = require('pg');
const XLSX = require('xlsx');
const app = express();
app.use(bodyParser.json());

// ========== 城市分流规则 ==========
function getGroupByCity(city, district) {
  let targetCity = city;
  if (district === "简阳市") targetCity = "成都市";
  if (targetCity === "上海市") return "上海组";
  if (["苏州市","南京市","无锡市","常州市","南通市"].includes(targetCity)) return "江苏组";
  if (["杭州市","宁波市","温州市","绍兴市","嘉兴市"].includes(targetCity)) return "浙江组";
  if (["广州市","深圳市","佛山市","东莞市","珠海市","中山市","江门市"].includes(targetCity)) return "广东组";
  if (targetCity === "成都市") return "四川组";
  if (targetCity === "重庆市") return "重庆组";
  if (targetCity === "北京市") return "北京组";
  if (targetCity === "天津市") return "天津组";
  return "全国客服池";
}

// ========= 数据库配置 =========
const pool = new Pool({
  connectionString: process.env.POSTGRES_URL,
  ssl: { rejectUnauthorized: false }
});

// 初始化数据表（首次启动自动建表）
async function initDB() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS leads (
      id SERIAL PRIMARY KEY,
      username TEXT,
      phone TEXT,
      province TEXT,
      city TEXT,
      district TEXT,
      platform TEXT,
      ad_account TEXT,
      ad_plan TEXT,
      ad_unit TEXT,
      creative TEXT,
      create_time TIMESTAMP,
      budget TEXT,
      can_add_wechat TEXT,
      demand TEXT,
      assign_city TEXT,
      assign_group TEXT,
      owner TEXT,
      status TEXT DEFAULT '待跟进',
      remark TEXT,
      push_status TEXT,
      created_at TIMESTAMP DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS app_users(
      id SERIAL PRIMARY KEY,
      account TEXT UNIQUE,
      password TEXT,
      group_access TEXT
    );
    INSERT INTO app_users(account,password,group_access) VALUES('admin','${process.env.ADMIN_PASSWORD || "Admin@123456"}','all') ON CONFLICT(account) DO NOTHING;
  `);
}
initDB();

// ========= API接口 =========
// Webhook接收快手线索
app.post('/api/kuaishou-webhook', async (req, res) => {
  const body = req.body;
  const phone = body.phone;
  if(!phone) return res.status(400).json({code:0,msg:"手机号缺失"});
  // 24小时手机号去重
  const repeat = await pool.query(`SELECT id FROM leads WHERE phone = $1 AND create_time > NOW() - INTERVAL '24 hours'`,[phone]);
  if(repeat.rows.length>0) return res.json({code:0,msg:"24小时内重复线索，丢弃"});
  const city = body.city || "";
  const district = body.district || "";
  const assignGroup = getGroupByCity(city, district);
  await pool.query(`INSERT INTO leads(
    username,phone,province,city,district,platform,ad_account,ad_plan,ad_unit,creative,
    create_time,budget,can_add_wechat,demand,assign_city,assign_group,owner,status,remark,push_status
  ) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20)`,
  [body.userName,phone,body.province,city,district,"快手",body.adAccountName,body.adPlanName,body.adUnitName,body.creativeName,
  body.createTime,body.customField_budget,body.customField_canAddWechat,body.customField_demand,city,assignGroup,"","待跟进",body.remark||"","成功"]);
  res.json({code:1,msg:"线索接收成功",group:assignGroup});
});

// 登录接口
app.post('/api/login', async(req,res)=>{
  const {account,password}=req.body;
  const adminPwd = process.env.ADMIN_PASSWORD || "Admin@123456";
  if(account === "admin" && password === adminPwd){
    return res.json({success:true});
  }
  return res.json({success:false});
});

// 获取线索列表、搜索筛选
app.get('/api/get-leads', async(req,res)=>{
  const {search,group,status} = req.query;
  let sql = "SELECT * FROM leads ORDER BY create_time DESC";
  const params = [];
  const where = [];
  if(search){params.push(`%${search}%`);where.push(`(username ILIKE $${params.length} OR phone ILIKE $${params.length})`);}
  if(group){params.push(group);where.push(`assign_group = $${params.length}`);}
  if(status){params.push(status);where.push(`status = $${params.length}`);}
  if(where.length>0) sql += " WHERE "+where.join(" AND ");
  const result = await pool.query(sql,params);
  res.json(result.rows);
});

// 修改线索状态/备注
app.post('/api/lead-update',async(req,res)=>{
  const {id,status,remark,owner}=req.body;
  await pool.query(`UPDATE leads SET status=$1,remark=$2,owner=$3 WHERE id=$4`,[status,remark,owner,id]);
  res.json({success:true});
});

// Excel导出
app.get('/api/export-excel',async(req,res)=>{
  const result = await pool.query("SELECT * FROM leads ORDER BY create_time DESC");
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(result.rows);
  XLSX.utils.book_append_sheet(wb,ws,"快手线索");
  const buf = XLSX.write(wb,{type:"buffer",bookType:"xlsx"});
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition','attachment;filename=快手线索.xlsx');
  res.end(buf);
});

// ========= 前端页面（内嵌HTML =========
// 登录页
app.get('/', (req,res)=>{
  res.send(`
    <html><body style="max-width:420px;margin:100px auto;padding:20px;font-family:system-ui">
      <h2>快手线索CRM登录</h2>
      <div>账号：<input id="account" value="admin" style="width:100%;padding:8px;margin:6px 0"/></div>
      <div>密码：<input id="pwd" type="password" style="width:100%;padding:8px;margin:6px 0"/></div>
      <button onclick="login()" style="width:100%;padding:10px;background:#0070f3;color:white;border:none;border-radius:6px;margin-top:10px">登录</button>
      <script>
        async function login(){
          const res=await fetch("/api/login",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({account:document.getElementById('account').value,password:document.getElementById('pwd').value})});
          const d=await res.json();
          if(d.success) window.location="/lead-list";
          else alert("账号密码错误");
        }
      </script>
    </body></html>
  `);
});

// 线索列表页
app.get('/lead-list', (req,res)=>{
  res.send(`
    <html><body style="padding:20px;font-family:system-ui">
      <h2>快手线索列表</h2>
      <div style="margin-bottom:12px">
        <input id="search" placeholder="搜索手机号/昵称">
        <select id="groupFilter">
          <option value="">全部分组</option>
          <option>上海组</option><option>江苏组</option><option>浙江组</option><option>广东组</option>
          <option>四川组</option><option>重庆组</option><option>北京组</option><option>天津组</option><option>全国客服池</option>
        </select>
        <select id="statusFilter">
          <option value="">全部状态</option>
          <option>待跟进</option><option>跟进中</option><option>已成交</option><option>无效</option>
        </select>
        <button onclick="loadData()">查询</button>
        <button onclick="window.open('/api/export-excel')">导出Excel</button>
        <button onclick="window.location='/dashboard'">统计看板</button>
      </div>
      <table border="1" cellpadding="6" style="width:100%;border-collapse:collapse">
        <thead><tr><th>昵称</th><th>手机号</th><th>省市</th><th>分配分组</th><th>状态</th><th>留资时间</th><th>操作</th></tr></thead>
        <tbody id="tableBody"></tbody>
      </table>
      <script>
        async function loadData(){
          const s = document.getElementById('search').value;
          const g = document.getElementById('groupFilter').value;
          const st = document.getElementById('statusFilter').value;
          const q = new URLSearchParams();
          if(s) q.append("search",s);if(g) q.append("group",g);if(st) q.append("status",st);
          const res = await fetch("/api/get-leads?"+q.toString());
          const data = await res.json();
          const tb = document.getElementById("tableBody");
          tb.innerHTML = "";
          data.forEach(item=>{
            const tr = document.createElement("tr");
            tr.innerHTML = `<td>${item.username||""}</td><td>${item.phone}</td><td>${item.province||""}${item.city||""}${item.district||""}</td><td>${item.assign_group}</td><td>${item.status}</td><td>${item.create_time}</td><td><select onchange="update(${item.id},this.value)">
              <option ${item.status==='待跟进'?'selected':''}>待跟进</option>
              <option ${item.status==='跟进中'?'selected':''}>跟进中</option>
              <option ${item.status==='已成交'?'selected':''}>已成交</option>
              <option ${item.status==='无效'?'selected':''}>无效</option>
            </select></td>`;
            tb.appendChild(tr);
          })
        }
        async function update(id,status){
          await fetch("/api/lead-update",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({id,status})});
          loadData();
        }
        loadData();
      </script>
    </body></html>
  `);
});

// 统计看板页面
app.get('/dashboard', async(req,res)=>{
  const result = await pool.query("SELECT assign_group,COUNT(*) cnt FROM leads GROUP BY assign_group");
  const total = await pool.query("SELECT COUNT(*) FROM leads");
  let html = `<html><body style="padding:20px;font-family:system-ui"><h2>线索统计看板</h2><button onclick="window.location='/lead-list'">返回线索列表</button><h3>总线索：${total.rows[0].count}</h3>`;
  result.rows.forEach(r=> html += `<div>${r.assign_group}: ${r.cnt}条</div>`);
  html += `</body></html>`;
  res.send(html);
});

// 启动服务
const PORT = process.env.PORT || 3000;
app.listen(PORT, ()=>{
  console.log(`CRM服务启动`);
})
