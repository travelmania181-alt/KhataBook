const $=id=>document.getElementById(id);
const categories={expense:["Food & groceries","Transport","Bills & utilities","Shopping","Health","Education","Rent","Entertainment","Travel","Other"],income:["Salary","Business","Freelance","Interest","Gift","Refund","Other"]};
const icons={"Food & groceries":"🥑","Transport":"🛵","Bills & utilities":"💡","Shopping":"🛍️","Health":"🩺","Education":"📚","Rent":"🏠","Entertainment":"🎧","Travel":"🧳","Other":"✦","Salary":"💼","Business":"📈","Freelance":"💻","Interest":"🏦","Gift":"🎁","Refund":"↩"};
let data=JSON.parse(localStorage.getItem("khataBookData")||"[]"), profile=JSON.parse(localStorage.getItem("khataBookProfile")||"null"), formType="expense", showAll=false;
const today=()=>new Date().toISOString().slice(0,10), monthNow=()=>today().slice(0,7), money=n=>new Intl.NumberFormat("en-IN",{style:"currency",currency:"INR",maximumFractionDigits:2}).format(n||0);
function save(){localStorage.setItem("khataBookData",JSON.stringify(data));render()}
function selectedMonth(){return $("reportMonth").value||monthNow()}
function monthData(m){return data.filter(t=>t.date.startsWith(m))}
function sum(arr,type){return arr.filter(t=>t.type===type).reduce((s,t)=>s+Number(t.amount),0)}
function escapeHTML(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
function render(){if(!profile)return;let m=selectedMonth(),md=monthData(m),inc=sum(md,"income"),exp=sum(md,"expense");$("monthLabel").textContent=new Date(m+"-02").toLocaleDateString("en-IN",{month:"short",year:"numeric"});$("netBalance").textContent=money(inc-exp);$("incomeTotal").textContent=money(inc);$("expenseTotal").textContent=money(exp);$("reportIncome").textContent=money(inc);$("reportExpense").textContent=money(exp);let max=Math.max(inc,exp,1);$("incomeBar").style.width=(inc/max*100)+"%";$("expenseBar").style.width=(exp/max*100)+"%";
let list=[...data].sort((a,b)=>(b.date+b.id).localeCompare(a.date+a.id));if(!showAll)list=list.slice(0,6);$("transactions").innerHTML=list.length?list.map(t=>`<article class="tx"><div class="tx-icon">${icons[t.category]||"✦"}</div><div class="tx-main"><b>${escapeHTML(t.note)}</b><small>${escapeHTML(t.category)} · ${new Date(t.date+"T12:00:00").toLocaleDateString("en-IN",{day:"numeric",month:"short",year:"numeric"})} · ${escapeHTML(t.payment)}</small></div><div class="tx-amount ${t.type==="income"?"positive":"negative"}">${t.type==="income"?"+":"−"}${money(t.amount)}<br><button class="text-btn" data-delete="${t.id}" style="font-size:11px;padding:0">Delete</button></div></article>`).join(""):'<div class="empty">No transactions yet. Add your first income or expense above.</div>';
const cats={};md.filter(t=>t.type==="expense").forEach(t=>cats[t.category]=(cats[t.category]||0)+Number(t.amount));let entries=Object.entries(cats).sort((a,b)=>b[1]-a[1]);$("categoryBreakdown").innerHTML=entries.length?entries.map(([c,v])=>`<div class="category-row"><span>${icons[c]||"✦"}</span><span>${escapeHTML(c)}<br><small>${Math.round(v/(exp||1)*100)}% of expenses</small></span><b>${money(v)}</b></div>`).join(""):'<div class="empty">Your spending categories will appear here.</div>';
}
function showApp(){ $("auth").hidden=true;$("main").hidden=false;$("greetingName").textContent=profile.name||"there";$("reportMonth").value=monthNow();render()}
function openForm(type){formType=type;$("txForm").reset();$("date").value=today();$("formTitle").textContent=type==="income"?"Add income":"Add expense";document.querySelectorAll("[data-formtype]").forEach(b=>b.classList.toggle("selected",b.dataset.formtype===type));$("category").innerHTML=categories[type].map(c=>`<option>${escapeHTML(c)}</option>`).join("");$("txDialog").showModal()}
document.querySelectorAll("[data-type]").forEach(b=>b.addEventListener("click",()=>openForm(b.dataset.type)));
document.querySelectorAll("[data-formtype]").forEach(b=>b.addEventListener("click",()=>{formType=b.dataset.formtype;$("formTitle").textContent=formType==="income"?"Add income":"Add expense";document.querySelectorAll("[data-formtype]").forEach(x=>x.classList.toggle("selected",x===b));$("category").innerHTML=categories[formType].map(c=>`<option>${escapeHTML(c)}</option>`).join("")}));
$("txForm").addEventListener("submit",e=>{e.preventDefault();const amount=Number($("amount").value);if(!(amount>0))return;data.push({id:Date.now().toString(),type:formType,amount,note:$("note").value.trim(),category:$("category").value,date:$("date").value,payment:$("payment").value});$("txDialog").close();save()});
$("closeDialog").onclick=$("cancelDialog").onclick=()=>$("txDialog").close();
$("transactions").addEventListener("click",e=>{let id=e.target.dataset.delete;if(id&&confirm("Delete this transaction?")){data=data.filter(t=>t.id!==id);save()}});
$("reportMonth").addEventListener("change",render);$("allBtn").onclick=()=>{showAll=!showAll;$("allBtn").textContent=showAll?"Show recent":"View all";render()};

function download(type) {
  const m = selectedMonth();
  const arr = monthData(m).filter(t => t.type === type);

  const rows = [
    ["Date", "Type", "Description", "Category", "Payment method", "Amount (INR)"],
    ...arr.map(t => [
      t.date,
      type,
      t.note,
      t.category,
      t.payment,
      Number(t.amount)
    ])
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(rows);

  // Convert ISO dates (YYYY-MM-DD) to genuine Excel date values.
  arr.forEach((t, index) => {
    const cell = XLSX.utils.encode_cell({ r: index + 1, c: 0 });
    const [year, month, day] = t.date.split("-").map(Number);

    worksheet[cell] = {
      t: "d",
      v: new Date(Date.UTC(year, month - 1, day)),
      z: "dd-mm-yyyy"
    };
  });

  worksheet["!cols"] = [
    { wch: 14 },
    { wch: 12 },
    { wch: 30 },
    { wch: 22 },
    { wch: 20 },
    { wch: 18 }
  ];

  // Format amounts as numeric currency values.
  arr.forEach((t, index) => {
    const cell = XLSX.utils.encode_cell({ r: index + 1, c: 5 });
    if (worksheet[cell]) {
      worksheet[cell].z = '#,##0.00';
    }
  });

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, type);

  XLSX.writeFile(workbook, `KhataBook-${type}-${m}.xlsx`);
}

$("incomeExport").onclick = () => download("income");
$("expenseExport").onclick = () => download("expense");

