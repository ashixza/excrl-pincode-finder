const $ = id => document.getElementById(id);
const fileInput = $("file"), drop = $("drop"), run = $("run");
let selectedFile = null, apiBase = localStorage.getItem("pincodeApiBase") || "https://YOUR-RENDER-SERVICE.onrender.com";

function setFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".xlsx")) { showError("Please choose an .xlsx file."); return; }
  if (file.size > 15 * 1024 * 1024) { showError("File exceeds the 15 MB limit."); return; }
  selectedFile = file; $("file-title").textContent = file.name;
  $("file-hint").textContent = `${(file.size/1024/1024).toFixed(2)} MB · Ready to process`;
  run.disabled = false; showError("");
}
fileInput.addEventListener("change", e => setFile(e.target.files[0]));
["dragenter","dragover"].forEach(ev => drop.addEventListener(ev, e => {e.preventDefault();drop.classList.add("active")}));
["dragleave","drop"].forEach(ev => drop.addEventListener(ev, e => {e.preventDefault();drop.classList.remove("active")}));
drop.addEventListener("drop", e => setFile(e.dataTransfer.files[0]));
function showError(msg){$("error").textContent=msg}
function escapeHtml(s){return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}

run.addEventListener("click", async () => {
  if (!selectedFile) return;
  let base = apiBase;
  if (base.includes("YOUR-RENDER-SERVICE")) {
    base = prompt("Enter your Render backend URL (example: https://excel-pincode-finder-api.onrender.com):", "");
    if (!base) return;
    apiBase = base.replace(/\/+$/,""); localStorage.setItem("pincodeApiBase",apiBase);
  }
  run.disabled = true; run.textContent = "Processing workbook…"; showError("");
  try {
    const form = new FormData();
    form.append("file", selectedFile);
    form.append("sheet_name", $("sheet").value || "Addresses");
    if ($("address").value.trim()) form.append("address_column", $("address").value.trim());
    if ($("pincode").value.trim()) form.append("pincode_column", $("pincode").value.trim());
    if ($("city").value.trim()) form.append("city_column", $("city").value.trim());
    const response = await fetch(`${apiBase}/process`, {method:"POST",body:form});
    const data = await response.json();
    if (!response.ok) throw new Error(data.detail || "Processing failed.");
    $("exact").textContent=data.filled_exact;
    $("fallback").textContent=data.filled_city_fallback;
    $("unresolved").textContent=data.unresolved;
    $("existing").textContent=data.already_had_pincode;
    $("rowcount").textContent=`${data.rows_checked} data rows checked`;
    $("audit").innerHTML=(data.audit||[]).map(r=>{
      const cls=r.status==="CITY_FALLBACK"?"fallback":r.status==="UNRESOLVED"?"bad":"";
      return `<tr><td>${r.row}</td><td>${escapeHtml(r.address)}</td><td>${escapeHtml(r.pincode||"—")}</td><td><span class="tag ${cls}">${escapeHtml(r.status)}</span></td></tr>`;
    }).join("");
    $("download").href=`${apiBase}/download/${data.download_token}`;
    $("results").classList.remove("hidden");
    $("results").scrollIntoView({behavior:"smooth"});
  } catch(e) { showError(e.message || "Could not connect to the backend. Check the Render URL and try again."); }
  finally {run.disabled=false;run.innerHTML='Find & fill pincodes <span>→</span>';}
});