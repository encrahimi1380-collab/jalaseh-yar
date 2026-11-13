let meetings = [];
const DAY_NAMES = ["شنبه","یکشنبه","دوشنبه","سه‌شنبه","چهارشنبه","پنجشنبه","جمعه"];

// ============ Geocoding با Neshan ============
async function geocode(address) {
  const url = `https://api.neshan.org/v4/geocoding?address=${encodeURIComponent(address)}`;
  const res = await fetch(url, { headers: { "Api-Key": CONFIG.NESHAN_API_KEY } });
  if (!res.ok) throw new Error("خطا در پیدا کردن آدرس: " + address);
  const data = await res.json();
  const loc = data.location || (data.result && data.result.location);
  if (!loc) throw new Error("مختصات پیدا نشد برای: " + address);
  return { lat: loc.y || loc.lat, lng: loc.x || loc.lng };
}

function distanceKm(a, b) {
  const R = 6371;
  const dLat = (b.lat - a.lat) * Math.PI / 180;
  const dLng = (b.lng - a.lng) * Math.PI / 180;
  const lat1 = a.lat * Math.PI / 180;
  const lat2 = b.lat * Math.PI / 180;
  const h = Math.sin(dLat/2)**2 + Math.cos(lat1)*Math.cos(lat2)*Math.sin(dLng/2)**2;
  return 2 * R * Math.asin(Math.sqrt(h));
}

function travelMinutes(km) {
  return Math.max(10, Math.round(km / 25 * 60));
}

// ============ افزودن دستی ============
document.getElementById("addBtn").onclick = async () => {
  const title = document.getElementById("title").value.trim();
  const address = document.getElementById("address").value.trim();
  const days = [...document.querySelectorAll(".days input:checked")].map(c => +c.value);
  if (!title || !address || days.length === 0) {
    alert("عنوان، آدرس و حداقل یک روز را پر کن.");
    return;
  }
  await addMeeting(title, address, days);
  document.getElementById("title").value = "";
  document.getElementById("address").value = "";
  document.querySelectorAll(".days input").forEach(c => c.checked = false);
};

// تابع مشترک برای اضافه کردن (هم دستی هم اکسل)
async function addMeeting(title, address, days) {
  try {
    const coords = await geocode(address);
    meetings.push({ id: Date.now() + Math.random(), title, address, days, lat: coords.lat, lng: coords.lng });
    renderList();
  } catch (e) {
    alert("خطا در جلسه: " + title + "\n" + e.message);
  }
}

// ============ آپلود و خواندن اکسل ============
document.getElementById("excelFile").addEventListener("change", function(e) {
  const file = e.target.files[0];
  if (!file) return;
  
  const reader = new FileReader();
  reader.onload = async function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type: 'array'});
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(worksheet);
    
    if(json.length === 0) return alert("فایل خالی است!");
    
    alert(`در حال خواندن ${json.length} جلسه از فایل... لطفاً صبر کنید.`);
    
    for (let row of json) {
      // فرض بر این است که ستون‌ها نام فارسی دارند: عنوان، آدرس، روزها
      const title = row["عنوان"] || row["Title"] || row["name"];
      const address = row["آدرس"] || row["Address"] || row["location"];
      const daysStr = row["روزها"] || row["Days"] || row["days"] || "";
      
      if(!title || !address) continue;
      
      // تبدیل روزها (مثلا "شنبه,دوشنبه" به آرایه [0,2])
      let daysArr = [];
      const dayMap = {"شنبه":0, "یکشنبه":1, "دوشنبه":2, "سه‌شنبه":3, "سه شنبه":3, "چهارشنبه":4, "پنجشنبه":5, "پنج شنبه":5};
      String(daysStr).split(/[,، ]+/).forEach(d => {
        if (dayMap[d] !== undefined) daysArr.push(dayMap[d]);
      });
      
      if (daysArr.length === 0) daysArr = [0,1,2,3,4,5]; // اگر روزی مشخص نشده بود، همه روزها
      
      await addMeeting(title, address, daysArr);
    }
    alert("تمام جلسات اکسل اضافه شدند!");
    e.target.value = ""; // ریست کردن اینپوت
  };
  reader.readAsArrayBuffer(file);
});

// ============ نمایش لیست ============
function renderList() {
  const ul = document.getElementById("meetingList");
  ul.innerHTML = "";
  meetings.forEach(m => {
    const li = document.createElement("li");
    li.innerHTML = `<span><b>${m.title}</b> — ${m.address} <small>(${m.days.map(d => DAY_NAMES[d]).join("، ")})</small></span>
      <button onclick="removeMeeting(${m.id})">حذف</button>`;
    ul.appendChild(li);
  });
}

function removeMeeting(id) {
  meetings = meetings.filter(m => m.id !== id);
  renderList();
}

// ============ ساخت برنامه ============
document.getElementById("planBtn").onclick = () => {
  if (meetings.length === 0) return alert("اول جلسه اضافه کن.");
  renderSchedule(buildSchedule(meetings));
};

function buildSchedule(list) {
  const byDay = Array.from({length: 6}, () => []);
  list.forEach(m => m.days.forEach(d => byDay[d].push(m)));
  const result = {};
  for (let d = 0; d < 6; d++) {
    let candidates = byDay[d];
    if (candidates.length === 0) continue;
    let chosen = candidates.slice(0, CONFIG.MAX_MEETINGS_PER_DAY);
    chosen = orderByNearestNeighbor(chosen);
    const slots = [];
    let current = CONFIG.START_HOUR * 60;
    const endLimit = CONFIG.END_HOUR * 60;
    for (let i = 0; i < chosen.length; i++) {
      const m = chosen[i];
      const duration = CONFIG.MEETING_MINUTES;
      if (current + duration > endLimit) {
        slots.push({ warning: `جلسه «${m.title}» در این روز جا نشد` });
        continue;
      }
      let gapInfo = null;
      if (i > 0) {
        const prev = chosen[i-1];
        const km = distanceKm(prev, m);
        const travel = travelMinutes(km);
        const gap = Math.min(CONFIG.GAP_MAX, Math.max(CONFIG.GAP_MIN, travel));
        current += gap;
        gapInfo = { km: km.toFixed(1), gap };
      }
      const start = current;
      const end = current + duration;
      slots.push({ meeting: m, start, end, gapInfo });
      current = end + CONFIG.GAP_MIN;
    }
    result[d] = slots;
  }
  return result;
}

function orderByNearestNeighbor(points) {
  if (points.length <= 2) return points.slice();
  const remaining = points.slice();
  const ordered = [remaining.shift()];
  while (remaining.length) {
    const last = ordered[ordered.length - 1];
    let bestIdx = 0, bestDist = Infinity;
    remaining.forEach((p, i) => {
      const d = distanceKm(last, p);
      if (d < bestDist) { bestDist = d; bestIdx = i; }
    });
    ordered.push(remaining.splice(bestIdx, 1)[0]);
  }
  return ordered;
}

function fmt(min) {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
}

function renderSchedule(schedule) {
  const panel = document.getElementById("resultPanel");
  const box = document.getElementById("schedule");
  box.innerHTML = "";
  panel.style.display = "block";
  Object.keys(schedule).forEach(day => {
    const card = document.createElement("div");
    card.className = "day-card";
    card.innerHTML = `<h3>${DAY_NAMES[day]}</h3>`;
    schedule[day].forEach(slot => {
      if (slot.warning) {
        card.innerHTML += `<div class="warn">⚠️ ${slot.warning}</div>`;
        return;
      }
      const m = slot.meeting;
      card.innerHTML += `<div class="slot">
          <div class="time">${fmt(slot.start)} – ${fmt(slot.end)}</div>
          <div class="info"><b>${m.title}</b><small>${m.address}</small>
          ${slot.gapInfo ? `<small>🚗 فاصله: ${slot.gapInfo.km} km — ${slot.gapInfo.gap} دقیقه</small>` : ""}
          </div></div>`;
    });
    box.appendChild(card);
  });
}
