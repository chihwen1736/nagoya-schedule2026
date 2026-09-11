/*
 * 2026 名古屋亞運｜中華台北代表團每日賽程查詢系統
 * 純前端 Vanilla JS。資料來源：data/schedule.js（由 Excel 轉換工具產生，見 README.md）
 */
(function () {
  "use strict";

  // 資料更新日期：以 data/schedule.js 內由轉換工具自動產生的 SCHEDULE_META 為主要來源，
  // 不在 HTML / JS 多處寫死日期。若該檔案尚未包含 metadata（例如舊版資料），才使用下方預設值。
  const CONFIG = {
    DATA_UPDATE_DATE:
      (window.SCHEDULE_META && window.SCHEDULE_META.updateDate) || "2026年9月11日",
  };

  const RAW = Array.isArray(window.SCHEDULE_DATA) ? window.SCHEDULE_DATA : [];

  // ------------------------------------------------------------------
  // 基礎資料準備
  // ------------------------------------------------------------------

  // 依日期彙整資料（沿用轉換工具已排序好的順序：日期 -> 開始時間 -> 運動種類）
  const recordsByDate = {};
  RAW.forEach((r) => {
    if (!recordsByDate[r.date]) recordsByDate[r.date] = [];
    recordsByDate[r.date].push(r);
  });

  // 產生完整日期序列（含 0 筆賽程的日期，例如休兵日），避免日期切換時漏掉某天
  function buildDateRange() {
    const dates = Object.keys(recordsByDate).sort();
    if (dates.length === 0) return [];
    const start = new Date(dates[0] + "T00:00:00+09:00");
    const end = new Date(dates[dates.length - 1] + "T00:00:00+09:00");
    const out = [];
    for (let t = start.getTime(); t <= end.getTime(); t += 86400000) {
      out.push(isoDateJST(new Date(t)));
    }
    return out;
  }

  function isoDateJST(d) {
    const fmt = new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Tokyo",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    });
    return fmt.format(d); // yyyy-mm-dd
  }

  const ALL_DATES = buildDateRange();

  const WEEKDAY_CHARS = ["日", "一", "二", "三", "四", "五", "六"];
  function weekdayLabel(dateStr) {
    // 純日曆運算，避免時區轉換造成的日期偏移（不可用 new Date(dateStr+offset) 後取 getUTCDay）
    const [y, m, d] = dateStr.split("-").map(Number);
    return WEEKDAY_CHARS[new Date(Date.UTC(y, m - 1, d)).getUTCDay()];
  }
  function monthDayLabel(dateStr) {
    const [, m, day] = dateStr.split("-");
    return `${parseInt(m, 10)}/${parseInt(day, 10)}`;
  }

  // 目前日本時間（JST）
  function getJSTNow() {
    const now = new Date();
    const dateStr = isoDateJST(now);
    const parts = new Intl.DateTimeFormat("en-GB", {
      timeZone: "Asia/Tokyo",
      hour: "2-digit",
      minute: "2-digit",
      hourCycle: "h23",
    }).formatToParts(now);
    let hour = 0, minute = 0;
    parts.forEach((p) => {
      if (p.type === "hour") hour = parseInt(p.value, 10);
      if (p.type === "minute") minute = parseInt(p.value, 10);
    });
    return { dateStr, minutes: hour * 60 + minute, timestamp: now.getTime() };
  }

  // 運動項目清單：依資料出現順序去重（不寫死在程式中）
  function buildSportList() {
    const seen = new Set();
    const list = [];
    RAW.forEach((r) => {
      const key = r.normalizedSport || r.sport;
      if (key && !seen.has(key)) {
        seen.add(key);
        list.push(key);
      }
    });
    return list;
  }
  const SPORT_LIST = buildSportList();

  // ------------------------------------------------------------------
  // 狀態
  // ------------------------------------------------------------------

  const jstNowInit = getJSTNow();

  function pickDefaultDate(jstNow) {
    if (ALL_DATES.length === 0) return null;
    if (ALL_DATES.includes(jstNow.dateStr)) return jstNow.dateStr;
    if (jstNow.dateStr < ALL_DATES[0]) return ALL_DATES[0];
    return ALL_DATES[ALL_DATES.length - 1];
  }

  const state = {
    selectedDate: pickDefaultDate(jstNowInit),
    scope: "today", // 'today' | 'all'
    sport: "all",
    search: "",
    view: "all", // 'all' | 'upcoming'
  };

  // ------------------------------------------------------------------
  // DOM refs
  // ------------------------------------------------------------------

  const dateStripEl = document.getElementById("dateStrip");
  const prevDayBtn = document.getElementById("prevDayBtn");
  const nextDayBtn = document.getElementById("nextDayBtn");
  const searchInput = document.getElementById("searchInput");
  const clearSearchBtn = document.getElementById("clearSearchBtn");
  const sportFilterEl = document.getElementById("sportFilter");
  const resultTitleEl = document.getElementById("resultTitle");
  const resultCountEl = document.getElementById("resultCount");
  const scheduleListEl = document.getElementById("scheduleList");
  const emptyStateEl = document.getElementById("emptyState");
  const emptyMessageEl = document.getElementById("emptyMessage");
  const emptyActionBtn = document.getElementById("emptyActionBtn");
  const footerDesktopEl = document.getElementById("footerDesktop");
  const footerMobileDateEl = document.getElementById("footerMobileDate");

  function renderFooter() {
    const d = CONFIG.DATA_UPDATE_DATE;
    footerDesktopEl.textContent =
      `資料更新：${d}｜資料僅供內部作業參考使用，正式資料請以主辦單位公告為準｜建置與維護：國家運動訓練中心競技運動處`;
    footerMobileDateEl.textContent = `資料更新：${d}`;
  }

  // ------------------------------------------------------------------
  // 搜尋比對用（與轉換工具的 clean_search 邏輯一致：移除空白）
  // ------------------------------------------------------------------
  function normalizeQuery(q) {
    return q.replace(/\s+/g, "").replace(/\u3000/g, "").toLowerCase();
  }

  // ------------------------------------------------------------------
  // 渲染：日期切換列
  // ------------------------------------------------------------------
  function renderDateStrip(jstNow) {
    dateStripEl.innerHTML = "";
    ALL_DATES.forEach((dateStr) => {
      const chip = document.createElement("button");
      chip.className = "date-chip";
      chip.dataset.date = dateStr;
      chip.setAttribute("role", "option");
      if (dateStr === jstNow.dateStr) chip.classList.add("is-today");
      if (dateStr === state.selectedDate) {
        chip.classList.add("is-selected");
        chip.setAttribute("aria-selected", "true");
      }
      chip.innerHTML =
        `<span class="d-num">${monthDayLabel(dateStr)}</span>` +
        `<span class="d-wd">${weekdayLabel(dateStr)}</span>`;
      chip.addEventListener("click", () => {
        state.selectedDate = dateStr;
        state.scope = "today";
        render();
        scrollSelectedIntoView();
      });
      dateStripEl.appendChild(chip);
    });
  }

  function scrollSelectedIntoView() {
    const sel = dateStripEl.querySelector(".date-chip.is-selected");
    if (sel) sel.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
  }

  function shiftDate(delta) {
    if (!state.selectedDate) return;
    const idx = ALL_DATES.indexOf(state.selectedDate);
    if (idx === -1) return;
    const next = idx + delta;
    if (next < 0 || next >= ALL_DATES.length) return;
    state.selectedDate = ALL_DATES[next];
    state.scope = "today";
    render();
    scrollSelectedIntoView();
  }
  prevDayBtn.addEventListener("click", () => shiftDate(-1));
  nextDayBtn.addEventListener("click", () => shiftDate(1));

  // ------------------------------------------------------------------
  // 渲染：搜尋範圍 / 項目篩選 / 全部接下來
  // ------------------------------------------------------------------
  document.querySelectorAll(".scope-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.scope = btn.dataset.scope;
      render();
    });
  });

  document.querySelectorAll(".toggle-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.view = btn.dataset.view;
      render();
    });
  });

  function renderSportFilter() {
    sportFilterEl.innerHTML = "";
    const makeChip = (label, value) => {
      const chip = document.createElement("button");
      chip.className = "sport-chip" + (state.sport === value ? " active" : "");
      chip.textContent = label;
      chip.addEventListener("click", () => {
        state.sport = value;
        render();
      });
      return chip;
    };
    sportFilterEl.appendChild(makeChip("全部項目", "all"));
    SPORT_LIST.forEach((sport) => sportFilterEl.appendChild(makeChip(sport, sport)));
  }

  // ------------------------------------------------------------------
  // 搜尋輸入
  // ------------------------------------------------------------------
  searchInput.addEventListener("input", () => {
    state.search = searchInput.value;
    clearSearchBtn.style.display = state.search ? "" : "none";
    render();
  });
  clearSearchBtn.addEventListener("click", () => {
    searchInput.value = "";
    state.search = "";
    clearSearchBtn.style.display = "none";
    render();
  });

  // ------------------------------------------------------------------
  // 篩選核心邏輯
  // ------------------------------------------------------------------
  function getFiltered(jstNow) {
    const q = normalizeQuery(state.search || "");
    const hasSearch = q.length > 0;

    let list = RAW;

    if (state.scope === "today" && state.selectedDate) {
      list = recordsByDate[state.selectedDate] || [];
    }

    if (state.sport !== "all") {
      list = list.filter((r) => (r.normalizedSport || r.sport) === state.sport);
    }

    if (hasSearch) {
      list = list.filter((r) => r._search && r._search.toLowerCase().includes(q));
    }

    if (state.view === "upcoming") {
      list = list.filter((r) => isUpcoming(r, jstNow));
    }

    return list;
  }

  // 判斷該筆賽程是否仍屬於「接下來」：
  // 只要多時段中「任一」開始時間 >= 目前 JST 時間，就視為接下來；
  // 完全無法解析時間的資料一律保留、不隱藏。
  function isUpcoming(r, jstNow) {
    const starts = r.timeStartMinutes || [];
    if (starts.length === 0) return true;
    const dateBase = new Date(r.date + "T00:00:00+09:00").getTime();
    return starts.some((m) => dateBase + m * 60000 >= jstNow.timestamp);
  }

  // ------------------------------------------------------------------
  // 標題與統計文字
  // ------------------------------------------------------------------
  function buildTitleAndCount(list) {
    const q = (state.search || "").trim();
    const parts = [];

    if (state.scope === "today" && state.selectedDate) {
      parts.push(`${monthDayLabel(state.selectedDate)}（${weekdayLabel(state.selectedDate)}）`);
    } else {
      parts.push("全部日期");
    }
    if (state.sport !== "all") parts.push(state.sport);
    if (q) parts.push(`搜尋「${q}」`);

    const title = parts.join("｜");

    let countText;
    if (q) {
      countText = `共 <strong>${list.length}</strong> 筆結果`;
    } else if (state.scope === "today" && state.sport === "all") {
      countText = `今日共 <strong>${list.length}</strong> 筆賽程`;
    } else {
      countText = `共 <strong>${list.length}</strong> 筆賽程`;
    }
    return { title, countText };
  }

  // ------------------------------------------------------------------
  // 賽程卡片渲染
  // ------------------------------------------------------------------
  function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str == null ? "" : String(str);
    return div.innerHTML;
  }

  function renderCard(r, jstNow) {
    // 「已過」判斷：多時段中若仍有任一時段尚未開始，就不算已過（維持與接下來一致的邏輯）
    const isPast = !isUpcoming(r, jstNow) && (r.timeStartMinutes || []).length > 0;

    let html = `<div class="card${isPast ? " is-past" : ""}">`;
    html += `<div class="card-top">`;
    html += `<span class="card-sport">${escapeHtml(r.sport)}</span>`;
    if (r.time) {
      html += `<span class="card-time${r.timeParseable ? "" : " is-unparsed"}">${escapeHtml(r.time)}</span>`;
    }
    html += `</div>`;

    if (r.event) html += `<div class="card-event">${escapeHtml(r.event)}</div>`;
    if (r.athletes) html += `<div class="card-row">${escapeHtml(r.athletes)}</div>`;
    if (r.opponent) html += `<div class="card-row"><span class="label">對手：</span>${escapeHtml(r.opponent)}</div>`;
    if (r.score) html += `<div class="card-result"><span class="label">成績：</span>${escapeHtml(r.score)}</div>`;
    if (r.rank) html += `<div class="card-result"><span class="label">名次：</span>${escapeHtml(r.rank)}</div>`;
    if (r.venue) html += `<div class="card-venue">📍 ${escapeHtml(r.venue)}</div>`;
    if (r.note) html += `<div class="card-note">備註：${escapeHtml(r.note)}</div>`;

    html += `</div>`;
    return html;
  }

  // ------------------------------------------------------------------
  // 空狀態
  // ------------------------------------------------------------------
  function showEmptyState(list, jstNow) {
    const q = (state.search || "").trim();
    const isRawEmptyDay =
      !q &&
      state.sport === "all" &&
      state.view === "all" &&
      state.scope === "today" &&
      state.selectedDate &&
      (recordsByDate[state.selectedDate] || []).length === 0;

    emptyStateEl.style.display = "";
    scheduleListEl.style.display = "none";

    if (isRawEmptyDay) {
      emptyMessageEl.textContent =
        state.selectedDate === jstNow.dateStr ? "今日無中華台北代表隊賽程" : "本日無中華台北代表隊賽程";
      const nextDate = findNextDateWithData(state.selectedDate);
      if (nextDate) {
        emptyActionBtn.style.display = "";
        emptyActionBtn.textContent = "查看下一個賽程日";
        emptyActionBtn.onclick = () => {
          state.selectedDate = nextDate;
          render();
          scrollSelectedIntoView();
        };
      } else {
        emptyActionBtn.style.display = "none";
      }
    } else {
      emptyMessageEl.textContent = "找不到符合條件的賽程";
      emptyActionBtn.style.display = "";
      emptyActionBtn.textContent = "清除搜尋條件";
      emptyActionBtn.onclick = () => {
        searchInput.value = "";
        state.search = "";
        state.sport = "all";
        clearSearchBtn.style.display = "none";
        render();
      };
    }
  }

  function findNextDateWithData(fromDate) {
    const idx = ALL_DATES.indexOf(fromDate);
    if (idx === -1) return null;
    for (let i = idx + 1; i < ALL_DATES.length; i++) {
      if ((recordsByDate[ALL_DATES[i]] || []).length > 0) return ALL_DATES[i];
    }
    return null;
  }

  // ------------------------------------------------------------------
  // 主渲染
  // ------------------------------------------------------------------
  function render() {
    const jstNow = getJSTNow(); // 每次渲染都重新取得目前 JST，避免頁面長時間開著後時間停留在載入當下

    renderDateStrip(jstNow);
    renderSportFilter();

    document.querySelectorAll(".scope-btn").forEach((b) => {
      b.classList.toggle("active", b.dataset.scope === state.scope);
    });
    document.querySelectorAll(".toggle-btn").forEach((b) => {
      b.classList.toggle("active", b.dataset.view === state.view);
    });

    const list = getFiltered(jstNow);
    const { title, countText } = buildTitleAndCount(list);
    resultTitleEl.textContent = title;
    resultCountEl.innerHTML = countText;

    if (list.length === 0) {
      showEmptyState(list, jstNow);
      scheduleListEl.innerHTML = "";
      return;
    }

    emptyStateEl.style.display = "none";
    scheduleListEl.style.display = "";
    scheduleListEl.innerHTML = list.map((r) => renderCard(r, jstNow)).join("");
  }

  // ------------------------------------------------------------------
  // 啟動
  // ------------------------------------------------------------------
  renderFooter();
  render();
  scrollSelectedIntoView();
})();
