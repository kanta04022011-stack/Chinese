(() => {
  const STORAGE_KEY = "zh_telop_flashcards_v2";
  const LEGACY_STORAGE_KEY = "zh_telop_flashcards_v1";

  const GAS_TEMPLATE_CODE = `const SHEET_WORDS = "単語";
const SHEET_SENTENCES = "文章";
const WORD_HEADERS = ["中国語", "ピンイン", "カタカナ", "品詞", "日本語訳", "ステータス", "更新日時"];
const SENTENCE_HEADERS = ["中国語", "日本語訳", "ステータス", "更新日時"];

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#eff6ff");
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#eff6ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function statusToLabel_(status) {
  if (status === "mastered") return "覚えた";
  if (status === "again") return "要復習";
  return "未学習";
}

function labelToStatus_(label) {
  const s = String(label || "").trim();
  if (s === "覚えた" || s === "mastered") return "mastered";
  if (s === "要復習" || s === "もう一度" || s === "again") return "again";
  return "new";
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wordSheet = getOrCreateSheet_(ss, SHEET_WORDS, WORD_HEADERS);
    const sentenceSheet = getOrCreateSheet_(ss, SHEET_SENTENCES, SENTENCE_HEADERS);
    const words = [];
    if (wordSheet.getLastRow() >= 2) {
      const rows = wordSheet.getRange(2, 1, wordSheet.getLastRow() - 1, WORD_HEADERS.length).getValues();
      for (const r of rows) {
        const zh = String(r[0] || "").trim();
        if (!zh) continue;
        words.push({
          zh,
          pinyin: String(r[1] || "").trim(),
          katakana: String(r[2] || "").trim(),
          pos: String(r[3] || "その他").trim() || "その他",
          ja: String(r[4] || "").trim(),
          status: labelToStatus_(r[5]),
          updatedAt: Number(r[6]) || 0
        });
      }
    }
    const sentences = [];
    if (sentenceSheet.getLastRow() >= 2) {
      const rows = sentenceSheet.getRange(2, 1, sentenceSheet.getLastRow() - 1, SENTENCE_HEADERS.length).getValues();
      for (const r of rows) {
        const zh = String(r[0] || "").trim();
        if (!zh) continue;
        sentences.push({
          zh,
          ja: String(r[1] || "").trim(),
          status: labelToStatus_(r[2]),
          updatedAt: Number(r[3]) || 0
        });
      }
    }
    return jsonResponse_({ ok: true, words, sentences });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wordSheet = getOrCreateSheet_(ss, SHEET_WORDS, WORD_HEADERS);
    const sentenceSheet = getOrCreateSheet_(ss, SHEET_SENTENCES, SENTENCE_HEADERS);

    if (payload.action === "updateStatus") {
      const isWord = payload.type === "word";
      const sheet = isWord ? wordSheet : sentenceSheet;
      const statusCol = isWord ? 6 : 3;
      const updatedCol = isWord ? 7 : 4;
      const lastRow = sheet.getLastRow();
      let foundRow = -1;
      if (lastRow >= 2) {
        const zhValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < zhValues.length; i++) {
          if (String(zhValues[i][0] || "").trim() === String(payload.zh || "").trim()) {
            foundRow = i + 2;
            break;
          }
        }
      }
      const label = statusToLabel_(payload.status);
      const updatedAt = Number(payload.updatedAt) || Date.now();
      if (foundRow !== -1) {
        sheet.getRange(foundRow, statusCol).setValue(label);
        sheet.getRange(foundRow, updatedCol).setValue(updatedAt);
      } else if (payload.card) {
        const c = payload.card;
        if (isWord) {
          sheet.appendRow([c.zh || "", c.pinyin || "", c.katakana || "", c.pos || "その他", c.ja || "", label, updatedAt]);
        } else {
          sheet.appendRow([c.zh || "", c.ja || "", label, updatedAt]);
        }
      }
      return jsonResponse_({ ok: true, updated: true });
    }

    if (Array.isArray(payload.words)) {
      if (wordSheet.getLastRow() >= 2) {
        wordSheet.getRange(2, 1, wordSheet.getLastRow() - 1, WORD_HEADERS.length).clearContent();
      }
      if (payload.words.length > 0) {
        const rows = payload.words.map((w) => [
          w.zh || "", w.pinyin || "", w.katakana || "", w.pos || "その他", w.ja || "",
          statusToLabel_(w.status), Number(w.updatedAt) || Date.now()
        ]);
        wordSheet.getRange(2, 1, rows.length, WORD_HEADERS.length).setValues(rows);
      }
    }

    if (Array.isArray(payload.sentences)) {
      if (sentenceSheet.getLastRow() >= 2) {
        sentenceSheet.getRange(2, 1, sentenceSheet.getLastRow() - 1, SENTENCE_HEADERS.length).clearContent();
      }
      if (payload.sentences.length > 0) {
        const rows = payload.sentences.map((s) => [
          s.zh || "", s.ja || "", statusToLabel_(s.status), Number(s.updatedAt) || Date.now()
        ]);
        sentenceSheet.getRange(2, 1, rows.length, SENTENCE_HEADERS.length).setValues(rows);
      }
    }
    return jsonResponse_({ ok: true, saved: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}`;

  const SAMPLE_SENTENCES = {
    zh: [
      "你今天过得怎么样？",
      "我觉得这个主意不错。",
      "不好意思，请问洗手间在哪里？",
      "等一下，我马上就到。",
      "这个多少钱？可以便宜一点吗？",
      "说实话，我也有点儿担心。",
      "没关系，别放在心上。",
      "我们好久不见了，最近忙什么呢？"
    ].join("\n"),
    ja: [
      "今日はどう過ごしましたか？",
      "このアイデアは悪くないと思います。",
      "すみません、お手洗いはどこですか？",
      "ちょっと待って、もうすぐ着きます。",
      "これはいくらですか？少し安くなりますか？",
      "正直なところ、私も少し心配しています。",
      "大丈夫ですよ、気にしないでください。",
      "久しぶりですね、最近は何か忙しいですか？"
    ].join("\n")
  };

  const SAMPLE_WORDS = {
    zh: [
      "觉得", "请问", "担心", "准备",
      "不错", "便宜", "开心", "重要",
      "主意", "洗手间", "时间", "朋友",
      "不好意思", "马上", "最近"
    ].join("\n"),
    pinyin: [
      "juéde", "qǐngwèn", "dānxīn", "zhǔnbèi",
      "bùcuò", "piányi", "kāixīn", "zhòngyào",
      "zhǔyì", "xǐshǒujiān", "shíjiān", "péngyou",
      "bù hǎoyìsi", "mǎshàng", "zuìjìn"
    ].join("\n"),
    katakana: [
      "ジュエダ", "チンウェン", "ダンシン", "ジュンベイ",
      "ブーツオ", "ピエンイー", "カイシン", "ジョンヤオ",
      "ジューイー", "シーショウジエン", "シージエン", "ポンヨウ",
      "ブ ハオイース", "マーシャン", "ズイジン"
    ].join("\n"),
    pos: [
      "動詞", "動詞", "動詞", "動詞",
      "形容詞", "形容詞", "形容詞", "形容詞",
      "名詞", "名詞", "名詞", "名詞",
      "その他", "その他", "名詞"
    ].join("\n"),
    ja: [
      "〜だと思う・感じる", "すみませんが（お尋ねします）", "心配する", "準備する",
      "悪くない・なかなか良い", "安い", "嬉しい・楽しい", "重要だ",
      "アイデア・考え", "お手洗い・トイレ", "時間", "友達",
      "すみません・申し訳ない", "すぐに・もうすぐ", "最近"
    ].join("\n")
  };

  // App State
  let activeMode = "word"; // "word" | "sentence"
  let activeTab = "study"; // "study" | "input" | "sync"
  let activePosFilter = "all";
  let gasUrl = "";
  let autoSync = true;
  let alwaysShowReading = false;

  const decks = {
    word: {
      cards: [], // { id, zh, pinyin, katakana, pos, ja, status, updatedAt }
      deckIndices: [],
      currentPos: 0,
      isRevealed: false,
      isReadingHintShown: false,
      unmasteredOnly: false
    },
    sentence: {
      cards: [], // { id, zh, ja, status, updatedAt }
      deckIndices: [],
      currentPos: 0,
      isRevealed: false,
      isReadingHintShown: false,
      unmasteredOnly: false
    }
  };

  // DOM Elements - Header
  const modeWordBtn = document.getElementById("mode-word");
  const modeSentenceBtn = document.getElementById("mode-sentence");
  const badgeWordCount = document.getElementById("badge-word-count");
  const badgeSentenceCount = document.getElementById("badge-sentence-count");
  const tabStudy = document.getElementById("tab-study");
  const tabInput = document.getElementById("tab-input");
  const tabSync = document.getElementById("tab-sync");
  const syncDot = document.getElementById("sync-dot");
  const viewStudy = document.getElementById("view-study");
  const viewInput = document.getElementById("view-input");
  const viewSync = document.getElementById("view-sync");

  // DOM Elements - Study
  const emptyState = document.getElementById("empty-state");
  const emptyTitle = document.getElementById("empty-title");
  const emptyDesc = document.getElementById("empty-desc");
  const studyWorkspace = document.getElementById("study-workspace");
  const posFilterBar = document.getElementById("pos-filter-bar");
  const posFilterButtons = document.querySelectorAll(".pos-filter-btn");
  const posCountAll = document.getElementById("pos-count-all");
  const posCountVerb = document.getElementById("pos-count-verb");
  const posCountAdj = document.getElementById("pos-count-adj");
  const posCountNoun = document.getElementById("pos-count-noun");
  const posCountOther = document.getElementById("pos-count-other");

  const cardCounter = document.getElementById("card-counter");
  const masteredCounter = document.getElementById("mastered-counter");
  const studySyncBadge = document.getElementById("study-sync-badge");
  const btnQuickSync = document.getElementById("btn-quick-sync");
  const progressBarFill = document.getElementById("progress-bar-fill");
  const chkAlwaysShowReading = document.getElementById("chk-always-show-reading");
  const chkUnmasteredOnly = document.getElementById("chk-unmastered-only");
  const btnShuffle = document.getElementById("btn-shuffle");
  const btnResetOrder = document.getElementById("btn-reset-order");

  const cardModePill = document.getElementById("card-mode-pill");
  const cardPosPill = document.getElementById("card-pos-pill");
  const cardStatusPill = document.getElementById("card-status-pill");
  const btnSpeak = document.getElementById("btn-speak");
  const zhLabel = document.getElementById("zh-label");
  const cardZh = document.getElementById("card-zh");
  const readingSection = document.getElementById("reading-section");
  const cardPinyin = document.getElementById("card-pinyin");
  const cardKatakana = document.getElementById("card-katakana");
  const jaSection = document.getElementById("ja-section");
  const cardJa = document.getElementById("card-ja");
  const btnHintReading = document.getElementById("btn-hint-reading");
  const btnReveal = document.getElementById("btn-reveal");

  const btnPrev = document.getElementById("btn-prev");
  const btnNext = document.getElementById("btn-next");
  const btnAgain = document.getElementById("btn-again");
  const btnMastered = document.getElementById("btn-mastered");
  const btnGoInput = document.getElementById("btn-go-input");
  const btnLoadSampleEmpty = document.getElementById("btn-load-sample-empty");

  // DOM Elements - Word Input
  const panelWordInput = document.getElementById("panel-word-input");
  const inputWordZh = document.getElementById("input-word-zh");
  const inputWordPinyin = document.getElementById("input-word-pinyin");
  const inputWordKatakana = document.getElementById("input-word-katakana");
  const inputWordPos = document.getElementById("input-word-pos");
  const inputWordJa = document.getElementById("input-word-ja");
  const countWordZh = document.getElementById("count-word-zh");
  const countWordPinyin = document.getElementById("count-word-pinyin");
  const countWordKatakana = document.getElementById("count-word-katakana");
  const countWordPos = document.getElementById("count-word-pos");
  const countWordJa = document.getElementById("count-word-ja");
  const chkWordIgnoreEmpty = document.getElementById("chk-word-ignore-empty");
  const wordLineMatchStatus = document.getElementById("word-line-match-status");
  const wordMismatchWarning = document.getElementById("word-mismatch-warning");
  const wordSavedNotice = document.getElementById("word-saved-notice");
  const btnExtractFromSentences = document.getElementById("btn-extract-from-sentences");
  const btnAutoFillReading = document.getElementById("btn-auto-fill-reading");
  const btnLoadWordSample = document.getElementById("btn-load-word-sample");
  const btnClearWordInputs = document.getElementById("btn-clear-word-inputs");
  const btnCreateWordCards = document.getElementById("btn-create-word-cards");

  // DOM Elements - Sentence Input
  const panelSentenceInput = document.getElementById("panel-sentence-input");
  const inputSentenceZh = document.getElementById("input-sentence-zh");
  const inputSentenceJa = document.getElementById("input-sentence-ja");
  const countSentenceZh = document.getElementById("count-sentence-zh");
  const countSentenceJa = document.getElementById("count-sentence-ja");
  const chkSentenceIgnoreEmpty = document.getElementById("chk-sentence-ignore-empty");
  const sentenceLineMatchStatus = document.getElementById("sentence-line-match-status");
  const sentenceMismatchWarning = document.getElementById("sentence-mismatch-warning");
  const sentenceSavedNotice = document.getElementById("sentence-saved-notice");
  const btnLoadSentenceSample = document.getElementById("btn-load-sentence-sample");
  const btnClearSentenceInputs = document.getElementById("btn-clear-sentence-inputs");
  const btnSendToWordExtractor = document.getElementById("btn-send-to-word-extractor");
  const btnCreateSentenceCards = document.getElementById("btn-create-sentence-cards");

  // DOM Elements - Sync
  const syncStatusPill = document.getElementById("sync-status-pill");
  const inputGasUrl = document.getElementById("input-gas-url");
  const btnSaveGasUrl = document.getElementById("btn-save-gas-url");
  const chkAutoSync = document.getElementById("chk-auto-sync");
  const btnSyncBidirectional = document.getElementById("btn-sync-bidirectional");
  const btnSyncPull = document.getElementById("btn-sync-pull");
  const btnSyncPush = document.getElementById("btn-sync-push");
  const syncLog = document.getElementById("sync-log");
  const btnCopyShareUrl = document.getElementById("btn-copy-share-url");
  const shareUrlNotice = document.getElementById("share-url-notice");
  const btnCopyGasCode = document.getElementById("btn-copy-gas-code");
  const gasCopyNotice = document.getElementById("gas-copy-notice");

  // --- Initialization ---
  function init() {
    loadFromStorage();
    const urlSyncParam = new URLSearchParams(window.location.search).get("sync");
    if (urlSyncParam && urlSyncParam.startsWith("https://")) {
      gasUrl = urlSyncParam.trim();
      inputGasUrl.value = gasUrl;
      saveToStorage();
      // URLから ?sync=... を消してスッキリさせる
      window.history.replaceState({}, "", window.location.pathname);
    }

    bindEvents();
    updateWordInputStats();
    updateSentenceInputStats();
    rebuildDeck("word", false);
    rebuildDeck("sentence", false);
    updateSyncUI();
    updateModeUI();

    if (gasUrl) {
      // 設定済みなら起動時にバックグラウンドで双方向同期
      syncBidirectional(true);
    }

    if (decks[activeMode].cards.length === 0) {
      if (decks.sentence.cards.length > 0) {
        activeMode = "sentence";
        updateModeUI();
        switchTab("study");
      } else if (!gasUrl) {
        switchTab("input");
      } else {
        switchTab("study");
      }
    } else {
      switchTab("study");
    }
  }

  function loadFromStorage() {
    try {
      const rawV2 = localStorage.getItem(STORAGE_KEY);
      if (rawV2) {
        const data = JSON.parse(rawV2);
        if (data.activeMode === "word" || data.activeMode === "sentence") activeMode = data.activeMode;
        if (typeof data.alwaysShowReading === "boolean") {
          alwaysShowReading = data.alwaysShowReading;
          chkAlwaysShowReading.checked = alwaysShowReading;
        }
        if (typeof data.gasUrl === "string") {
          gasUrl = data.gasUrl;
          inputGasUrl.value = gasUrl;
        }
        if (typeof data.autoSync === "boolean") {
          autoSync = data.autoSync;
          chkAutoSync.checked = autoSync;
        }
        if (data.wordInputs) {
          inputWordZh.value = data.wordInputs.zh || "";
          inputWordPinyin.value = data.wordInputs.pinyin || "";
          inputWordKatakana.value = data.wordInputs.katakana || "";
          inputWordPos.value = data.wordInputs.pos || "";
          inputWordJa.value = data.wordInputs.ja || "";
          if (typeof data.wordInputs.ignoreEmpty === "boolean") chkWordIgnoreEmpty.checked = data.wordInputs.ignoreEmpty;
        }
        if (data.sentenceInputs) {
          inputSentenceZh.value = data.sentenceInputs.zh || "";
          inputSentenceJa.value = data.sentenceInputs.ja || "";
          if (typeof data.sentenceInputs.ignoreEmpty === "boolean") chkSentenceIgnoreEmpty.checked = data.sentenceInputs.ignoreEmpty;
        }
        if (Array.isArray(data.wordCards)) {
          decks.word.cards = data.wordCards.map((c) => ({
            ...c,
            pos: window.ZhDict ? window.ZhDict.normalizePos(c.pos) : c.pos || "その他",
            updatedAt: Number(c.updatedAt) || 0
          }));
        }
        if (Array.isArray(data.sentenceCards)) {
          decks.sentence.cards = data.sentenceCards.map((c) => ({
            ...c,
            updatedAt: Number(c.updatedAt) || 0
          }));
        }
        return;
      }

      const rawV1 = localStorage.getItem(LEGACY_STORAGE_KEY);
      if (rawV1) {
        const v1 = JSON.parse(rawV1);
        if (typeof v1.rawZh === "string") inputSentenceZh.value = v1.rawZh;
        if (typeof v1.rawJa === "string") inputSentenceJa.value = v1.rawJa;
        if (typeof v1.ignoreEmpty === "boolean") chkSentenceIgnoreEmpty.checked = v1.ignoreEmpty;
        if (Array.isArray(v1.cards)) decks.sentence.cards = v1.cards;
      }
    } catch (err) {
      console.warn("保存データの読み込みに失敗しました:", err);
    }
  }

  function saveToStorage() {
    try {
      const payload = {
        activeMode,
        alwaysShowReading,
        gasUrl,
        autoSync,
        wordInputs: {
          zh: inputWordZh.value,
          pinyin: inputWordPinyin.value,
          katakana: inputWordKatakana.value,
          pos: inputWordPos.value,
          ja: inputWordJa.value,
          ignoreEmpty: chkWordIgnoreEmpty.checked
        },
        sentenceInputs: {
          zh: inputSentenceZh.value,
          ja: inputSentenceJa.value,
          ignoreEmpty: chkSentenceIgnoreEmpty.checked
        },
        wordCards: decks.word.cards,
        sentenceCards: decks.sentence.cards
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (err) {
      console.warn("データの保存に失敗しました:", err);
    }
  }

  // --- Mode & Tab Switching ---
  function switchMode(mode) {
    activeMode = mode;
    saveToStorage();
    updateModeUI();
  }

  function updateModeUI() {
    const isWord = activeMode === "word";
    modeWordBtn.classList.toggle("active", isWord);
    modeSentenceBtn.classList.toggle("active", !isWord);

    panelWordInput.classList.toggle("hidden", !isWord);
    panelSentenceInput.classList.toggle("hidden", isWord);

    badgeWordCount.textContent = String(decks.word.cards.length);
    badgeSentenceCount.textContent = String(decks.sentence.cards.length);
    chkUnmasteredOnly.checked = decks[activeMode].unmasteredOnly;

    if (activeTab === "study") {
      renderStudyView();
    }
  }

  function switchTab(tab) {
    activeTab = tab;
    tabStudy.classList.toggle("active", tab === "study");
    tabInput.classList.toggle("active", tab === "input");
    tabSync.classList.toggle("active", tab === "sync");

    tabStudy.setAttribute("aria-selected", String(tab === "study"));
    tabInput.setAttribute("aria-selected", String(tab === "input"));
    tabSync.setAttribute("aria-selected", String(tab === "sync"));

    viewStudy.classList.toggle("active", tab === "study");
    viewInput.classList.toggle("active", tab === "input");
    viewSync.classList.toggle("active", tab === "sync");

    if (tab === "study") {
      renderStudyView();
    }
  }

  // --- Parsing & Input Helpers ---
  function parseLines(text, ignoreEmpty) {
    const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const lines = normalized.split("\n").map((line) => line.trim());
    return ignoreEmpty ? lines.filter((line) => line.length > 0) : lines;
  }

  function showNotice(el, message) {
    el.textContent = message;
    setTimeout(() => {
      if (el.textContent === message) el.textContent = "";
    }, 3000);
  }

  function syncTextareasFromCards() {
    inputWordZh.value = decks.word.cards.map((c) => c.zh).join("\n");
    inputWordPinyin.value = decks.word.cards.map((c) => c.pinyin || "").join("\n");
    inputWordKatakana.value = decks.word.cards.map((c) => c.katakana || "").join("\n");
    inputWordPos.value = decks.word.cards.map((c) => c.pos || "その他").join("\n");
    inputWordJa.value = decks.word.cards.map((c) => c.ja || "").join("\n");

    inputSentenceZh.value = decks.sentence.cards.map((c) => c.zh).join("\n");
    inputSentenceJa.value = decks.sentence.cards.map((c) => c.ja || "").join("\n");

    updateWordInputStats();
    updateSentenceInputStats();
  }

  // --- Word Input Logic ---
  function updateWordInputStats() {
    const ignore = chkWordIgnoreEmpty.checked;
    const zh = parseLines(inputWordZh.value, ignore);
    const pinyin = parseLines(inputWordPinyin.value, ignore);
    const kata = parseLines(inputWordKatakana.value, ignore);
    const pos = parseLines(inputWordPos.value, ignore);
    const ja = parseLines(inputWordJa.value, ignore);

    countWordZh.textContent = `${zh.length}行`;
    countWordPinyin.textContent = `${pinyin.length}行`;
    countWordKatakana.textContent = `${kata.length}行`;
    countWordPos.textContent = `${pos.length}行`;
    countWordJa.textContent = `${ja.length}行`;

    wordLineMatchStatus.textContent = `単語: ${zh.length}行 ／ 発音: ${pinyin.length}行 ／ カタカナ: ${kata.length}行 ／ 品詞: ${pos.length}行 ／ 訳: ${ja.length}行`;
    wordLineMatchStatus.classList.remove("matched", "mismatched");

    if (zh.length > 0 && zh.length === ja.length) {
      wordLineMatchStatus.classList.add("matched");
      wordMismatchWarning.classList.add("hidden");
    } else if (zh.length !== ja.length) {
      wordLineMatchStatus.classList.add("mismatched");
      wordMismatchWarning.classList.remove("hidden");
    } else {
      wordMismatchWarning.classList.add("hidden");
    }
  }

  function autoFillWordReadings() {
    const zhLines = parseLines(inputWordZh.value, false);
    const pinyinLines = parseLines(inputWordPinyin.value, false);
    const kataLines = parseLines(inputWordKatakana.value, false);
    const posLines = parseLines(inputWordPos.value, false);
    const jaLines = parseLines(inputWordJa.value, false);

    const maxLen = Math.max(zhLines.length, pinyinLines.length, kataLines.length, posLines.length, jaLines.length);
    let filledCount = 0;

    for (let i = 0; i < maxLen; i++) {
      const w = zhLines[i] || "";
      let p = pinyinLines[i] || "";
      let k = kataLines[i] || "";
      let ps = posLines[i] || "";
      let j = jaLines[i] || "";

      if (w && window.ZhDict) {
        const entry = window.ZhDict.lookupWord(w);
        if (entry) {
          if (!p && entry.pinyin) { p = entry.pinyin; filledCount++; }
          if (!k && entry.katakana) { k = entry.katakana; filledCount++; }
          if (!ps && entry.pos) { ps = entry.pos; filledCount++; }
          if (!j && entry.ja) { j = entry.ja; filledCount++; }
        }
      }
      if (p && !k && window.ZhDict) {
        k = window.ZhDict.pinyinToKatakana(p);
        if (k) filledCount++;
      }

      pinyinLines[i] = p;
      kataLines[i] = k;
      posLines[i] = ps;
      jaLines[i] = j;
    }

    inputWordPinyin.value = pinyinLines.join("\n");
    inputWordKatakana.value = kataLines.join("\n");
    inputWordPos.value = posLines.join("\n");
    inputWordJa.value = jaLines.join("\n");
    updateWordInputStats();
    saveToStorage();
    showNotice(wordSavedNotice, filledCount > 0 ? `${filledCount}箇所を自動補完しました` : "補完できる項目がありませんでした");
  }

  function extractWordsFromSentenceInput() {
    const sentences = parseLines(inputSentenceZh.value, true);
    if (sentences.length === 0) {
      showNotice(wordSavedNotice, "文章モードに中国語テロップが入力されていません");
      return;
    }
    const extracted = window.ZhDict.extractWordsFromSentences(sentences);
    const posOrder = { "動詞": 1, "形容詞": 2, "名詞": 3, "その他": 4 };
    extracted.sort((a, b) => (posOrder[a.pos] || 9) - (posOrder[b.pos] || 9));

    inputWordZh.value = extracted.map((e) => e.word).join("\n");
    inputWordPinyin.value = extracted.map((e) => e.pinyin || "").join("\n");
    inputWordKatakana.value = extracted.map((e) => e.katakana || "").join("\n");
    inputWordPos.value = extracted.map((e) => e.pos || "その他").join("\n");
    inputWordJa.value = extracted.map((e) => e.ja || "").join("\n");

    chkWordIgnoreEmpty.checked = false;
    updateWordInputStats();
    saveToStorage();
    switchMode("word");
    switchTab("input");
    showNotice(wordSavedNotice, `${extracted.length}個の単語を品詞分けして抽出しました`);
  }

  function buildWordCards() {
    const ignore = chkWordIgnoreEmpty.checked;
    const zhLines = parseLines(inputWordZh.value, ignore);
    const pinyinLines = parseLines(inputWordPinyin.value, ignore);
    const kataLines = parseLines(inputWordKatakana.value, ignore);
    const posLines = parseLines(inputWordPos.value, ignore);
    const jaLines = parseLines(inputWordJa.value, ignore);

    const maxLen = Math.max(zhLines.length, jaLines.length);
    const prevMap = new Map();
    for (const c of decks.word.cards) {
      prevMap.set(c.zh, c);
    }

    const now = Date.now();
    const nextCards = [];
    for (let i = 0; i < maxLen; i++) {
      const zh = zhLines[i] || "";
      const ja = jaLines[i] || "";
      if (!zh && !ja) continue;

      let pinyin = pinyinLines[i] || "";
      let katakana = kataLines[i] || "";
      let rawPos = posLines[i] || "";

      if (zh && window.ZhDict) {
        const entry = window.ZhDict.lookupWord(zh);
        if (entry) {
          if (!pinyin) pinyin = entry.pinyin;
          if (!katakana) katakana = entry.katakana;
          if (!rawPos) rawPos = entry.pos;
        }
      }
      if (!katakana && pinyin && window.ZhDict) {
        katakana = window.ZhDict.pinyinToKatakana(pinyin);
      }

      const pos = window.ZhDict ? window.ZhDict.normalizePos(rawPos) : rawPos || "その他";
      const prev = prevMap.get(zh);
      nextCards.push({
        id: i,
        zh: zh || "（中国語未入力）",
        pinyin,
        katakana,
        pos,
        ja: ja || "（日本語訳未入力）",
        status: prev ? prev.status : "new",
        updatedAt: prev ? prev.updatedAt || now : now
      });
    }

    decks.word.cards = nextCards;
    activePosFilter = "all";
    updatePosFilterButtonsUI();
    saveToStorage();
    rebuildDeck("word", false);
    switchMode("word");
    switchTab("study");

    if (gasUrl && autoSync) {
      pushToSheet(true);
    }
  }

  // --- Sentence Input Logic ---
  function updateSentenceInputStats() {
    const ignore = chkSentenceIgnoreEmpty.checked;
    const zh = parseLines(inputSentenceZh.value, ignore);
    const ja = parseLines(inputSentenceJa.value, ignore);

    countSentenceZh.textContent = `${zh.length}行`;
    countSentenceJa.textContent = `${ja.length}行`;
    sentenceLineMatchStatus.textContent = `中国語: ${zh.length}行 ／ 日本語: ${ja.length}行`;

    sentenceLineMatchStatus.classList.remove("matched", "mismatched");
    if (zh.length > 0 && zh.length === ja.length) {
      sentenceLineMatchStatus.classList.add("matched");
      sentenceMismatchWarning.classList.add("hidden");
    } else if (zh.length !== ja.length) {
      sentenceLineMatchStatus.classList.add("mismatched");
      sentenceMismatchWarning.classList.remove("hidden");
    } else {
      sentenceMismatchWarning.classList.add("hidden");
    }
  }

  function buildSentenceCards() {
    const ignore = chkSentenceIgnoreEmpty.checked;
    const zhLines = parseLines(inputSentenceZh.value, ignore);
    const jaLines = parseLines(inputSentenceJa.value, ignore);
    const maxLen = Math.max(zhLines.length, jaLines.length);

    const prevMap = new Map();
    for (const c of decks.sentence.cards) {
      prevMap.set(c.zh, c);
    }

    const now = Date.now();
    const nextCards = [];
    for (let i = 0; i < maxLen; i++) {
      const zh = zhLines[i] || "";
      const ja = jaLines[i] || "";
      if (!zh && !ja) continue;
      const prev = prevMap.get(zh);
      nextCards.push({
        id: i,
        zh: zh || "（中国語未入力）",
        ja: ja || "（日本語訳未入力）",
        status: prev ? prev.status : "new",
        updatedAt: prev ? prev.updatedAt || now : now
      });
    }

    decks.sentence.cards = nextCards;
    saveToStorage();
    rebuildDeck("sentence", false);
    switchMode("sentence");
    switchTab("study");

    if (gasUrl && autoSync) {
      pushToSheet(true);
    }
  }

  // --- POS Filter & Deck Building ---
  function updatePosCounts() {
    const cards = decks.word.cards;
    let verb = 0, adj = 0, noun = 0, other = 0;
    for (const c of cards) {
      const p = c.pos || "その他";
      if (p === "動詞") verb++;
      else if (p === "形容詞") adj++;
      else if (p === "名詞") noun++;
      else other++;
    }
    posCountAll.textContent = String(cards.length);
    posCountVerb.textContent = String(verb);
    posCountAdj.textContent = String(adj);
    posCountNoun.textContent = String(noun);
    posCountOther.textContent = String(other);
  }

  function updatePosFilterButtonsUI() {
    posFilterButtons.forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.pos === activePosFilter);
    });
  }

  function rebuildDeck(mode, shuffle = false) {
    const deck = decks[mode];
    deck.deckIndices = deck.cards
      .map((card, idx) => ({ card, idx }))
      .filter(({ card }) => {
        if (mode === "word" && activePosFilter !== "all") {
          if ((card.pos || "その他") !== activePosFilter) return false;
        }
        if (deck.unmasteredOnly && card.status === "mastered") return false;
        return true;
      })
      .map(({ idx }) => idx);

    if (shuffle) {
      for (let i = deck.deckIndices.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [deck.deckIndices[i], deck.deckIndices[j]] = [deck.deckIndices[j], deck.deckIndices[i]];
      }
    }

    deck.currentPos = 0;
    deck.isRevealed = false;
    deck.isReadingHintShown = false;

    badgeWordCount.textContent = String(decks.word.cards.length);
    badgeSentenceCount.textContent = String(decks.sentence.cards.length);
    if (mode === "word") updatePosCounts();

    if (mode === activeMode && activeTab === "study") {
      renderStudyView();
    }
  }

  function getSentenceReading(sentence) {
    if (!window.ZhDict) return { pinyin: "", katakana: "" };
    const words = window.ZhDict.extractWordsFromSentences([sentence]);
    const known = words.filter((w) => w.pinyin);
    if (known.length === 0) return { pinyin: "", katakana: "" };
    return {
      pinyin: known.map((w) => `${w.word}(${w.pinyin})`).join("  "),
      katakana: known.map((w) => `${w.word}: ${w.katakana}`).join(" ／ ")
    };
  }

  function getPosPillClass(pos) {
    if (pos === "動詞") return "pos-pill pos-verb";
    if (pos === "形容詞") return "pos-pill pos-adj";
    if (pos === "名詞") return "pos-pill pos-noun";
    return "pos-pill pos-other";
  }

  function renderStudyView() {
    const deck = decks[activeMode];
    const isWord = activeMode === "word";

    badgeWordCount.textContent = String(decks.word.cards.length);
    badgeSentenceCount.textContent = String(decks.sentence.cards.length);
    posFilterBar.classList.toggle("hidden", !isWord);
    if (isWord) {
      updatePosCounts();
      updatePosFilterButtonsUI();
    }

    if (deck.cards.length === 0) {
      emptyState.classList.remove("hidden");
      studyWorkspace.classList.add("hidden");
      emptyTitle.textContent = isWord ? "まだ単語カードがありません" : "まだ文章カードがありません";
      emptyDesc.textContent = isWord
        ? "「入力・編集」タブから入力するか、「シート同期」からスプレッドシートの単語を読み込んでください。"
        : "「入力・編集」タブからテロップを入力するか、「シート同期」から読み込んでください。";
      return;
    }

    emptyState.classList.add("hidden");
    studyWorkspace.classList.remove("hidden");

    const masteredTotal = deck.cards.filter((c) => c.status === "mastered").length;
    masteredCounter.textContent = `覚えた: ${masteredTotal} / ${deck.cards.length}`;

    cardModePill.textContent = isWord ? "単語モード" : "文章モード";
    zhLabel.textContent = isWord ? "中国語（単語）" : "中国語（テロップ）";
    cardZh.classList.toggle("is-sentence", !isWord);

    if (deck.deckIndices.length === 0) {
      cardCounter.textContent = "0 / 0";
      progressBarFill.style.width = "100%";
      cardPosPill.classList.add("hidden");
      readingSection.classList.add("hidden");
      jaSection.classList.remove("hidden");
      btnHintReading.classList.add("hidden");
      btnReveal.classList.add("hidden");
      btnPrev.disabled = true;
      btnNext.disabled = true;
      btnAgain.disabled = true;
      btnMastered.disabled = true;
      btnSpeak.disabled = true;

      if (isWord && activePosFilter !== "all" && !deck.unmasteredOnly) {
        cardZh.textContent = `「${activePosFilter}」のカードがありません`;
        cardJa.textContent = "上の品詞フィルターから別の品詞を選ぶか、「入力・編集」から追加してください。";
        cardStatusPill.textContent = "0件";
        cardStatusPill.className = "status-pill";
      } else {
        cardZh.textContent = "すべて覚えました！🎉";
        cardJa.textContent = "「未習得のみ」のチェックを外すか、別の品詞を選んで学習を続けられます。";
        cardStatusPill.textContent = "完了";
        cardStatusPill.className = "status-pill mastered";
      }
      return;
    }

    btnHintReading.classList.remove("hidden");
    btnReveal.classList.remove("hidden");
    btnAgain.disabled = false;
    btnMastered.disabled = false;
    btnSpeak.disabled = false;

    if (deck.currentPos >= deck.deckIndices.length) deck.currentPos = deck.deckIndices.length - 1;
    if (deck.currentPos < 0) deck.currentPos = 0;

    const card = deck.cards[deck.deckIndices[deck.currentPos]];
    cardCounter.textContent = `${deck.currentPos + 1} / ${deck.deckIndices.length}`;
    progressBarFill.style.width = `${((deck.currentPos + 1) / deck.deckIndices.length) * 100}%`;

    cardZh.textContent = card.zh;
    cardJa.textContent = card.ja;

    if (isWord) {
      const posLabel = card.pos || "その他";
      cardPosPill.textContent = posLabel;
      cardPosPill.className = getPosPillClass(posLabel);
    } else {
      cardPosPill.classList.add("hidden");
    }

    let pinyinStr = card.pinyin || "";
    let kataStr = card.katakana || "";
    if (!isWord) {
      const aux = getSentenceReading(card.zh);
      pinyinStr = aux.pinyin;
      kataStr = aux.katakana;
    } else if (!kataStr && pinyinStr && window.ZhDict) {
      kataStr = window.ZhDict.pinyinToKatakana(pinyinStr);
    }

    const hasReading = Boolean(pinyinStr || kataStr);
    const shouldShowReading = hasReading && (alwaysShowReading || deck.isReadingHintShown || deck.isRevealed);

    readingSection.classList.toggle("hidden", !shouldShowReading);
    cardPinyin.textContent = pinyinStr;
    cardPinyin.classList.toggle("hidden", !pinyinStr);
    cardKatakana.textContent = kataStr;
    cardKatakana.classList.toggle("hidden", !kataStr);

    btnHintReading.classList.toggle("hidden", !hasReading || alwaysShowReading || deck.isRevealed);
    btnHintReading.innerHTML = deck.isReadingHintShown
      ? '発音・カタカナを隠す <span class="shortcut-hint">(H)</span>'
      : '発音・カタカナを見る <span class="shortcut-hint">(H)</span>';

    jaSection.classList.toggle("hidden", !deck.isRevealed);
    btnReveal.innerHTML = deck.isRevealed
      ? '日本語訳を隠す <span class="shortcut-hint">(Space)</span>'
      : '日本語訳を見る <span class="shortcut-hint">(Space)</span>';

    cardStatusPill.className = "status-pill";
    if (card.status === "mastered") {
      cardStatusPill.textContent = "覚えた";
      cardStatusPill.classList.add("mastered");
    } else if (card.status === "again") {
      cardStatusPill.textContent = "要復習";
      cardStatusPill.classList.add("again");
    } else {
      cardStatusPill.textContent = "未学習";
    }

    btnPrev.disabled = deck.currentPos === 0;
    btnNext.disabled = deck.currentPos >= deck.deckIndices.length - 1;
  }

  function toggleReveal() {
    const deck = decks[activeMode];
    if (deck.deckIndices.length === 0) return;
    deck.isRevealed = !deck.isRevealed;
    renderStudyView();
  }

  function toggleReadingHint() {
    const deck = decks[activeMode];
    if (deck.deckIndices.length === 0) return;
    deck.isReadingHintShown = !deck.isReadingHintShown;
    renderStudyView();
  }

  function goPrev() {
    const deck = decks[activeMode];
    if (deck.currentPos > 0) {
      deck.currentPos--;
      deck.isRevealed = false;
      deck.isReadingHintShown = false;
      renderStudyView();
    }
  }

  function goNext() {
    const deck = decks[activeMode];
    if (deck.currentPos < deck.deckIndices.length - 1) {
      deck.currentPos++;
      deck.isRevealed = false;
      deck.isReadingHintShown = false;
      renderStudyView();
    }
  }

  function markStatus(status) {
    const deck = decks[activeMode];
    if (deck.deckIndices.length === 0) return;
    const cardIdx = deck.deckIndices[deck.currentPos];
    const card = deck.cards[cardIdx];
    const now = Date.now();
    card.status = status;
    card.updatedAt = now;
    saveToStorage();

    if (gasUrl && autoSync) {
      sendStatusUpdateToSheet(activeMode, card);
    }

    if (deck.unmasteredOnly && status === "mastered") {
      deck.deckIndices.splice(deck.currentPos, 1);
      deck.isRevealed = false;
      deck.isReadingHintShown = false;
      if (deck.currentPos >= deck.deckIndices.length && deck.deckIndices.length > 0) {
        deck.currentPos = deck.deckIndices.length - 1;
      }
      renderStudyView();
      return;
    }

    if (deck.currentPos < deck.deckIndices.length - 1) {
      deck.currentPos++;
      deck.isRevealed = false;
      deck.isReadingHintShown = false;
    }
    renderStudyView();
  }

  function speakCurrentChinese() {
    const deck = decks[activeMode];
    if (deck.deckIndices.length === 0) return;
    if (!("speechSynthesis" in window)) {
      alert("お使いのブラウザは音声読み上げに対応していません。");
      return;
    }
    const card = deck.cards[deck.deckIndices[deck.currentPos]];
    window.speechSynthesis.cancel();
    const utterance = new SpeechSynthesisUtterance(card.zh);
    utterance.lang = "zh-CN";
    utterance.rate = activeMode === "word" ? 0.85 : 0.92;
    window.speechSynthesis.speak(utterance);
  }

  // --- Google Sheets Sync Logic ---
  function updateSyncUI(badgeText = "", isError = false) {
    const configured = Boolean(gasUrl);
    syncDot.classList.toggle("hidden", !configured);
    btnQuickSync.classList.toggle("hidden", !configured);
    syncStatusPill.textContent = configured ? "接続設定済み" : "未設定";
    syncStatusPill.className = configured ? "sync-status-pill connected" : "sync-status-pill";
    if (isError) syncStatusPill.className = "sync-status-pill error";

    if (configured && badgeText) {
      studySyncBadge.textContent = badgeText;
      studySyncBadge.classList.remove("hidden", "error");
      if (isError) studySyncBadge.classList.add("error");
    } else if (configured) {
      studySyncBadge.textContent = "☁️ 同期ON";
      studySyncBadge.classList.remove("hidden", "error");
    } else {
      studySyncBadge.classList.add("hidden");
    }
  }

  function setSyncLog(message, type = "info") {
    syncLog.textContent = message;
    syncLog.className = `sync-log ${type === "success" ? "success" : type === "error" ? "error" : ""}`;
    syncLog.classList.remove("hidden");
  }

  async function gasGet() {
    const res = await fetch(gasUrl, { method: "GET", redirect: "follow" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "取得に失敗しました");
    return data;
  }

  async function gasPost(payload) {
    // CORSプリフライト回避のため text/plain を使用
    const res = await fetch(gasUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
      redirect: "follow"
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "保存に失敗しました");
    return data;
  }

  async function sendStatusUpdateToSheet(type, card) {
    if (!gasUrl) return;
    try {
      updateSyncUI("☁️ 保存中...");
      await gasPost({
        action: "updateStatus",
        type,
        zh: card.zh,
        status: card.status,
        updatedAt: card.updatedAt,
        card
      });
      updateSyncUI("☁️ 同期済み");
    } catch (err) {
      console.warn("ステータス自動同期エラー:", err);
      updateSyncUI("☁️ 同期エラー", true);
    }
  }

  function mergeCardLists(localList, remoteList, isWord) {
    const map = new Map();
    const order = [];

    for (const r of remoteList) {
      const item = {
        ...r,
        pos: isWord ? (window.ZhDict ? window.ZhDict.normalizePos(r.pos) : r.pos || "その他") : undefined,
        updatedAt: Number(r.updatedAt) || 0
      };
      map.set(item.zh, item);
      order.push(item.zh);
    }

    for (const l of localList) {
      const existing = map.get(l.zh);
      if (!existing) {
        map.set(l.zh, { ...l });
        order.push(l.zh);
      } else {
        // updatedAt が新しい方のステータスを優先
        const localTime = Number(l.updatedAt) || 0;
        const remoteTime = Number(existing.updatedAt) || 0;
        if (localTime > remoteTime) {
          existing.status = l.status;
          existing.updatedAt = localTime;
        }
        // メタ情報は空欄を補う
        if (isWord) {
          if (!existing.pinyin && l.pinyin) existing.pinyin = l.pinyin;
          if (!existing.katakana && l.katakana) existing.katakana = l.katakana;
          if ((!existing.pos || existing.pos === "その他") && l.pos) existing.pos = l.pos;
        }
        if (!existing.ja && l.ja) existing.ja = l.ja;
      }
    }

    return order.map((zh, idx) => ({ ...map.get(zh), id: idx }));
  }

  async function syncBidirectional(silent = false) {
    if (!gasUrl) {
      if (!silent) setSyncLog("先にGoogle Apps ScriptのURLを入力・保存してください。", "error");
      return;
    }
    try {
      updateSyncUI("🔄 同期中...");
      if (!silent) setSyncLog("スプレッドシートと双方向同期しています...");

      const remote = await gasGet();
      const mergedWords = mergeCardLists(decks.word.cards, remote.words || [], true);
      const mergedSentences = mergeCardLists(decks.sentence.cards, remote.sentences || [], false);

      decks.word.cards = mergedWords;
      decks.sentence.cards = mergedSentences;
      syncTextareasFromCards();
      saveToStorage();
      rebuildDeck("word", false);
      rebuildDeck("sentence", false);

      await gasPost({
        action: "saveAll",
        words: mergedWords,
        sentences: mergedSentences
      });

      updateSyncUI("☁️ 同期済み");
      if (!silent) {
        setSyncLog(`同期完了！（単語 ${mergedWords.length}件・文章 ${mergedSentences.length}件）`, "success");
      }
    } catch (err) {
      updateSyncUI("☁️ 同期エラー", true);
      if (!silent) setSyncLog(`同期に失敗しました: ${err.message}`, "error");
    }
  }

  async function pullFromSheet() {
    if (!gasUrl) {
      setSyncLog("先にGoogle Apps ScriptのURLを入力・保存してください。", "error");
      return;
    }
    try {
      updateSyncUI("⬇️ 読込中...");
      setSyncLog("スプレッドシートからデータを読み込んでいます...");
      const remote = await gasGet();
      decks.word.cards = (remote.words || []).map((w, idx) => ({
        ...w,
        id: idx,
        pos: window.ZhDict ? window.ZhDict.normalizePos(w.pos) : w.pos || "その他"
      }));
      decks.sentence.cards = (remote.sentences || []).map((s, idx) => ({ ...s, id: idx }));
      syncTextareasFromCards();
      saveToStorage();
      rebuildDeck("word", false);
      rebuildDeck("sentence", false);
      updateSyncUI("☁️ 同期済み");
      setSyncLog(`読み込み完了！（単語 ${decks.word.cards.length}件・文章 ${decks.sentence.cards.length}件）`, "success");
    } catch (err) {
      updateSyncUI("☁️ 同期エラー", true);
      setSyncLog(`読み込みに失敗しました: ${err.message}`, "error");
    }
  }

  async function pushToSheet(silent = false) {
    if (!gasUrl) {
      if (!silent) setSyncLog("先にGoogle Apps ScriptのURLを入力・保存してください。", "error");
      return;
    }
    try {
      updateSyncUI("⬆️ 保存中...");
      if (!silent) setSyncLog("スプレッドシートへ保存しています...");
      await gasPost({
        action: "saveAll",
        words: decks.word.cards,
        sentences: decks.sentence.cards
      });
      updateSyncUI("☁️ 同期済み");
      if (!silent) {
        setSyncLog(`スプレッドシートへ保存しました！（単語 ${decks.word.cards.length}件・文章 ${decks.sentence.cards.length}件）`, "success");
      }
    } catch (err) {
      updateSyncUI("☁️ 同期エラー", true);
      if (!silent) setSyncLog(`保存に失敗しました: ${err.message}`, "error");
    }
  }

  function loadWordSample() {
    inputWordZh.value = SAMPLE_WORDS.zh;
    inputWordPinyin.value = SAMPLE_WORDS.pinyin;
    inputWordKatakana.value = SAMPLE_WORDS.katakana;
    inputWordPos.value = SAMPLE_WORDS.pos;
    inputWordJa.value = SAMPLE_WORDS.ja;
    chkWordIgnoreEmpty.checked = true;
    updateWordInputStats();
    saveToStorage();
    showNotice(wordSavedNotice, "単語サンプル（ピンイン・カタカナ・品詞付き）を入力しました");
  }

  function loadSentenceSample() {
    inputSentenceZh.value = SAMPLE_SENTENCES.zh;
    inputSentenceJa.value = SAMPLE_SENTENCES.ja;
    chkSentenceIgnoreEmpty.checked = true;
    updateSentenceInputStats();
    saveToStorage();
    showNotice(sentenceSavedNotice, "文章サンプルを入力しました");
  }

  // --- Event Binding ---
  function bindEvents() {
    modeWordBtn.addEventListener("click", () => switchMode("word"));
    modeSentenceBtn.addEventListener("click", () => switchMode("sentence"));
    tabStudy.addEventListener("click", () => switchTab("study"));
    tabInput.addEventListener("click", () => switchTab("input"));
    tabSync.addEventListener("click", () => switchTab("sync"));

    btnGoInput.addEventListener("click", () => switchTab("input"));
    btnLoadSampleEmpty.addEventListener("click", () => {
      if (activeMode === "word") {
        loadWordSample();
        buildWordCards();
      } else {
        loadSentenceSample();
        buildSentenceCards();
      }
    });

    posFilterButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        activePosFilter = btn.dataset.pos || "all";
        updatePosFilterButtonsUI();
        rebuildDeck("word", false);
      });
    });

    [inputWordZh, inputWordPinyin, inputWordKatakana, inputWordPos, inputWordJa].forEach((el) => {
      el.addEventListener("input", () => { updateWordInputStats(); saveToStorage(); });
    });
    chkWordIgnoreEmpty.addEventListener("change", () => { updateWordInputStats(); saveToStorage(); });
    btnExtractFromSentences.addEventListener("click", extractWordsFromSentenceInput);
    btnAutoFillReading.addEventListener("click", autoFillWordReadings);
    btnLoadWordSample.addEventListener("click", loadWordSample);
    btnClearWordInputs.addEventListener("click", () => {
      inputWordZh.value = "";
      inputWordPinyin.value = "";
      inputWordKatakana.value = "";
      inputWordPos.value = "";
      inputWordJa.value = "";
      updateWordInputStats();
      saveToStorage();
    });
    btnCreateWordCards.addEventListener("click", buildWordCards);

    [inputSentenceZh, inputSentenceJa].forEach((el) => {
      el.addEventListener("input", () => { updateSentenceInputStats(); saveToStorage(); });
    });
    chkSentenceIgnoreEmpty.addEventListener("change", () => { updateSentenceInputStats(); saveToStorage(); });
    btnLoadSentenceSample.addEventListener("click", loadSentenceSample);
    btnClearSentenceInputs.addEventListener("click", () => {
      inputSentenceZh.value = "";
      inputSentenceJa.value = "";
      updateSentenceInputStats();
      saveToStorage();
    });
    btnSendToWordExtractor.addEventListener("click", extractWordsFromSentenceInput);
    btnCreateSentenceCards.addEventListener("click", buildSentenceCards);

    btnReveal.addEventListener("click", toggleReveal);
    btnHintReading.addEventListener("click", toggleReadingHint);
    btnPrev.addEventListener("click", goPrev);
    btnNext.addEventListener("click", goNext);
    btnAgain.addEventListener("click", () => markStatus("again"));
    btnMastered.addEventListener("click", () => markStatus("mastered"));
    btnSpeak.addEventListener("click", speakCurrentChinese);
    btnQuickSync.addEventListener("click", () => syncBidirectional(false));

    chkAlwaysShowReading.addEventListener("change", (e) => {
      alwaysShowReading = e.target.checked;
      saveToStorage();
      renderStudyView();
    });
    chkUnmasteredOnly.addEventListener("change", (e) => {
      decks[activeMode].unmasteredOnly = e.target.checked;
      rebuildDeck(activeMode, false);
    });
    btnShuffle.addEventListener("click", () => rebuildDeck(activeMode, true));
    btnResetOrder.addEventListener("click", () => rebuildDeck(activeMode, false));

    // Sync controls
    btnSaveGasUrl.addEventListener("click", () => {
      gasUrl = inputGasUrl.value.trim();
      saveToStorage();
      updateSyncUI();
      if (gasUrl) {
        setSyncLog("URLを保存しました。「双方向同期」を押して接続を確認してください。", "success");
      } else {
        setSyncLog("同期URLをクリアしました。", "info");
      }
    });
    chkAutoSync.addEventListener("change", (e) => {
      autoSync = e.target.checked;
      saveToStorage();
    });
    btnSyncBidirectional.addEventListener("click", () => syncBidirectional(false));
    btnSyncPull.addEventListener("click", pullFromSheet);
    btnSyncPush.addEventListener("click", () => pushToSheet(false));

    btnCopyShareUrl.addEventListener("click", async () => {
      if (!gasUrl) {
        showNotice(shareUrlNotice, "先に上の欄にGASのURLを保存してください");
        return;
      }
      const base = `${window.location.origin}${window.location.pathname}`;
      const shareUrl = `${base}?sync=${encodeURIComponent(gasUrl)}`;
      try {
        await navigator.clipboard.writeText(shareUrl);
        showNotice(shareUrlNotice, "同期設定付きURLをコピーしました！");
      } catch {
        prompt("このURLをコピーしてiPhone/iPadで開いてください:", shareUrl);
      }
    });

    btnCopyGasCode.addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(GAS_TEMPLATE_CODE);
        showNotice(gasCopyNotice, "GASコードをコピーしました！");
      } catch {
        prompt("GASコードをコピーしてください:", GAS_TEMPLATE_CODE);
      }
    });

    document.addEventListener("keydown", (e) => {
      if (activeTab !== "study" || decks[activeMode].cards.length === 0) return;
      const tag = e.target.tagName;
      if (tag === "TEXTAREA" || tag === "INPUT") return;

      if (e.code === "Space") { e.preventDefault(); toggleReveal(); }
      else if (e.key.toLowerCase() === "h") { e.preventDefault(); toggleReadingHint(); }
      else if (e.code === "ArrowLeft") { e.preventDefault(); goPrev(); }
      else if (e.code === "ArrowRight") { e.preventDefault(); goNext(); }
      else if (e.key === "1") { e.preventDefault(); markStatus("again"); }
      else if (e.key === "2") { e.preventDefault(); markStatus("mastered"); }
      else if (e.key.toLowerCase() === "a") { e.preventDefault(); speakCurrentChinese(); }
    });
  }

  init();
})();
