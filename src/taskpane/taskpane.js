/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { GoogleGenerativeAI } from "@google/generative-ai";

let currentOriginalText = "";
let currentDiff = []; // Stores the diff structure
let currentCheckMode = 'selection';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("check-selection-btn").onclick = () => checkSpelling('selection');
    document.getElementById("check-all-btn").onclick = () => checkSpelling('whole');
    document.getElementById("cancel-btn").onclick = resetUI;
    document.getElementById("save-key-btn").onclick = saveApiKey;

    // New Settings Toggle
    document.getElementById("settings-toggle").onclick = () => {
      document.getElementById("settings-panel").classList.toggle("hidden");
    };

    // Load saved API Key
    const savedKey = localStorage.getItem("gemini_api_key");
    if (savedKey) {
      document.getElementById("api-key-input").value = savedKey;
    } else {
      // Auto-open settings if no key found
      document.getElementById("settings-panel").classList.remove("hidden");
    }
  }
});

function saveApiKey() {
  const key = document.getElementById("api-key-input").value.trim();
  if (key) {
    localStorage.setItem("gemini_api_key", key);
    showStatus("Đã lưu!");
  } else {
    localStorage.removeItem("gemini_api_key");
    showStatus("Đã xóa!");
  }
}

function showStatus(msg) {
  const status = document.getElementById("save-status");
  status.innerText = msg;
  status.style.display = "inline";
  setTimeout(() => { status.style.display = "none"; }, 2000);
}

async function checkSpelling(mode) {
  currentCheckMode = mode;
  return Word.run(async (context) => {
    let range;
    if (mode === 'whole') {
      range = context.document.body;
    } else {
      range = context.document.getSelection();
    }

    range.load("text");
    await context.sync();

    const originalText = range.text;
    if (!originalText.trim()) {
      const msg = mode === 'whole' ? "Tài liệu trống!" : "Vui lòng bôi đen văn bản trước!";
      // Use a simple alert or inject into a specific error container if diff-view is hidden
      // For now, let's just alert or use the diff view as before but ensure it's visible
      document.getElementById("diff-view").innerHTML = `<div style='padding:20px; text-align:center; color:#a4262c;'>${msg}</div>`;
      document.getElementById("result-step").classList.remove("hidden");
      document.getElementById("instruction-step").classList.add("hidden");
      return;
    }

    currentOriginalText = originalText;

    const btnId = mode === 'whole' ? "check-all-btn" : "check-selection-btn";
    const originalBtnContent = document.getElementById(btnId).innerHTML;
    document.getElementById(btnId).innerHTML = '<span class="ms-Button-label">Đang xử lý...</span>';

    // Get API Key
    const apiKey = localStorage.getItem("gemini_api_key");
    if (!apiKey) {
      document.getElementById("diff-view").innerHTML = "<div style='padding:20px; text-align:center; color:#a4262c;'>Vui lòng nhập Gemini API Key trong phần Cài đặt!</div>";
      document.getElementById("result-step").classList.remove("hidden");
      document.getElementById("instruction-step").classList.add("hidden");
      document.getElementById(btnId).innerHTML = originalBtnContent;
      return;
    }

    try {
      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" });

      const prompt = `
        Bạn là một chuyên gia biên tập tiếng Việt.
        Nhiệm vụ:
        1. Sửa lỗi chính tả, lỗi đánh máy (VD: "dukhách" -> "du khách").
        2. Bổ sung từ bị thiếu trong các cụm từ, danh từ riêng hoặc chức danh phổ biến (VD: "Ban quản Núi Bà Đen" -> "Ban Quản lý Núi Bà Đen").
        3. Sửa lỗi ngữ pháp và dấu câu.

        QUAN TRỌNG: 
        - KHÔNG thay đổi phong cách văn bản.
        - KHÔNG viết lại câu nếu không cần thiết.
        - Giữ nguyên ý nghĩa gốc.
        
        Văn bản gốc: "${originalText}"
        
        Yêu cầu output JSON format:
        {
            "corrected": "Văn bản đã sửa hoàn chỉnh"
        }
      `;

      const result = await model.generateContent(prompt);
      const response = await result.response;
      let text = response.text();

      console.log("Raw AI response:", text); // Debug log

      // Robust JSON extraction
      const jsonStartIndex = text.indexOf('{');
      const jsonEndIndex = text.lastIndexOf('}');

      if (jsonStartIndex !== -1 && jsonEndIndex !== -1 && jsonEndIndex > jsonStartIndex) {
        text = text.substring(jsonStartIndex, jsonEndIndex + 1);
      } else {
        throw new Error("Không tìm thấy dữ liệu JSON hợp lệ trong phản hồi của AI.");
      }

      let data;
      try {
        data = JSON.parse(text);
      } catch (e) {
        console.error("JSON Parse Error:", e);
        throw new Error("Lỗi phân tích dữ liệu từ AI.");
      }

      // 1. Compute LCS Diff
      currentDiff = computeDiff(originalText, data.corrected);

      // 2. Show Diff in UI
      renderDiffUI();

      // 3. Highlight errors in Document
      await highlightErrorsInDocument(context, range);

      document.getElementById("instruction-step").classList.add("hidden");
      document.getElementById("result-step").classList.remove("hidden");

    } catch (error) {
      console.error(error);
      document.getElementById("diff-view").innerHTML = `<div style='padding:20px; text-align:center; color:#a4262c;'>Lỗi: ${error.message}</div>`;
      document.getElementById("result-step").classList.remove("hidden");
      document.getElementById("instruction-step").classList.add("hidden");
    } finally {
      document.getElementById(btnId).innerHTML = originalBtnContent;
    }
  });
}

// --- LCS Diff Algorithm ---
function computeDiff(original, corrected) {
  const originalWords = original.trim().split(/\s+/);
  const correctedWords = corrected.trim().split(/\s+/);

  const m = originalWords.length;
  const n = correctedWords.length;
  const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (originalWords[i - 1] === correctedWords[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  let i = m, j = n;
  const diff = [];

  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && originalWords[i - 1] === correctedWords[j - 1]) {
      diff.unshift({ type: 'keep', value: originalWords[i - 1] });
      i--; j--;
    } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
      diff.unshift({ type: 'add', value: correctedWords[j - 1] });
      j--;
    } else {
      diff.unshift({ type: 'delete', value: originalWords[i - 1] });
      i--;
    }
  }

  // Post-processing to group contiguous items
  const groupedDiff = [];
  let k = 0;
  while (k < diff.length) {
    const item = diff[k];

    if (item.type === 'keep') {
      // Collect contiguous keeps
      let keeps = [item.value];
      let currentK = k + 1;
      while (currentK < diff.length && diff[currentK].type === 'keep') {
        keeps.push(diff[currentK].value);
        currentK++;
      }
      groupedDiff.push({ type: 'keep', value: keeps.join(" ") });
      k = currentK;
      continue;
    }

    // Collect contiguous deletes and adds
    let deletes = [];
    let adds = [];

    let currentK = k;
    while (currentK < diff.length && (diff[currentK].type === 'delete' || diff[currentK].type === 'add')) {
      if (diff[currentK].type === 'delete') deletes.push(diff[currentK].value);
      else if (diff[currentK].type === 'add') adds.push(diff[currentK].value);
      currentK++;
    }

    if (deletes.length > 0 && adds.length > 0) {
      // Replace block
      groupedDiff.push({
        type: 'replace',
        oldValue: deletes.join(" "),
        newValue: adds.join(" ")
      });
    } else if (deletes.length > 0) {
      // Pure delete block (possibly multiple words)
      groupedDiff.push({
        type: 'delete',
        value: deletes.join(" ")
      });
    } else if (adds.length > 0) {
      // Pure add block
      groupedDiff.push({
        type: 'add',
        value: adds.join(" ")
      });
    }

    k = currentK;
  }

  return groupedDiff;
}

// --- UI Rendering ---
function renderDiffUI() {
  let html = "";
  let errorIndex = 0;
  let hasErrors = false;
  let errorCount = 0;

  // We need to look ahead/behind to get context if we want.
  // For now, let's just list the errors.

  currentDiff.forEach((item, index) => {
    if (item.type === 'replace' || item.type === 'delete' || item.type === 'add') {
      hasErrors = true;
      errorCount++;
      item.errorIndex = errorIndex++;

      // Determine Label and Color
      let label = "Sửa lỗi";
      let labelClass = "label-error"; // CSS class for styling
      let icon = "Edit";

      if (item.type === 'add') {
        label = "Gợi ý bổ sung";
        labelClass = "label-suggestion";
        icon = "Add";
      } else if (item.type === 'delete') {
        label = "Gợi ý lược bỏ";
        labelClass = "label-delete";
        icon = "Delete";
      }

      // Context
      let contextStr = "... ";
      if (index > 0) {
        let prevItem = currentDiff[index - 1];
        if (prevItem.type === 'keep') contextStr += prevItem.value.split(" ").slice(-5).join(" ") + " ";
      }

      // Highlight the change in context
      if (item.type === 'replace') contextStr += `<strong style="color:#d13438; text-decoration:line-through">${item.oldValue}</strong>`;
      else if (item.type === 'delete') contextStr += `<strong style="color:#a4262c; text-decoration:line-through">${item.value}</strong>`;

      // For 'add', we don't show old value in context, just the insertion point essentially.
      // But to make it clear, let's show the NEW value in Green in the context? 
      // Or just keep context as surrounding words.
      // Let's stick to showing the "Old" state in context for replace/delete.
      // For Add, there is no old state.

      if (index < currentDiff.length - 1) {
        let nextItem = currentDiff[index + 1];
        if (nextItem.type === 'keep') contextStr += " " + nextItem.value.split(" ").slice(0, 3).join(" ");
      }
      contextStr += " ...";

      // Display Values
      let oldValDisplay = item.type === 'add' ? '(Trống)' : (item.type === 'replace' ? item.oldValue : item.value);
      let newValDisplay = item.type === 'delete' ? '(Xóa)' : (item.type === 'replace' ? item.newValue : item.value);

      html += `
            <div class="error-card" id="error-card-${item.errorIndex}">
                <div class="error-header">
                    <span class="error-label ${labelClass}">${label}</span>
                </div>
                <div class="error-info">
                    <div class="error-context">${contextStr}</div>
                    <div class="word-change">
                        <span class="old-word">${oldValDisplay}</span>
                        <i class="ms-Icon ms-Icon--Forward"></i>
                        <span class="new-word">${newValDisplay}</span>
                    </div>
                </div>
                <button class="fix-btn" onclick="window.fixSpecificError(${item.errorIndex})">Áp dụng</button>
            </div>
            `;
    }
  });

  if (!hasErrors) {
    html = "<div style='padding:40px; text-align:center; color:#107c10;'><i class='ms-Icon ms-Icon--CheckMark' style='font-size:32px; display:block; margin-bottom:10px;'></i>Tuyệt vời! Không tìm thấy lỗi nào.</div>";
  }

  document.getElementById("diff-view").innerHTML = html;
  document.getElementById("error-count").innerText = `${errorCount} lỗi`;
}

// Expose function to window for onclick
window.fixSpecificError = async (index) => {
  await applySpecificCorrection(index);
};

// --- Word Interaction ---
async function highlightErrorsInDocument(context, range) {
  range.font.highlightColor = null; // Clear old highlights

  // We need a more robust way to map diff items to document content.
  // Instead of complex index mapping which can fail with punctuation,
  // let's try a sequential search strategy within the selection range.

  // We will iterate through the 'keep' and 'error' items in order.
  // We search for the text. Since 'search' returns all occurrences, 
  // we need to track our "current position" in the range, but Word JS API doesn't give easy offsets.

  // Alternative: Split the range into words (ranges) and iterate? Slow.

  // Let's stick to the occurrence counting method but improve tokenization.
  // We will treat punctuation as separate tokens if possible, or just use the raw values from diff.

  // 1. Map each error word to its occurrence index (e.g., "thich" is the 2nd "thich" in the text)
  const wordCounts = {}; // Total counts in original text
  const errorOccurrences = {}; // Map: "word_occurrenceIndex" -> true (if it's an error)

  let originalIdx = 0;

  // Re-construct original text flow from diff to count occurrences correctly
  currentDiff.forEach(item => {
    if (item.type === 'add') return; // Adds don't exist in original

    const word = (item.type === 'replace') ? item.oldValue : item.value;

    if (!wordCounts[word]) wordCounts[word] = 0;

    if (item.type === 'replace' || item.type === 'delete') {
      // Mark this specific occurrence as an error
      errorOccurrences[`${word}_${wordCounts[word]}`] = true;
    }

    wordCounts[word]++;
  });

  // 2. Perform Search and Highlight
  // We search for every unique word that has an error.
  const uniqueErrorWords = Object.keys(wordCounts).filter(word => {
    // Check if this word has ANY errors associated with it
    for (let i = 0; i < wordCounts[word]; i++) {
      if (errorOccurrences[`${word}_${i}`]) return true;
    }
    return false;
  });

  // Limit to prevent timeouts
  const wordsToProcess = uniqueErrorWords.slice(0, 50);

  for (const word of wordsToProcess) {
    // Skip very short words unless they are clearly errors
    if (word.length < 2 && !errorOccurrences[`${word}_0`]) continue;

    // Search within the selection
    const searchResults = range.search(word, { matchCase: false, matchWholeWord: true });
    searchResults.load("items");
    await context.sync();

    // Iterate through search results (occurrences)
    for (let i = 0; i < searchResults.items.length; i++) {
      if (errorOccurrences[`${word}_${i}`]) {
        // This is the N-th occurrence, and it is marked as an error
        searchResults.items[i].font.highlightColor = "Yellow";
      }
    }
  }
}

async function applySpecificCorrection(errorIndex) {
  await Word.run(async (context) => {
    let targetItem = null;
    for (const item of currentDiff) {
      if (item.errorIndex === errorIndex) {
        targetItem = item;
        break;
      }
    }

    if (!targetItem || targetItem.fixed) return;

    const range = currentCheckMode === 'whole' ? context.document.body : context.document.getSelection();
    await fixErrorInContext(context, range, targetItem);
  });
}

async function fixErrorInContext(context, range, targetItem) {
  let wordToFind = "";
  if (targetItem.type === 'replace') wordToFind = targetItem.oldValue;
  else if (targetItem.type === 'delete') wordToFind = targetItem.value;
  else return;

  // Find the N-th occurrence of this word corresponding to this error
  // We must skip occurrences that correspond to ALREADY FIXED items
  let occurrenceIndex = 0;
  for (const item of currentDiff) {
    if (item === targetItem) break;
    if (item.type !== 'add' && !item.fixed) {
      const w = (item.type === 'replace') ? item.oldValue : item.value;
      if (w === wordToFind) occurrenceIndex++;
    }
  }

  const searchResults = range.search(wordToFind, { matchCase: false, matchWholeWord: true });
  searchResults.load("items");
  await context.sync();

  if (occurrenceIndex < searchResults.items.length) {
    const targetRange = searchResults.items[occurrenceIndex];

    if (targetItem.type === 'replace') {
      targetRange.insertText(targetItem.newValue, "Replace");
      targetRange.font.highlightColor = null;
    } else if (targetItem.type === 'delete') {
      targetRange.delete();
    }

    // Mark as fixed so subsequent searches don't count it
    targetItem.fixed = true;

    await context.sync();

    // Update UI
    if (targetItem.errorIndex !== undefined) {
      const card = document.getElementById(`error-card-${targetItem.errorIndex}`);
      if (card) card.style.display = 'none';
    }
  }
}

async function applyAllCorrections() {
  return Word.run(async (context) => {
    const range = currentCheckMode === 'whole' ? context.document.body : context.document.getSelection();

    // Filter errors
    const errors = currentDiff.filter(item => (item.type === 'replace' || item.type === 'delete') && !item.fixed);

    // Process in REVERSE order to minimize index shifting impact (though our 'fixed' logic handles it too)
    // Reverse is generally safer for document mutations.
    for (let i = errors.length - 1; i >= 0; i--) {
      await fixErrorInContext(context, range, errors[i]);
    }

    resetUI();
  });
}

async function resetUI() {
  document.getElementById("instruction-step").classList.remove("hidden");
  document.getElementById("result-step").classList.add("hidden");
  document.getElementById("diff-view").innerHTML = "";
  currentOriginalText = "";
  currentDiff = [];

  try {
    await Word.run(async (context) => {
      // We try to clear highlights on both body and selection to be safe, 
      // or just use the last mode. Let's clear body if we can, or just selection.
      // Actually, clearing body highlight is safer to ensure clean state.
      const body = context.document.body;
      body.font.highlightColor = null;
      await context.sync();
    });
  } catch (error) {
    console.error("Error clearing highlights:", error);
  }
}

export async function run() { }
