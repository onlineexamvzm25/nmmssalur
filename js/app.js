(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);

  // Make the manual PDF button easy to notice after results are displayed.
  const pdfAttentionStyle = document.createElement("style");
  pdfAttentionStyle.textContent = `
    #downloadResultPdf.pdf-attention {
      background: linear-gradient(135deg, #ffdf35, #ff9f1c) !important;
      color: #172033 !important;
      border: 2px solid #fff2a6 !important;
      font-weight: 800 !important;
      box-shadow: 0 0 0 0 rgba(255, 183, 0, .75), 0 0 18px rgba(255, 183, 0, .8) !important;
      animation: pdfButtonPulse 1.15s ease-in-out infinite;
      transform-origin: center;
    }
    #downloadResultPdf.pdf-attention:hover { filter: brightness(1.08); }
    @keyframes pdfButtonPulse {
      0%, 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 183, 0, .65), 0 0 12px rgba(255, 183, 0, .65); }
      50% { transform: scale(1.045); box-shadow: 0 0 0 7px rgba(255, 183, 0, 0), 0 0 24px rgba(255, 183, 0, .95); }
    }
    @media (prefers-reduced-motion: reduce) { #downloadResultPdf.pdf-attention { animation: none; } }
  `;
  document.head.appendChild(pdfAttentionStyle);

  const requiredHeaders = [
    "QUESTIONID", "SECTIONID", "EXAMID", "QNAME",
    "OPT1", "OPT2", "OPT3", "OPT4", "ANSWER", "HINT"
  ];
  // IMAGEID is optional for older workbooks; missing or blank values default to 0.
  const allowedSections = new Set(["MA", "SCI", "SS", "MATH"]);
  const allowedAnswers = new Set(["A", "B", "C", "D"]);
  let supabase = null;
  let sessionToken = null;
  let parsedQuestions = [];

  const screens = ["homeScreen", "adminLoginScreen", "uploadScreen", "studentLoginScreen", "examScreen", "resultScreen"];
  function showScreen(id) {
    screens.forEach((screenId) => $(screenId).classList.toggle("hidden", screenId !== id));
  }
  function setMessage(id, message, type = "") {
    const el = $(id);
    el.textContent = message || "";
    el.className = `message ${type}`.trim();
  }
  function toast(message) {
    const el = $("toast");
    el.textContent = message;
    el.classList.remove("hidden");
    window.setTimeout(() => el.classList.add("hidden"), 3500);
  }
  function getClient() {
    const cfg = window.EXAM_APP_CONFIG || {};
    if (!cfg.SUPABASE_URL || cfg.SUPABASE_URL.includes("YOUR-PROJECT-REF") ||
        !cfg.SUPABASE_ANON_KEY || cfg.SUPABASE_ANON_KEY.includes("YOUR_SUPABASE")) {
      throw new Error("Configure SUPABASE_URL and SUPABASE_ANON_KEY in js/config.js first.");
    }
    if (!window.supabase?.createClient) throw new Error("Supabase JavaScript library did not load.");
    if (!supabase) supabase = window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);
    return supabase;
  }
  function cellText(value) {
    return value === null || value === undefined ? "" : String(value).trim();
  }
  function normalizeHeader(value) {
    return cellText(value).toUpperCase().replace(/\s+/g, "");
  }
  function readWorkbook(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error("Could not read the selected file."));
      reader.onload = (event) => {
        try {
          const workbook = XLSX.read(event.target.result, { type: "array" });
          if (!workbook.SheetNames.length) throw new Error("Workbook has no sheets.");
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
          if (raw.length < 2) throw new Error("The Questions sheet has no question rows.");
          const headers = raw[0].map(normalizeHeader);
          const missing = requiredHeaders.filter((header) => !headers.includes(header));
          if (missing.length) throw new Error(`Missing required columns: ${missing.join(", ")}`);
          const index = Object.fromEntries(requiredHeaders.map((h) => [h, headers.indexOf(h)]));
          index.IMAGEID = headers.indexOf("IMAGEID");
          const rows = raw.slice(1).filter((row) => row.some((value) => cellText(value) !== ""));
          const questions = rows.map((row, idx) => ({
            _excelRow: idx + 2,
            QUESTIONID: cellText(row[index.QUESTIONID]),
            SECTIONID: cellText(row[index.SECTIONID]).toUpperCase(),
            EXAMID: cellText(row[index.EXAMID]),
            QNAME: cellText(row[index.QNAME]),
            OPT1: cellText(row[index.OPT1]),
            OPT2: cellText(row[index.OPT2]),
            OPT3: cellText(row[index.OPT3]),
            OPT4: cellText(row[index.OPT4]),
            ANSWER: cellText(row[index.ANSWER]).toUpperCase(),
            HINT: cellText(row[index.HINT]),
            IMAGEID: index.IMAGEID >= 0 ? cellText(row[index.IMAGEID]) : "0"
          }));
          resolve(questions);
        } catch (error) {
          reject(error);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }
  function validateQuestions(questions, requestedExamId) {
    const errors = [];
    const seen = new Set();
    if (!questions.length) errors.push("No question rows found.");
    questions.forEach((q) => {
      const row = q._excelRow;
      const required = ["QUESTIONID", "SECTIONID", "EXAMID", "QNAME", "OPT1", "OPT2", "OPT3", "OPT4", "ANSWER"];
      required.forEach((key) => { if (!q[key]) errors.push(`Row ${row}: ${key} is required.`); });
      if (q.EXAMID !== requestedExamId) errors.push(`Row ${row}: EXAMID ${q.EXAMID} does not match ${requestedExamId}.`);
      if (!allowedSections.has(q.SECTIONID)) errors.push(`Row ${row}: SECTIONID must be MA, SCI, SS or MATH.`);
      if (!allowedAnswers.has(q.ANSWER)) errors.push(`Row ${row}: ANSWER must be A, B, C or D.`);
      if (q.IMAGEID !== "" && (!/^\d+$/.test(q.IMAGEID) || !Number.isSafeInteger(Number(q.IMAGEID)) || Number(q.IMAGEID) < 0)) {
        errors.push(`Row ${row}: IMAGEID must be a whole number greater than or equal to 0.`);
      }
      const duplicateKey = `${q.EXAMID}::${q.QUESTIONID}`;
      if (q.QUESTIONID && seen.has(duplicateKey)) errors.push(`Row ${row}: duplicate QUESTIONID ${q.QUESTIONID}.`);
      if (q.QUESTIONID) seen.add(duplicateKey);
    });
    return errors;
  }
  function renderPreview(questions, errors) {
    const summary = $("previewSummary");
    summary.classList.remove("hidden");
    summary.textContent = errors.length
      ? `Validation failed with ${errors.length} issue(s): ${errors.slice(0, 8).join(" ")}${errors.length > 8 ? " …" : ""}`
      : `${questions.length} questions validated. Review the preview, then confirm import.`;
    summary.style.background = errors.length ? "#fef2f2" : "#eff6ff";
    summary.style.color = errors.length ? "#991b1b" : "#173b67";
    const wrap = $("previewTableWrap");
    wrap.classList.remove("hidden");
    const shown = questions.slice(0, 12);
    const cols = ["QUESTIONID", "SECTIONID", "EXAMID", "QNAME", "OPT1", "OPT2", "OPT3", "OPT4", "ANSWER", "IMAGEID"];
    const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    wrap.innerHTML = `<table><thead><tr>${cols.map(c => `<th>${c}</th>`).join("")}</tr></thead><tbody>${
      shown.map(q => `<tr>${cols.map(c => `<td>${esc(q[c])}</td>`).join("")}</tr>`).join("")
    }</tbody></table>`;
    if (questions.length > shown.length) {
      wrap.insertAdjacentHTML("beforeend", `<p class="small muted" style="padding:10px">Showing ${shown.length} of ${questions.length} rows.</p>`);
    }
    $("importButton").classList.toggle("hidden", errors.length > 0);
  }

  $("adminChoice").addEventListener("click", () => {
    setMessage("loginMessage", "");
    $("adminUsername").value = "MAHESH";
    $("adminPassword").value = "";
    showScreen("adminLoginScreen");
  });
  $("studentChoice").addEventListener("click", () => {
    setMessage("studentLoginMessage", "");
    showScreen("studentLoginScreen");
  });
  document.querySelectorAll("[data-go-home]").forEach(btn => btn.addEventListener("click", () => showScreen("homeScreen")));

  $("adminLoginForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    setMessage("loginMessage", "Checking credentials…");
    try {
      const client = getClient();
      const { data, error } = await client.rpc("exam_admin_login", {
        p_username: $("adminUsername").value.trim(),
        p_password: $("adminPassword").value
      });
      if (error) throw error;
      if (!data?.success || !data?.session_token) throw new Error(data?.message || "Login failed.");
      sessionToken = data.session_token;
      sessionStorage.setItem("exam_admin_session_token", sessionToken);
      $("adminPassword").value = "";
      setMessage("loginMessage", "");
      $("uploadMessage").textContent = "";
      showScreen("uploadScreen");
    } catch (error) {
      setMessage("loginMessage", error.message || "Unable to log in.", "error");
    }
  });

  $("excelFile").addEventListener("change", () => {
    const file = $("excelFile").files[0];
    $("fileInfo").textContent = file ? `${file.name} • ${(file.size / 1024).toFixed(1)} KB` : "";
    parsedQuestions = [];
    $("previewSummary").classList.add("hidden");
    $("previewTableWrap").classList.add("hidden");
    $("importButton").classList.add("hidden");
    setMessage("uploadMessage", "");
  });

  $("previewButton").addEventListener("click", async () => {
    setMessage("uploadMessage", "");
    try {
      const file = $("excelFile").files[0];
      const examId = $("examId").value.trim();
      if (!examId) throw new Error("Enter the Exam ID.");
      if (!file) throw new Error("Choose an Excel file.");
      if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error("Choose an .xlsx or .xls file.");
      parsedQuestions = await readWorkbook(file);
      const errors = validateQuestions(parsedQuestions, examId);
      renderPreview(parsedQuestions, errors);
      if (errors.length) setMessage("uploadMessage", "Fix the validation issues before importing.", "error");
      else setMessage("uploadMessage", "Validation passed. Review the rows and confirm import.", "success");
    } catch (error) {
      parsedQuestions = [];
      $("importButton").classList.add("hidden");
      setMessage("uploadMessage", error.message || "Unable to validate workbook.", "error");
    }
  });

  $("uploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    setMessage("uploadMessage", "");
    if (!sessionToken) {
      setMessage("uploadMessage", "Your admin session has expired. Please log in again.", "error");
      showScreen("adminLoginScreen");
      return;
    }
    if (!parsedQuestions.length) {
      setMessage("uploadMessage", "Validate the Excel file before importing.", "error");
      return;
    }
    const examId = $("examId").value.trim();
    const questions = parsedQuestions.map(q => ({
      question_id: q.QUESTIONID,
      section_id: q.SECTIONID,
      exam_id: q.EXAMID,
      qname: q.QNAME,
      opt1: q.OPT1,
      opt2: q.OPT2,
      opt3: q.OPT3,
      opt4: q.OPT4,
      answer_key: q.ANSWER,
      hint: q.HINT || null,
      image_id: Number(q.IMAGEID || 0)
    }));
    const button = $("importButton");
    button.disabled = true;
    button.textContent = "Importing…";
    try {
      const client = getClient();
      const { data, error } = await client.rpc("exam_admin_import_questions", {
        p_session_token: sessionToken,
        p_exam_id: examId,
        p_questions: questions
      });
      if (error) throw error;
      if (!data?.success) throw new Error(data?.message || "Import was rejected.");
      toast(`Success! ${data.inserted_count} questions imported.`);
      sessionStorage.removeItem("exam_admin_session_token");
      sessionToken = null;
      $("uploadForm").reset();
      $("fileInfo").textContent = "";
      parsedQuestions = [];
      $("previewSummary").classList.add("hidden");
      $("previewTableWrap").classList.add("hidden");
      $("importButton").classList.add("hidden");
      setMessage("uploadMessage", "");
      showScreen("homeScreen");
    } catch (error) {
      setMessage("uploadMessage", error.message || "Import failed.", "error");
    } finally {
      button.disabled = false;
      button.textContent = "Confirm & Import Questions";
    }
  });

  $("logoutButton").addEventListener("click", () => {
    sessionStorage.removeItem("exam_admin_session_token");
    sessionToken = null;
    parsedQuestions = [];
    $("uploadForm").reset();
    $("fileInfo").textContent = "";
    $("previewSummary").classList.add("hidden");
    $("previewTableWrap").classList.add("hidden");
    $("importButton").classList.add("hidden");
    showScreen("homeScreen");
  });

  // Student exam flow. Saved answers are written only by Save & Next or Mark for Review.
  let currentAttempt = null;
  let examQuestions = [];
  let questionIndex = 0;
  let draftSelections = {};
  let savedStatuses = {};
  let timerHandle = null;
  let submitting = false;
  let latestResult = null;
  const sectionLabels = { MA: "MA", SCI: "SCI", SS: "SS", MATH: "MAT" };
  const sectionOrder = ["MA", "SCI", "SS", "MATH"];
  function normalizeSectionMarks(items) {
    const scores = Object.fromEntries(sectionOrder.map(code => [code, 0]));
    if (Array.isArray(items)) {
      items.forEach(row => {
        const code = String(row.section_id || "").trim().toUpperCase();
        if (Object.hasOwn(scores, code)) scores[code] = Number(row.correct_count ?? row.total_marks ?? 0);
      });
    }
    return scores;
  }
  function renderSectionMarks(items) {
    const container = $("sectionMarks");
    container.replaceChildren();
    const scores = normalizeSectionMarks(items);
    for (const code of sectionOrder) {
      const cell = document.createElement("div"); cell.className = "section-mark";
      const name = document.createElement("span"); name.textContent = sectionLabels[code];
      const count = document.createElement("strong"); count.textContent = String(scores[code]);
      cell.append(name, count); container.appendChild(cell);
    }
    return scores;
  }
  function pdfText(value) { return value === null || value === undefined || String(value).trim() === "" ? "Not provided" : String(value); }
  function answerLetter(value) {
    const v = optionValue(value);
    return v ? String.fromCharCode(64 + Number(v.slice(-1))) : "Not answered";
  }
  function downloadResultPdf(options = {}) {
    if (!latestResult) return;
    const JsPdf = window.jspdf?.jsPDF;
    if (!JsPdf) { $("sectionMessage").textContent = "PDF library did not load. Check your internet connection."; return; }
    if (!Array.isArray(latestResult.question_review) || latestResult.question_review.length !== examQuestions.length) {
      $("sectionMessage").textContent = "Question-wise answer key and hint are missing from the submission response. Please apply the SQL update described below before downloading the complete PDF.";
      return;
    }
    const doc = new JsPdf({ unit: "mm", format: "a4" });
    const r = latestResult;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const left = 16, right = 16, contentWidth = pageWidth - left - right;
    let y = 20;
    const ensureSpace = (needed = 8) => { if (y + needed > pageHeight - 16) { doc.addPage(); y = 18; } };
    const writeLine = (label, value, opts = {}) => {
      const { bold = false, gap = 6, size = 10 } = opts;
      doc.setFont("helvetica", bold ? "bold" : "normal"); doc.setFontSize(size);
      const lines = doc.splitTextToSize(`${label}${value === undefined ? "" : value}`, contentWidth);
      ensureSpace(lines.length * gap + 2); doc.text(lines, left, y); y += lines.length * gap + 2;
    };
    doc.setFont("helvetica", "bold"); doc.setFontSize(17);
    doc.text("ONLINE EXAMINATION - RESULT", left, y); y += 11;
    writeLine("Student Name: ", r.student_name, { gap: 5.5 });
    writeLine("School Name: ", r.school_name, { gap: 5.5 });
    writeLine("Phone Number: ", r.phone_no, { gap: 5.5 });
    writeLine("Exam ID: ", r.exam_id, { gap: 5.5 });
    writeLine("Total Marks: ", `${r.total_marks} / ${r.max_marks}`, { gap: 5.5 });
    writeLine("Correct: ", `${r.correct_count}    Wrong: ${r.wrong_count}    Unanswered: ${r.unanswered_count}`, { gap: 5.5 });
    writeLine("Status: ", r.status, { gap: 5.5 });
    y += 2;
    writeLine("SECTION-WISE CORRECT ANSWERS", "", { bold: true, gap: 6, size: 12 });
    sectionOrder.forEach(code => writeLine(`${sectionLabels[code]}: `, r.scores[code], { gap: 5.5 }));
    doc.setDrawColor(100); doc.line(left, y + 1, pageWidth - right, y + 1); y += 9;
    writeLine("QUESTION-WISE REVIEW", "", { bold: true, gap: 7, size: 14 });

    const reviewById = new Map(r.question_review.map(item => [String(item.question_row_id ?? item.question_id ?? item.id ?? ""), item]));
    examQuestions.forEach((q, index) => {
      const detail = reviewById.get(String(q.id)) || r.question_review[index] || {};
      const selectedRaw = detail.selected_option ?? detail.student_answer ?? detail.response ?? q.selected_option ?? q.saved_answer;
      const keyRaw = detail.answer_key ?? detail.correct_answer ?? detail.key ?? detail.correct_option;
      const selectedLetter = answerLetter(selectedRaw);
      const keyLetter = answerLetter(keyRaw);
      const selectedOpt = optionValue(selectedRaw);
      const keyOpt = optionValue(keyRaw);
      const selectedText = selectedOpt ? q[`opt${selectedOpt.slice(-1)}`] : "Not answered";
      const keyText = keyOpt ? (detail.correct_answer_text ?? detail.key_text ?? q[`opt${keyOpt.slice(-1)}`]) : "Not provided";
      const question = detail.qname ?? detail.question_text ?? detail.question ?? q.qname;
      const hint = detail.hint ?? detail.question_hint ?? detail.explanation ?? "Not provided";
      const section = detail.section_id ?? q.section_id ?? "—";
      const startY = y;
      writeLine(`QUESTION ${index + 1}  |  SECTION: ${section}`, "", { bold: true, gap: 6, size: 11 });
      writeLine("Question: ", question, { gap: 5.2 });
      [1, 2, 3, 4].forEach(n => writeLine(`${String.fromCharCode(64+n)}. `, detail[`opt${n}`] ?? detail[`OPT${n}`] ?? q[`opt${n}`] ?? "", { gap: 5.0 }));
      // Color-code the student's response and draw vector icons so they render
      // reliably in the PDF across desktop and mobile PDF viewers.
      const responseIsCorrect = Boolean(selectedOpt && keyOpt && selectedOpt === keyOpt);
      const responseIsWrong = Boolean(selectedOpt && keyOpt && selectedOpt !== keyOpt);
      const responseLabel = selectedOpt
        ? `${selectedLetter} — ${selectedText}`
        : "Not answered";

      ensureSpace(8);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10);

      // Draw a green tick for correct answers, a red cross for wrong answers,
      // and a muted dash for unanswered questions.
      const iconX = left + 1.2;
      const iconY = y - 1.2;
      if (responseIsCorrect) {
        doc.setDrawColor(21, 128, 61);
        doc.setLineWidth(0.8);
        doc.line(iconX, iconY, iconX + 1.6, iconY + 1.7);
        doc.line(iconX + 1.6, iconY + 1.7, iconX + 4.5, iconY - 2.0);
        doc.setTextColor(21, 128, 61);
      } else if (responseIsWrong) {
        doc.setDrawColor(185, 28, 28);
        doc.setLineWidth(0.8);
        doc.line(iconX, iconY - 1.8, iconX + 4.2, iconY + 2.2);
        doc.line(iconX, iconY + 2.2, iconX + 4.2, iconY - 1.8);
        doc.setTextColor(185, 28, 28);
      } else {
        doc.setDrawColor(100, 116, 139);
        doc.setLineWidth(0.7);
        doc.line(iconX, iconY, iconX + 4, iconY);
        doc.setTextColor(100, 116, 139);
      }

      const responseLines = doc.splitTextToSize(`Student Response: ${responseLabel}`, contentWidth - 8);
      ensureSpace(responseLines.length * 5.2 + 2);
      doc.text(responseLines, left + 8, y);
      y += responseLines.length * 5.2 + 2;
      doc.setTextColor(0, 0, 0);

      // Keep the correct answer key green for quick review.
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10);
      doc.setTextColor(21, 128, 61);
      const keyValue = `${keyLetter}${keyOpt ? ` — ${keyText}` : keyRaw ? ` — ${keyRaw}` : "Not provided"}`;
      const keyLines = doc.splitTextToSize(`Correct Answer Key: ${keyValue}`, contentWidth);
      ensureSpace(keyLines.length * 5.2 + 2);
      doc.text(keyLines, left, y);
      y += keyLines.length * 5.2 + 2;
      doc.setTextColor(0, 0, 0);

      writeLine("Hint: ", hint, { gap: 5.2 });
      ensureSpace(8); doc.setDrawColor(145); doc.setLineWidth(0.35); doc.line(left, y + 1, pageWidth - right, y + 1); y += 8;
    });
    const filename = `exam_result_${String(r.exam_id || "exam").replace(/[^a-zA-Z0-9_-]/g, "")}.pdf`;
    const pdfBlob = doc.output("blob");
    const pdfUrl = URL.createObjectURL(pdfBlob);

    // Open in a tab created directly from the student's submit action when possible.
    if (options.openWindow && !options.openWindow.closed) {
      options.openWindow.location.href = pdfUrl;
    }

    // Always trigger a download so the student has a saved copy.
    const link = document.createElement("a");
    link.href = pdfUrl;
    link.download = filename;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();

    // Give the browser time to start both the PDF tab and download.
    window.setTimeout(() => URL.revokeObjectURL(pdfUrl), 60000);
  }
  $("downloadResultPdf").addEventListener("click", () => {
    $("downloadResultPdf").classList.remove("pdf-attention");
    downloadResultPdf();
  });

  function unwrapQuestions(data) {
    if (Array.isArray(data)) return data;
    if (Array.isArray(data?.questions)) return data.questions;
    if (Array.isArray(data?.data)) return data.data;
    if (Array.isArray(data?.result)) return data.result;
    if (data && typeof data === "object" && Array.isArray(data?.exam_questions)) return data.exam_questions;
    throw new Error("Questions were returned in an unexpected format. Please share the exam_get_questions function definition.");
  }
  function optionValue(value) {
    const v = String(value || "").trim().toUpperCase();
    if (["A", "OPT1", "OPTION1", "1"].includes(v)) return "OPT1";
    if (["B", "OPT2", "OPTION2", "2"].includes(v)) return "OPT2";
    if (["C", "OPT3", "OPTION3", "3"].includes(v)) return "OPT3";
    if (["D", "OPT4", "OPTION4", "4"].includes(v)) return "OPT4";
    return "";
  }
  function questionId(q) { return q.id ?? q.question_row_id ?? q.questionRowId; }
  function questionText(q) { return q.qname ?? q.QNAME ?? q.question ?? q.question_text ?? ""; }
  function optionText(q, n) { return q[`opt${n}`] ?? q[`OPT${n}`] ?? ""; }
  function normalizeQuestion(q) {
    return {
      ...q,
      id: questionId(q),
      qname: questionText(q),
      opt1: optionText(q, 1),
      opt2: optionText(q, 2),
      opt3: optionText(q, 3),
      opt4: optionText(q, 4),
      section_id: q.section_id ?? q.SECTIONID ?? "",
      image_id: Number(q.image_id ?? q.IMAGEID ?? q.imageid ?? 0) || 0
    };
  }
  function showStudentError(message) { setMessage("studentLoginMessage", message, "error"); }
  function paintQuestionNav() {
    const nav = $("questionNav");
    nav.innerHTML = "";
    examQuestions.forEach((q, i) => {
      const b = document.createElement("button");
      b.type = "button"; b.className = "question-nav-item";
      const state = savedStatuses[String(q.id)];
      b.classList.add(state === "review" ? "nav-review" : state === "saved" ? "nav-saved" : "nav-empty");
      b.textContent = String(i + 1); b.setAttribute("aria-label", `Question ${i + 1}${state ? `, ${state}` : ", not saved"}`);
      if (i === questionIndex) b.classList.add("active");
      b.addEventListener("click", () => { questionIndex = i; renderQuestion(); });
      nav.appendChild(b);
    });
  }
  // Exam images are stored in a public GitHub Pages folder named IMGS + Exam ID.
  // Try common extensions because IMAGEID does not store the file extension.
  const imageExtensions = ["png", "jpg", "jpeg", "webp", "gif","PNG", "JPG", "JPEG", "WEBP", "GIF"];
  let imageLoadSequence = 0;

  function currentExamIdForImages() {
    return String(currentAttempt?.exam_id ?? $("studentExamId")?.value ?? "").trim();
  }

  function imageFolderUrl() {
    const examId = currentExamIdForImages();
    if (!examId) return "";
    return new URL(`IMGS${encodeURIComponent(examId)}/`, document.baseURI).href;
  }

  function loadQuestionImage(imageId) {
    return new Promise((resolve, reject) => {
      const folder = imageFolderUrl();
      const id = Number(imageId);
      if (!folder || !Number.isSafeInteger(id) || id <= 0) {
        reject(new Error("This question does not have a valid image ID."));
        return;
      }
      let index = 0;
      const tryNext = () => {
        if (index >= imageExtensions.length) {
          reject(new Error(`Image ${id} was not found in ${`IMGS${currentExamIdForImages()}`}. Check its filename and extension.`));
          return;
        }
        const ext = imageExtensions[index++];
        const img = new Image();
        img.onload = () => resolve({ src: img.src, extension: ext });
        img.onerror = tryNext;
        img.src = new URL(`${id}.${ext}`, folder).href;
      };
      tryNext();
    });
  }

  function closeQuestionImageModal() {
    const modal = $("questionImageModal");
    modal.hidden = true;
    $("questionImageModalContent").replaceChildren();
    document.body.classList.remove("image-modal-open");
  }

  async function openQuestionImageModal(q, questionNumber) {
    const modal = $("questionImageModal");
    const content = $("questionImageModalContent");
    const title = $("questionImageModalTitle");
    const sequence = ++imageLoadSequence;
    title.textContent = `Question ${questionNumber} · Image ${q.image_id}`;
    content.replaceChildren();
    const loading = document.createElement("p");
    loading.className = "muted";
    loading.textContent = "Loading image…";
    content.appendChild(loading);
    modal.hidden = false;
    document.body.classList.add("image-modal-open");

    console.log("Question image mapping:", {
  displayedQuestion: questionNumber,
  questionId: q.question_id,
  imageId: q.image_id,
  folder: imageFolderUrl()
  });
  
    try {
      const loaded = await loadQuestionImage(q.image_id);
      if (sequence !== imageLoadSequence || modal.hidden) return;
      const img = document.createElement("img");
      img.src = loaded.src;
      img.alt = `Image for question ${questionNumber}`;
      content.replaceChildren(img);
    } catch (error) {
      if (sequence !== imageLoadSequence || modal.hidden) return;
      const message = document.createElement("p");
      message.className = "image-modal-message";
      message.textContent = `${error.message} Supported extensions: ${imageExtensions.map(x => "." + x).join(", ")}.`;
      content.replaceChildren(message);
    }
  }

  $("closeQuestionImageModal").addEventListener("click", closeQuestionImageModal);
  $("questionImageModal").addEventListener("click", event => {
    if (event.target === $("questionImageModal")) closeQuestionImageModal();
  });
  document.addEventListener("keydown", event => {
    if (event.key === "Escape" && !$("questionImageModal").hidden) closeQuestionImageModal();
  });

  function renderQuestion() {
    const q = examQuestions[questionIndex];
    if (!q) return;
    $("questionCounter").textContent = `Question ${questionIndex + 1} of ${examQuestions.length} · Section ${q.section_id || "—"}`;
    $("questionText").textContent = q.qname;
    // Show the image control only when Supabase returns a positive image_id.
    const oldImageButton = $("questionImageButton");
    if (oldImageButton) oldImageButton.remove();
    if (Number(q.image_id) > 0) {
      const imageButton = document.createElement("button");
      imageButton.id = "questionImageButton";
      imageButton.type = "button";
      imageButton.className = "question-image-action";
      imageButton.innerHTML = '<span class="image-symbol" aria-hidden="true">🖼️</span><span>View Question Image</span>';
      imageButton.addEventListener("click", () => openQuestionImageModal(q, questionIndex + 1));
      $("questionText").insertAdjacentElement("afterend", imageButton);
    }
    const list = $("optionList"); list.innerHTML = "";
    [1, 2, 3, 4].forEach(n => {
      const val = `OPT${n}`;
      const label = document.createElement("label"); label.className = "option-row";
      const radio = document.createElement("input"); radio.type = "radio"; radio.name = "selectedOption"; radio.value = val;
      radio.checked = draftSelections[String(q.id)] === val;
      const letter = document.createElement("strong"); letter.textContent = `${String.fromCharCode(64+n)}.`;
      const text = document.createElement("span"); text.textContent = q[`opt${n}`] || "";
      label.append(radio, letter, text); list.appendChild(label);
      radio.addEventListener("change", () => { draftSelections[String(q.id)] = val; });
    });
    $("previousQuestion").disabled = questionIndex === 0;
    $("skipQuestion").disabled = questionIndex === examQuestions.length - 1;
    $("saveNextQuestion").textContent = questionIndex === examQuestions.length - 1 ? "Save Answer" : "Save & Next";
    paintQuestionNav();
  }
  function formatRemaining(ms) {
    const sec = Math.max(0, Math.ceil(ms / 1000));
    return `${String(Math.floor(sec / 60)).padStart(2,"0")}:${String(sec % 60).padStart(2,"0")}`;
  }
  function startCountdown(deadline) {
    if (timerHandle) clearInterval(timerHandle);
    const end = new Date(deadline).getTime();
    const tick = async () => {
      const left = end - Date.now();
      $("examTimer").textContent = formatRemaining(left);
      if (left <= 0 && !submitting) {
        clearInterval(timerHandle); timerHandle = null;
        await submitExam(true);
      }
    };
    tick(); timerHandle = setInterval(tick, 1000);
  }
  async function loadExam(attempt) {
    const client = getClient();
    const { data, error } = await client.rpc("exam_get_questions", { p_attempt_id: attempt.attempt_id });
    if (error) throw error;
    console.table(
      (Array.isArray(data) ? data : data?.questions || []).map(q => ({
        question_id: q.question_id,
        question_row_id: q.question_row_id,
        image_id: q.image_id,
        IMAGEID: q.IMAGEID
      }))
    );
    examQuestions = unwrapQuestions(data).map(normalizeQuestion);
    if (!examQuestions.length) throw new Error("No questions were returned for this exam.");
    if (examQuestions.some(q => !q.id || !q.qname)) throw new Error("Question data is missing the row ID or question text.");
    currentAttempt = attempt;
    questionIndex = 0; draftSelections = {}; savedStatuses = {};
    // If the RPC returns saved answer/review fields, restore their display state.
    examQuestions.forEach(q => {
      const saved = optionValue(q.selected_option ?? q.saved_answer);
      if (saved) draftSelections[String(q.id)] = saved;
      if (q.is_review === true || q.is_review === "true") savedStatuses[String(q.id)] = "review";
      else if (saved) savedStatuses[String(q.id)] = "saved";
    });
    $("examMeta").textContent = `Exam ${attempt.exam_id} · Maximum marks ${attempt.max_marks ?? examQuestions.length}`;
    setMessage("examMessage", "");
    showScreen("examScreen"); renderQuestion(); startCountdown(attempt.deadline_at);
  }
  // Registration input validation: trim names, restrict phone to 10 digits,
  // and allow any characters in an Exam ID up to 6 characters.
  const studentNameInput = $("studentName");
  const schoolNameInput = $("schoolName");
  const phoneInput = $("phoneNo");
  const examIdInput = $("studentExamId");

  phoneInput.addEventListener("input", () => {
    const digitsOnly = phoneInput.value.replace(/\\D/g, "").slice(0, 10);
    if (phoneInput.value !== digitsOnly) phoneInput.value = digitsOnly;
    phoneInput.setCustomValidity("");
  });
  examIdInput.addEventListener("input", () => {
    if (examIdInput.value.length > 8) examIdInput.value = examIdInput.value.slice(0, 8);
  });

  $("studentLoginForm").addEventListener("submit", async event => {
    event.preventDefault();
    const studentName = studentNameInput.value.trim();
    const schoolName = schoolNameInput.value.trim();
    const phoneNo = phoneInput.value.trim();
    const examId = examIdInput.value.trim();

    studentNameInput.value = studentName;
    schoolNameInput.value = schoolName;
    examIdInput.value = examId;

    if (studentName.length < 3 || studentName.length > 12) {
      showStudentError("Student name must be 3 to 12 characters.");
      studentNameInput.focus();
      return;
    }
    if (schoolName.length < 4 || schoolName.length > 20) {
      showStudentError("School name must be 4 to 20 characters.");
      schoolNameInput.focus();
      return;
    }
    if (!/^[0-9]{10}$/.test(phoneNo)) {
      showStudentError("Phone number must contain exactly 10 digits (numbers only).");
      phoneInput.focus();
      return;
    }
    if (examId.length > 6) {
      showStudentError("Exam ID must be no more than 6 characters.");
      examIdInput.focus();
      return;
    }

    const btn = $("studentLoginButton"); btn.disabled = true; btn.textContent = "Checking details…";
    showStudentError("");
    try {
      const client = getClient();
      const { data, error } = await client.rpc("exam_register_or_login", {
        p_student_name: studentName,
        p_school_name: schoolName,
        p_exam_id: examId,
        p_phone_no: phoneNo,
        p_exam_password: $("studentExamPassword").value
      });
      if (error) throw error;
      if (!data?.success) throw new Error(data?.message || "Unable to register or continue.");
      if (!data.attempt_id || !data.deadline_at) throw new Error("Login succeeded, but the server did not return the attempt ID and deadline.");
      data.student_name = data.student_name || $("studentName").value.trim();
      data.school_name = data.school_name || $("schoolName").value.trim();
      data.phone_no = data.phone_no || $("phoneNo").value.trim();
      await loadExam(data);
    } catch (error) { showStudentError(error.message || "Student login failed."); }
    finally { btn.disabled = false; btn.textContent = "Register / Continue to Exam"; }
  });
  async function saveCurrent(isReview) {
    const q = examQuestions[questionIndex];
    const selected = draftSelections[String(q.id)] || null;
    const { error } = await getClient().rpc("exam_save_answer", {
      p_attempt_id: currentAttempt.attempt_id,
      p_question_row_id: q.id,
      p_selected_option: selected,
      p_is_review: isReview
    });
    if (error) throw error;
    q.selected_option = selected;
    q.saved_answer = selected;
    q.is_review = Boolean(isReview);
    if (isReview) savedStatuses[String(q.id)] = "review";
    else if (selected) savedStatuses[String(q.id)] = "saved";
    else delete savedStatuses[String(q.id)];
    paintQuestionNav();
  }
  $("saveNextQuestion").addEventListener("click", async () => {
    const btn = $("saveNextQuestion"); btn.disabled = true;
    try {
      await saveCurrent(false); setMessage("examMessage", "Answer saved.", "success");
      if (questionIndex < examQuestions.length - 1) { questionIndex++; renderQuestion(); }
    } catch (e) { setMessage("examMessage", e.message || "Could not save answer.", "error"); }
    finally { btn.disabled = false; }
  });
  $("reviewQuestion").addEventListener("click", async () => {
    const btn = $("reviewQuestion"); btn.disabled = true;
    try { await saveCurrent(true); setMessage("examMessage", "Saved and marked for review.", "success"); if (questionIndex < examQuestions.length - 1) { questionIndex++; renderQuestion(); } }
    catch (e) { setMessage("examMessage", e.message || "Could not mark for review.", "error"); }
    finally { btn.disabled = false; }
  });
  $("previousQuestion").addEventListener("click", () => { if (questionIndex > 0) { questionIndex--; renderQuestion(); setMessage("examMessage", "Previous question opened. No answer was saved by this action."); } });
  $("skipQuestion").addEventListener("click", () => { if (questionIndex < examQuestions.length - 1) { questionIndex++; renderQuestion(); setMessage("examMessage", "Question skipped without saving."); } });
  async function submitExam(auto = false) {
    if (!currentAttempt || submitting) return;
    if (!auto && !window.confirm("Submit your exam now? Unsaved selections will not be saved.")) return;

    // PDF download/open is manual only to avoid WebView/browser errors.
    submitting = true;
    $("submitExamButton").disabled = true;
    try {
      const { data, error } = await getClient().rpc("exam_submit_and_score", { p_attempt_id: currentAttempt.attempt_id });
      if (error) throw error;
      if (!data?.success) throw new Error(data?.message || "Could not submit the exam.");
      if (timerHandle) clearInterval(timerHandle); timerHandle = null;
      $("resultSummary").textContent = `Exam ${data.exam_id || currentAttempt.exam_id} · Score ${data.total_marks ?? "—"} / ${data.max_marks ?? currentAttempt.max_marks ?? examQuestions.length}`;
      const details = $("resultDetails"); details.innerHTML = "";
      const rows = [ ["Correct", data.correct_count], ["Wrong", data.wrong_count], ["Unanswered", data.unanswered_count], ["Status", data.status] ];
      rows.forEach(([label, value]) => { const p = document.createElement("p"); p.textContent = `${label}: ${value ?? "—"}`; details.appendChild(p); });
      const sectionRows = Array.isArray(data.section_marks) ? data.section_marks : null;
      const scores = renderSectionMarks(sectionRows);
      const sectionTotal = Object.values(scores).reduce((sum, n) => sum + n, 0);
      const scoreMatches = sectionRows && sectionTotal === Number(data.correct_count);
      $("sectionMarks").classList.toggle("hidden", !scoreMatches);
      const questionReview = Array.isArray(data.question_review) ? data.question_review
        : Array.isArray(data.question_details) ? data.question_details
        : Array.isArray(data.questions_review) ? data.questions_review : null;
      const completeReview = questionReview && questionReview.length === examQuestions.length;
      $("downloadResultPdf").disabled = !(scoreMatches && completeReview);
      $("sectionMessage").textContent = !scoreMatches
        ? "Section results are not yet available from Supabase. Apply the section-marks SQL update before downloading the PDF."
        : !completeReview
          ? "Section marks are ready. The PDF also needs question_review (all questions, student responses, answer keys and hints) from exam_submit_and_score. Update the SQL function before downloading."
          : "";
      latestResult = scoreMatches && completeReview ? {
        exam_id: data.exam_id || currentAttempt.exam_id,
        student_name: currentAttempt.student_name || $("studentName").value.trim(),
        school_name: currentAttempt.school_name || $("schoolName").value.trim(),
        phone_no: currentAttempt.phone_no || $("phoneNo").value.trim(),
        total_marks: data.total_marks, max_marks: data.max_marks ?? currentAttempt.max_marks,
        correct_count: data.correct_count, wrong_count: data.wrong_count,
        unanswered_count: data.unanswered_count, status: data.status, scores,
        question_review: questionReview
      } : null;
      showScreen("resultScreen");

      // PDF download stays manual. Highlight the button only when the complete PDF is ready.
      const pdfButton = $("downloadResultPdf");
      pdfButton.classList.toggle("pdf-attention", Boolean(latestResult));
      if (resultPdfWindow && !resultPdfWindow.closed) resultPdfWindow.close();
      resultPdfWindow = null;

      toast(auto ? "Time ended. Exam submitted." : "Exam submitted successfully.");
    } catch (e) {
      if (resultPdfWindow && !resultPdfWindow.closed) resultPdfWindow.close();
      setMessage("examMessage", e.message || "Submission failed.", "error");
    }
    finally { submitting = false; $("submitExamButton").disabled = false; }
  }
  $("submitExamButton").addEventListener("click", () => submitExam(false));
  $("returnHomeButton").addEventListener("click", () => {
    closeQuestionImageModal();
    if (timerHandle) clearInterval(timerHandle); timerHandle = null;
    currentAttempt = null; latestResult = null; examQuestions = []; draftSelections = {}; savedStatuses = {};
    $("studentLoginForm").reset(); showScreen("homeScreen");
  });

  // Never silently restore an admin session across browser tabs.
  sessionStorage.removeItem("exam_admin_session_token");
  showScreen("homeScreen");
})();
