(function () {
  "use strict";

  const state = { file: null, rows: [], filtered: [], headers: new Map(), mode: "plan" };
  const elements = {
    input: document.querySelector("#file-input"),
    dropzone: document.querySelector("#dropzone"),
    summary: document.querySelector("#file-summary"),
    fileName: document.querySelector("#file-name"),
    fileDetails: document.querySelector("#file-details"),
    remove: document.querySelector("#remove-file"),
    datePanel: document.querySelector("#date-panel"),
    date: document.querySelector("#start-date"),
    count: document.querySelector("#result-count"),
    preview: document.querySelector("#preview"),
    previewBody: document.querySelector("#preview-body"),
    previewHeaders: [...document.querySelectorAll("#preview thead th")],
    status: document.querySelector("#status"),
    generate: document.querySelector("#generate-button"),
    empty: document.querySelector("#empty-result"),
    loading: document.querySelector("#loading-overlay"),
    privacyButton: document.querySelector("#privacy-button"),
    privacyModal: document.querySelector("#privacy-modal"),
    privacyClose: document.querySelector("#privacy-close"),
    privacyConfirm: document.querySelector("#privacy-confirm"),
    modePlan: document.querySelector("#mode-plan"),
    modeCaju: document.querySelector("#mode-caju"),
  };

  const normalize = (value) => String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

  const text = (value) => (value == null ? "" : String(value).trim());

  function excelDateToLocal(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }
    if (typeof value === "number" && window.XLSX?.SSF) {
      const parts = XLSX.SSF.parse_date_code(value);
      return parts ? new Date(parts.y, parts.m - 1, parts.d) : null;
    }
    const raw = text(value);
    if (!raw) return null;
    const iso = raw.match(/^(\d{4})[-/]([01]?\d)[-/]([0-3]?\d)/);
    if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    const br = raw.match(/^([0-3]?\d)[-/]([01]?\d)[-/](\d{4})/);
    if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]));
    return null;
  }

  function isoDate(date) {
    if (!date) return "";
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function brDate(date) {
    if (!date) return "";
    return `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`;
  }

  function addDays(date, days) {
    const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    copy.setDate(copy.getDate() + days);
    return copy;
  }

  function buildHeaderMap(headerRow) {
    const map = new Map();
    headerRow.forEach((header, index) => map.set(normalize(header), index));
    return map;
  }

  function findColumn(...candidates) {
    for (const candidate of candidates) {
      const normalized = normalize(candidate);
      if (state.headers.has(normalized)) return state.headers.get(normalized);
    }
    return -1;
  }

  function valueFrom(row, ...headers) {
    const index = findColumn(...headers);
    return index >= 0 ? row[index] : "";
  }

  function formatCivilStatus(value) {
    const status = normalize(value);
    const mapping = {
      C: "Casado", CASADO: "Casado",
      S: "Solteiro", SOLTEIRO: "Solteiro",
      V: "Viúvo", VIUVO: "Viúvo",
      D: "Divorciado", DIVORCIADO: "Divorciado",
      SE: "Separado", SEPARADO: "Separado",
      O: "Outros", OUTROS: "Outros",
    };
    return mapping[status] || text(value);
  }

  function setStatus(message, type = "") {
    elements.status.textContent = message;
    elements.status.className = `status${type ? ` is-${type}` : ""}`;
  }

  function setMode(mode) {
    state.mode = mode;
    const plan = mode === "plan";
    elements.modePlan.classList.toggle("is-active", plan);
    elements.modeCaju.classList.toggle("is-active", !plan);
    elements.modePlan.setAttribute("aria-pressed", String(plan));
    elements.modeCaju.setAttribute("aria-pressed", String(!plan));
    elements.generate.lastChild.textContent = plan ? " Gerar e baixar layout" : " Gerar arquivo da Caju";
    if (state.filtered.length) updateFilteredRows();
  }

  function clearFile() {
    state.file = null;
    state.rows = [];
    state.filtered = [];
    state.headers = new Map();
    elements.input.value = "";
    elements.summary.hidden = true;
    elements.datePanel.hidden = true;
    elements.preview.hidden = true;
    elements.empty.hidden = false;
    elements.generate.disabled = true;
    elements.previewBody.replaceChildren();
    setStatus("");
  }

  function updateFilteredRows() {
    const selected = excelDateToLocal(elements.date.value);
    const admissionColumn = findColumn("Data Admis.", "Data Admis", "Data Admissão", "Data Admissao");
    state.filtered = state.rows.filter((row) => {
      const date = excelDateToLocal(row[admissionColumn]);
      return selected && date && date >= selected;
    });

    elements.count.textContent = `${state.filtered.length} ${state.filtered.length === 1 ? "pessoa" : "pessoas"}`;
    elements.generate.disabled = state.filtered.length === 0;
    elements.preview.hidden = state.filtered.length === 0;
    elements.empty.hidden = state.filtered.length > 0;
    elements.previewBody.replaceChildren();

    const cajuMode = state.mode === "caju";
    const headerLabels = cajuMode
      ? ["Nome completo", "CPF", "E-mail", "Telefone"]
      : ["Matrícula", "Nome", "Admissão", "Município"];
    elements.previewHeaders.forEach((header, index) => { header.textContent = headerLabels[index]; });

    state.filtered.slice(0, 5).forEach((row) => {
      const tr = document.createElement("tr");
      const mobile = `${digits(valueFrom(row, "DDD Celular"))}${digits(valueFrom(row, "Num. Celular"))}`;
      const values = cajuMode
        ? [
            valueFrom(row, "Nome completo", "Nome complet"),
            digits(valueFrom(row, "CPF")),
            valueFrom(row, "Email Princ", "Email Principal"),
            mobile || `${digits(valueFrom(row, "DDD Telefone"))}${digits(valueFrom(row, "Telefone"))}`,
          ]
        : [
            valueFrom(row, "Matricula", "Matrícula"),
            valueFrom(row, "Nome complet", "Nome completo"),
            brDate(excelDateToLocal(valueFrom(row, "Data Admis.", "Data Admissão"))),
            valueFrom(row, "Municipio", "Município"),
          ];
      values.forEach((value) => {
        const td = document.createElement("td");
        td.textContent = text(value);
        tr.appendChild(td);
      });
      elements.previewBody.appendChild(tr);
    });

    if (selected && state.filtered.length === 0) {
      setStatus("Não há admissões a partir da data escolhida.", "error");
    } else if (state.mode === "caju") {
      const missingEmail = state.filtered.filter((row) => !text(valueFrom(row, "Email Princ", "Email Principal"))).length;
      setStatus(
        missingEmail ? `${missingEmail} ${missingEmail === 1 ? "pessoa está" : "pessoas estão"} sem e-mail no efetivo.` : "",
        missingEmail ? "error" : "",
      );
    } else {
      setStatus("");
    }
  }

  async function loadFile(file) {
    if (!file || !/\.(xlsx|xls)$/i.test(file.name)) {
      setStatus("Escolha um arquivo Excel no formato XLSX ou XLS.", "error");
      return;
    }

    try {
      elements.loading.hidden = false;
      setStatus("Lendo a planilha…");
      await new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 420)));
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      if (matrix.length < 2) throw new Error("A planilha está vazia.");

      state.headers = buildHeaderMap(matrix[0]);
      const required = [
        { label: "Matricula", aliases: ["Matricula", "Matrícula"] },
        { label: "Nome completo", aliases: ["Nome complet", "Nome completo"] },
        { label: "Data Admis.", aliases: ["Data Admis.", "Data Admissão"] },
      ];
      const missing = required.filter((item) => findColumn(...item.aliases) < 0).map((item) => item.label);
      if (missing.length) throw new Error(`Coluna não encontrada: ${missing.join(", ")}.`);

      state.file = file;
      state.rows = matrix.slice(1).filter((row) => row.some((value) => text(value) !== ""));
      const admissionColumn = findColumn("Data Admis.");
      const dates = state.rows
        .map((row) => excelDateToLocal(row[admissionColumn]))
        .filter(Boolean)
        .sort((a, b) => a - b);
      if (!dates.length) throw new Error("Nenhuma data de admissão válida foi encontrada.");

      elements.fileName.textContent = file.name;
      elements.fileDetails.textContent = `${state.rows.length} registros encontrados`;
      elements.summary.hidden = false;
      elements.datePanel.hidden = false;
      elements.date.min = isoDate(dates[0]);
      elements.date.max = isoDate(dates[dates.length - 1]);
      elements.date.value = isoDate(dates[dates.length - 1]);
      updateFilteredRows();
    } catch (error) {
      clearFile();
      setStatus(error.message || "Não foi possível ler esse arquivo.", "error");
    } finally {
      elements.loading.hidden = true;
    }
  }

  function toLayoutRow(row) {
    const admission = excelDateToLocal(valueFrom(row, "Data Admis.", "Data Admissão"));
    const birth = excelDateToLocal(valueFrom(row, "Data Nasc.", "Data Nasc", "Data de nascimento"));
    const cpf = text(valueFrom(row, "CPF"));
    return [
      text(valueFrom(row, "Matricula", "Matrícula")),
      text(valueFrom(row, "Nome complet", "Nome completo")),
      isoDate(birth),
      "",
      text(valueFrom(row, "Sexo")),
      cpf,
      text(valueFrom(row, "Nome Mae", "Nome Mãe")),
      brDate(addDays(admission, 1)),
      "TITULAR",
      text(valueFrom(row, "Endereço", "Endereco")),
      text(valueFrom(row, "Num.Endereço", "Num.Endereco", "NrLogradouro")),
      text(valueFrom(row, "Compl.Ender.", "Compl.Ender")),
      text(valueFrom(row, "Bairro")),
      text(valueFrom(row, "Municipio", "Município")),
      text(valueFrom(row, "Cep", "CEP")),
      "",
      cpf,
      isoDate(admission),
      "",
      formatCivilStatus(valueFrom(row, "Est. Civil", "Estado Civil")),
    ];
  }

  function escapeXml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function base64Bytes(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }

  function cellXml(column, row, style, value) {
    const reference = `${XLSX.utils.encode_col(column)}${row}`;
    if (value === "" || value == null) return `<c r="${reference}" s="${style}"/>`;
    const safe = escapeXml(value);
    const preserve = /^\s|\s$|\s{2,}/.test(String(value)) ? ' xml:space="preserve"' : "";
    return `<c r="${reference}" s="${style}" t="inlineStr"><is><t${preserve}>${safe}</t></is></c>`;
  }

  async function generateWorkbook() {
    if (!state.filtered.length) return;
    try {
      elements.loading.hidden = false;
      const zip = await JSZip.loadAsync(base64Bytes(window.LAYOUT_TEMPLATE_BASE64));
      const sheetPath = "xl/worksheets/sheet1.xml";
      let xml = await zip.file(sheetPath).async("string");
      const sheetDataMatch = xml.match(/<sheetData>([\s\S]*?)<\/sheetData>/);
      if (!sheetDataMatch) throw new Error("A estrutura do modelo não pôde ser lida.");

      const fixedRows = [...sheetDataMatch[1].matchAll(/<row\b[^>]*\br="([1-4])"[^>]*>[\s\S]*?<\/row>/g)].map((match) => match[0]).join("");
      const styles = [36, 25, 37, 37, 25, 25, 25, 27, 26, 25, 25, 25, 25, 25, 25, 28, 25, 37, 29, 25];
      const dataRows = state.filtered.map(toLayoutRow).map((values, index) => {
        const rowNumber = index + 5;
        const cells = values.map((value, column) => cellXml(column, rowNumber, styles[column], value)).join("");
        return `<row r="${rowNumber}">${cells}</row>`;
      }).join("");
      const lastRow = state.filtered.length + 4;

      xml = xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${fixedRows}${dataRows}</sheetData>`);
      xml = xml.replace(/<dimension ref="[^"]+"\/>/, `<dimension ref="A1:T${lastRow}"/>`);
      xml = xml.replace(/(<autoFilter\b[^>]*\bref=")A4:T\d+("[^>]*\/>)/, `$1A4:T${lastRow}$2`);
      xml = xml.replace(/(<dataValidation\b[^>]*\bsqref=")E5:[^" ]+("[^>]*>)/, `$1E5:E${lastRow}$2`);
      xml = xml.replace(/(<dataValidation\b[^>]*\bsqref=")I5:[^" ]+("[^>]*>)/, `$1I5:I${lastRow}$2`);
      xml = xml.replace(
        /<dataValidation\b([^>]*)sqref="[^"]+"([^>]*)>((?:(?!<\/dataValidation>)[\s\S])*?<formula1>"Casado(?:(?!<\/dataValidation>)[\s\S])*?<\/dataValidation>)/,
        `<dataValidation$1sqref="T5:T${lastRow}"$2>$3`,
      );

      zip.file(sheetPath, xml);
      const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
      const selectedDate = elements.date.value;
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `Layout_Hapvida_Admissoes_${selectedDate}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(link.href), 60000);
      setStatus(`Layout gerado com ${state.filtered.length} registros.`, "success");
    } catch (error) {
      setStatus(error.message || "Não foi possível gerar o layout.", "error");
    } finally {
      elements.loading.hidden = true;
    }
  }

  function csvCell(value) {
    const string = String(value ?? "");
    return /[;"\r\n]/.test(string) ? `"${string.replace(/"/g, '""')}"` : string;
  }

  function digits(value) {
    return String(value ?? "").replace(/\D/g, "");
  }

  async function generateCajuCsv() {
    if (!state.filtered.length) return;
    try {
      elements.loading.hidden = false;
      const headers = ["Nome completo", "CPF", "Email", "Telefone Celular", "CEP", "Logradouro", "Numero", "Complemento", "Bairro", "Cidade", "Estado"];
      const rows = state.filtered.map((row) => {
        const mobile = `${digits(valueFrom(row, "DDD Celular"))}${digits(valueFrom(row, "Num. Celular"))}`;
        const phone = mobile || `${digits(valueFrom(row, "DDD Telefone"))}${digits(valueFrom(row, "Telefone"))}`;
        return [
          text(valueFrom(row, "Nome completo", "Nome complet")),
          digits(valueFrom(row, "CPF")),
          text(valueFrom(row, "Email Princ", "Email Principal")),
          phone,
          "", "", "", "", "", "", "",
        ];
      });
      const csv = [headers, ...rows].map((row) => row.map(csvCell).join(";")).join("\r\n");
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "colaboradores-exemplo (20).csv";
      document.body.appendChild(link);
      link.click();
      const objectUrl = link.href;
      link.remove();
      setTimeout(() => URL.revokeObjectURL(objectUrl), 60000);
      setStatus(`Arquivo da Caju gerado com ${state.filtered.length} registros.`, "success");
    } catch (error) {
      setStatus(error.message || "Não foi possível gerar o arquivo da Caju.", "error");
    } finally {
      elements.loading.hidden = true;
    }
  }

  function generateOutput() {
    return state.mode === "caju" ? generateCajuCsv() : generateWorkbook();
  }

  elements.input.addEventListener("change", () => loadFile(elements.input.files[0]));
  elements.remove.addEventListener("click", clearFile);
  elements.date.addEventListener("change", updateFilteredRows);
  elements.generate.addEventListener("click", generateOutput);
  elements.modePlan.addEventListener("click", () => setMode("plan"));
  elements.modeCaju.addEventListener("click", () => setMode("caju"));
  elements.privacyButton.addEventListener("click", () => elements.privacyModal.showModal());
  elements.privacyClose.addEventListener("click", () => elements.privacyModal.close());
  elements.privacyConfirm.addEventListener("click", () => elements.privacyModal.close());
  elements.privacyModal.addEventListener("click", (event) => {
    if (event.target === elements.privacyModal) elements.privacyModal.close();
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropzone.classList.add("is-dragging");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropzone.classList.remove("is-dragging");
    });
  });
  elements.dropzone.addEventListener("drop", (event) => loadFile(event.dataTransfer.files[0]));
})();
