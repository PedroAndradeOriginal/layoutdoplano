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
    dateStep: document.querySelector("#date-step"),
    dateLabel: document.querySelector("#date-label"),
    dateHelp: document.querySelector("#date-help"),
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
    modeCajuPayment: document.querySelector("#mode-caju-payment"),
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
    const cajuCard = mode === "caju-card";
    const cajuPayment = mode === "caju-payment";
    elements.modePlan.classList.toggle("is-active", plan);
    elements.modeCaju.classList.toggle("is-active", cajuCard);
    elements.modeCajuPayment.classList.toggle("is-active", cajuPayment);
    elements.modePlan.setAttribute("aria-pressed", String(plan));
    elements.modeCaju.setAttribute("aria-pressed", String(cajuCard));
    elements.modeCajuPayment.setAttribute("aria-pressed", String(cajuPayment));
    elements.dateStep.textContent = "02";
    elements.date.type = cajuPayment ? "month" : "date";
    elements.dateLabel.textContent = cajuPayment ? "Competência do pagamento" : "Admissões a partir de";
    elements.dateHelp.textContent = cajuPayment
      ? "Inclui quem esteve ativo em qualquer dia da competência."
      : "Inclui a data escolhida e as posteriores.";
    elements.generate.lastChild.textContent = plan
      ? " Gerar e baixar layout"
      : cajuCard
        ? " Gerar arquivo da Caju"
        : " Gerar pagamento mensal da Caju";
    if (state.rows.length) {
      configureDateControl();
      updateFilteredRows();
    }
  }

  function allAdmissionDates() {
    const admissionColumn = findColumn("Data Admis.", "Data Admis", "Data Admissão", "Data Admissao");
    return state.rows
      .map((row) => excelDateToLocal(row[admissionColumn]))
      .filter(Boolean)
      .sort((a, b) => a - b);
  }

  function configureDateControl() {
    const dates = allAdmissionDates();
    if (!dates.length) return;
    const first = isoDate(dates[0]);
    const last = isoDate(dates[dates.length - 1]);
    if (state.mode === "caju-payment") {
      elements.date.min = first.slice(0, 7);
      elements.date.max = last.slice(0, 7);
      elements.date.value = last.slice(0, 7);
    } else {
      elements.date.min = first;
      elements.date.max = last;
      elements.date.value = last;
    }
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
    const paymentMode = state.mode === "caju-payment";
    const selected = paymentMode
      ? excelDateToLocal(`${elements.date.value}-01`)
      : excelDateToLocal(elements.date.value);
    const admissionColumn = findColumn("Data Admis.", "Data Admis", "Data Admissão", "Data Admissao");
    const dismissalColumn = findColumn("Dt. Demissao", "Dt. Demissão", "Data Demissão", "Data Demissao");
    const monthEnd = selected ? new Date(selected.getFullYear(), selected.getMonth() + 1, 0) : null;
    state.filtered = state.rows.filter((row) => {
      const admission = excelDateToLocal(row[admissionColumn]);
      if (!selected || !admission) return false;
      if (!paymentMode) return admission >= selected;
      const dismissal = dismissalColumn >= 0 ? excelDateToLocal(row[dismissalColumn]) : null;
      return admission <= monthEnd && (!dismissal || dismissal >= selected);
    });

    elements.count.textContent = `${state.filtered.length} ${state.filtered.length === 1 ? "pessoa" : "pessoas"}`;
    elements.generate.disabled = state.filtered.length === 0;
    elements.preview.hidden = state.filtered.length === 0;
    elements.empty.hidden = state.filtered.length > 0;
    elements.previewBody.replaceChildren();

    const cajuCard = state.mode === "caju-card";
    const headerLabels = cajuCard
      ? ["Nome completo", "CPF", "E-mail", "Telefone"]
      : paymentMode
        ? ["Matrícula", "Nome", "Função", "CPF"]
        : ["Matrícula", "Nome", "Admissão", "Município"];
    elements.previewHeaders.forEach((header, index) => { header.textContent = headerLabels[index]; });

    state.filtered.slice(0, 5).forEach((row) => {
      const tr = document.createElement("tr");
      const mobile = `${digits(valueFrom(row, "DDD Celular"))}${digits(valueFrom(row, "Num. Celular"))}`;
      const values = cajuCard
        ? [
            valueFrom(row, "Nome completo", "Nome complet"),
            digits(valueFrom(row, "CPF")),
            valueFrom(row, "Email Princ", "Email Principal"),
            mobile || `${digits(valueFrom(row, "DDD Telefone"))}${digits(valueFrom(row, "Telefone"))}`,
          ]
        : paymentMode
          ? [
              valueFrom(row, "Matricula", "Matrícula"),
              valueFrom(row, "Nome complet", "Nome completo"),
              valueFrom(row, "Desc.Funcao", "Desc. Função", "Função"),
              valueFrom(row, "CPF"),
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
      setStatus(paymentMode ? "Não há pessoas ativas nessa competência." : "Não há admissões a partir da data escolhida.", "error");
    } else if (cajuCard) {
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
      const dates = allAdmissionDates();
      if (!dates.length) throw new Error("Nenhuma data de admissão válida foi encontrada.");

      elements.fileName.textContent = file.name;
      elements.fileDetails.textContent = `${state.rows.length} registros encontrados`;
      elements.summary.hidden = false;
      elements.datePanel.hidden = false;
      configureDateControl();
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

  function employeeKey(value) {
    const number = digits(value);
    return number ? number.padStart(6, "0") : "";
  }

  function excelSerial(value) {
    const date = excelDateToLocal(value);
    return date ? Math.round((Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - Date.UTC(1899, 11, 30)) / 86400000) : "";
  }

  function monthlyCellXml(column, row, style, cell = {}) {
    const reference = `${XLSX.utils.encode_col(column)}${row}`;
    const styleAttribute = style == null ? "" : ` s="${style}"`;
    if (cell.formula) {
      const cached = cell.value === "" || cell.value == null ? "" : `<v>${escapeXml(cell.value)}</v>`;
      return `<c r="${reference}"${styleAttribute}><f>${escapeXml(cell.formula)}</f>${cached}</c>`;
    }
    if (cell.value === "" || cell.value == null) return `<c r="${reference}"${styleAttribute}/>`;
    if (cell.type === "number") return `<c r="${reference}"${styleAttribute}><v>${Number(cell.value) || 0}</v></c>`;
    const preserve = /^\s|\s$|\s{2,}/.test(String(cell.value)) ? ' xml:space="preserve"' : "";
    return `<c r="${reference}"${styleAttribute} t="inlineStr"><is><t${preserve}>${escapeXml(cell.value)}</t></is></c>`;
  }

  function stylesFromTemplateRow(rowXml) {
    const styles = new Map();
    for (const match of rowXml.matchAll(/<c\b[^>]*\br="([A-Z]+)\d+"[^>]*>/g)) {
      const style = match[0].match(/\bs="(\d+)"/);
      styles.set(XLSX.utils.decode_col(match[1]), style ? Number(style[1]) : null);
    }
    return styles;
  }

  function benefitCells(sheet, sourceRow, targetRow) {
    const cells = [];
    for (let column = 13; column < 28; column += 1) {
      const source = sheet[`${XLSX.utils.encode_col(column)}${sourceRow}`];
      if (!source) {
        cells.push({ value: "" });
        continue;
      }
      const formula = source.f
        ? source.f.replace(/(\$?[A-Z]{1,3}\$?)\d+/g, (match, prefix) => `${prefix}${targetRow}`)
        : "";
      cells.push({
        value: source.v ?? "",
        type: source.t === "n" ? "number" : "text",
        formula,
      });
    }
    return cells;
  }

  function newEmployeeBenefits(targetRow) {
    return [
      { value: 0, type: "number" },
      { value: 0, type: "number" },
      { formula: `N${targetRow}*O${targetRow}` },
      { value: 0, type: "number" },
      { formula: `Q${targetRow}` },
      { value: 12.198, type: "number" },
      { value: 0, type: "number" },
      { value: 27.713, type: "number" },
      { value: 0, type: "number" },
      { value: 27.713, type: "number" },
      { value: 0, type: "number" },
      { value: 0, type: "number" },
      { formula: `(S${targetRow}*T${targetRow})+(U${targetRow}*V${targetRow})+(W${targetRow}*X${targetRow})+Y${targetRow}` },
      { value: 0, type: "number" },
      { value: "" },
    ];
  }

  function paymentStatus(row) {
    const status = normalize(valueFrom(row, "Sit. Folha", "Situação Folha", "Situacao Folha"));
    if (status === "D") return "DEMITIDO";
    if (status === "A") return "AFASTADO";
    if (status === "F") return "FÉRIAS";
    return status || "ATIVO";
  }

  function paymentBaseCells(row, templateRow) {
    const mobile = `${digits(valueFrom(row, "DDD Celular"))}${digits(valueFrom(row, "Num. Celular"))}`;
    const phone = mobile || `${digits(valueFrom(row, "DDD Telefone"))}${digits(valueFrom(row, "Telefone"))}`;
    return [
      { value: templateRow?.filial || "00301001" },
      { value: text(valueFrom(row, "Centro Custo")) },
      { value: employeeKey(valueFrom(row, "Matricula", "Matrícula")) },
      { value: text(valueFrom(row, "Nome completo", "Nome complet")) },
      { value: text(valueFrom(row, "Email Princ", "Email Principal")) },
      { value: phone },
      { value: text(valueFrom(row, "Desc.Funcao", "Desc. Função", "Função")) },
      { value: 0, type: "number" },
      { value: excelSerial(valueFrom(row, "Data Admis.", "Data Admissão")), type: "number" },
      { value: paymentStatus(row) },
      { value: excelSerial(valueFrom(row, "Dt. Demissao", "Dt. Demissão")), type: "number" },
      { value: text(valueFrom(row, "Desc. Depto", "Desc. Depto.", "Departamento")) },
      { value: text(valueFrom(row, "CPF")) },
    ];
  }

  function competenceText(value) {
    const [year, month] = value.split("-").map(Number);
    const months = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
    return `${months[month - 1]}/${year}`;
  }

  async function generateCajuMonthly() {
    if (!state.filtered.length) return;
    try {
      elements.loading.hidden = false;
      if (!window.MONTHLY_TEMPLATE_BASE64) throw new Error("O modelo mensal interno não foi carregado.");
      const templateBytes = base64Bytes(window.MONTHLY_TEMPLATE_BASE64);
      const zip = await JSZip.loadAsync(templateBytes);
      const workbook = XLSX.read(templateBytes, { type: "array", cellDates: true, cellFormula: true });
      const sourceSheet = workbook.Sheets["2-Funcionarios"] || workbook.Sheets[workbook.SheetNames[0]];
      const sheetPath = "xl/worksheets/sheet1.xml";
      let xml = await zip.file(sheetPath).async("string");
      const rowMatches = [...xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g)];
      const sourceRows = new Map(rowMatches.map((match) => [Number(match[1]), match[0]]));
      const defaultRowXml = sourceRows.get(8);
      if (!defaultRowXml) throw new Error("O modelo mensal da Caju não pôde ser lido.");

      const templateEmployees = new Map();
      const sourceRange = XLSX.utils.decode_range(sourceSheet["!ref"] || "A1:AB1");
      for (let rowNumber = 8; rowNumber <= sourceRange.e.r + 1; rowNumber += 1) {
        const key = employeeKey(sourceSheet[`C${rowNumber}`]?.v);
        if (!key) continue;
        templateEmployees.set(key, {
          rowNumber,
          filial: text(sourceSheet[`A${rowNumber}`]?.v),
          rowXml: sourceRows.get(rowNumber) || defaultRowXml,
        });
      }

      const fixedRows = rowMatches.filter((match) => Number(match[1]) <= 7).map((match) => match[0]).join("");
      const dataRows = state.filtered.map((row, index) => {
        const targetRow = index + 8;
        const key = employeeKey(valueFrom(row, "Matricula", "Matrícula"));
        const templateRow = templateEmployees.get(key);
        const rowXml = templateRow?.rowXml || defaultRowXml;
        const styles = stylesFromTemplateRow(rowXml);
        const base = paymentBaseCells(row, templateRow);
        const benefits = templateRow
          ? benefitCells(sourceSheet, templateRow.rowNumber, targetRow)
          : newEmployeeBenefits(targetRow);
        const values = [...base, ...benefits];
        const cells = values.map((cell, column) => monthlyCellXml(column, targetRow, styles.get(column), cell)).join("");
        return `<row r="${targetRow}" spans="1:28">${cells}</row>`;
      }).join("");
      const lastRow = state.filtered.length + 7;

      xml = xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${fixedRows}${dataRows}</sheetData>`);
      xml = xml.replace(/<dimension ref="[^"]+"\/>/, `<dimension ref="A1:AB${lastRow}"/>`);
      xml = xml.replace(/(<autoFilter\b[^>]*\bref=")A7:AB\d+("[^>]*\/>)/, `$1A7:AB${lastRow}$2`);
      xml = xml.replace(/([A-Z]+)8:\1\d+/g, (match, column) => `${column}8:${column}${lastRow}`);
      zip.file(sheetPath, xml);

      const competence = competenceText(elements.date.value);
      for (const path of ["xl/sharedStrings.xml", "xl/workbook.xml"]) {
        const file = zip.file(path);
        if (!file) continue;
        const content = await file.async("string");
        zip.file(path, content.replace(/SETEMBRO\/2026/g, competence));
      }
      const workbookFile = zip.file("xl/workbook.xml");
      if (workbookFile) {
        let workbookXml = await workbookFile.async("string");
        workbookXml = workbookXml.replace(/<calcPr\b[^>]*\/>/, '<calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>');
        zip.file("xl/workbook.xml", workbookXml);
      }

      const blob = await zip.generateAsync({ type: "blob", compression: "STORE" });
      const [year, month] = elements.date.value.split("-");
      window.__lastGeneratedBlob = blob;
      window.__lastGeneratedName = `CAJU HC4_${month}-${year}_Mensal.xlsx`;
      document.querySelector("#last-generated-download")?.remove();
      const link = document.createElement("a");
      link.id = "last-generated-download";
      link.hidden = true;
      link.href = URL.createObjectURL(blob);
      link.download = `CAJU HC4_${month}-${year}_Mensal.xlsx`;
      document.body.appendChild(link);
      link.click();
      const objectUrl = link.href;
      setTimeout(() => {
        URL.revokeObjectURL(objectUrl);
        link.remove();
      }, 60000);
      setStatus(`Arquivo mensal gerado com ${state.filtered.length} pessoas.`, "success");
    } catch (error) {
      setStatus(error.message || "Não foi possível gerar o pagamento mensal da Caju.", "error");
    } finally {
      elements.loading.hidden = true;
    }
  }

  function generateOutput() {
    if (state.mode === "caju-card") return generateCajuCsv();
    if (state.mode === "caju-payment") return generateCajuMonthly();
    return generateWorkbook();
  }

  elements.input.addEventListener("change", () => loadFile(elements.input.files[0]));
  elements.remove.addEventListener("click", clearFile);
  elements.date.addEventListener("change", updateFilteredRows);
  elements.generate.addEventListener("click", generateOutput);
  elements.modePlan.addEventListener("click", () => setMode("plan"));
  elements.modeCaju.addEventListener("click", () => setMode("caju-card"));
  elements.modeCajuPayment.addEventListener("click", () => setMode("caju-payment"));
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
