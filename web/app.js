const LOG_PATTERN = /^(?<clock>\d{2}(?::|h)\d{2}h?)\s*[-–]\s*Abertura da tela de Despacho\s*[-–]\s*(?<company>[A-Z]{3})\s*[-–]\s*EXCEDIDO EM:\s*(?<percent>\d+)%/i;

const state = {
  records: [],
  lastClockMinutes: null,
};

function formatClock(date) {
  return date.toLocaleTimeString("pt-BR", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
}

function adjustedIso(date) {
  const clone = new Date(date.getTime());
  clone.setHours(clone.getHours() + 3);
  const iso = clone.toISOString();
  return iso.slice(0, 16) + "Z";
}

function normalizeClock(rawClock) {
  let clock = rawClock.trim().toLowerCase();
  if (clock.endsWith("h")) clock = clock.slice(0, -1);
  clock = clock.replace("h", ":");

  if (!/^\d{2}:\d{2}$/.test(clock)) {
    throw new Error(`Horário inválido: "${rawClock}"`);
  }

  return clock;
}

function parseLine(line, shiftDate, previousClock) {
  const match = line.match(LOG_PATTERN);
  if (!match) {
    throw new Error(`Linha não corresponde ao padrão esperado: "${line}"`);
  }

  const normalizedClock = normalizeClock(match.groups.clock);
  const [hours, minutes] = normalizedClock.split(":").map(Number);
  const company = match.groups.company.toUpperCase();
  const percent = Number(match.groups.percent);

  const currentClockMinutes = hours * 60 + minutes;
  const dayOffset = previousClock !== null && currentClockMinutes < previousClock ? 1 : 0;

  const date = new Date(shiftDate);
  date.setDate(date.getDate() + dayOffset);
  date.setHours(hours, minutes, 0, 0);

  return { timestamp: date, company, percent, currentClockMinutes };
}

function addRecord(record) {
  state.records.push(record);
  state.lastClockMinutes = record.currentClockMinutes;
  render();
}

function render() {
  const total = state.records.length;
  const totalElem = document.getElementById("total-count");
  const avgElem = document.getElementById("average");
  const listElem = document.getElementById("record-list");
  const subtitle = document.getElementById("list-subtitle");

  totalElem.textContent = total;
  const avg = total
    ? (state.records.reduce((acc, r) => acc + r.percent, 0) / total).toFixed(1)
    : 0;
  avgElem.textContent = `${avg}%`;

  listElem.innerHTML = "";

  if (!total) {
    subtitle.textContent = "Nenhum registro adicionado";
    return;
  }

  subtitle.textContent = `${total} registro${total > 1 ? "s" : ""} adicionado${
    total > 1 ? "s" : ""
  }`;

  state.records.forEach((record) => {
    const item = document.createElement("div");
    item.className = "record";

    const clock = document.createElement("div");
    clock.className = "clock";
    clock.textContent = formatClock(record.timestamp);

    const description = document.createElement("div");
    description.className = "description";
    description.textContent = `Abertura da tela de Despacho – ${record.company}`;

    const details = document.createElement("div");
    details.className = "details";
    details.textContent = `TEMPO EXCEDIDO EM: ${record.percent}% | ${
      record.timestamp.toLocaleDateString()
    }`;

    const percent = document.createElement("div");
    percent.className = "percent";
    percent.textContent = `${record.percent}%`;

    item.append(clock, description, percent);
    item.appendChild(details);
    listElem.appendChild(item);
  });
}

function ensureShiftDate() {
  const value = document.getElementById("shift-date").value;
  if (!value) {
    throw new Error("Informe a data do turno antes de adicionar registros.");
  }
  return new Date(value + "T00:00:00");
}

function handleFormSubmit(event) {
  event.preventDefault();
  const errorElem = document.getElementById("error");
  errorElem.textContent = "";

  try {
    const shiftDate = ensureShiftDate();
    const time = document.getElementById("time").value;
    const company = document.getElementById("company").value.trim().toUpperCase();
    const percent = document.getElementById("percent").value.trim();

    if (!time || !company || !percent) {
      throw new Error("Preencha todos os campos do registro.");
    }

    const syntheticLine = `${time} - Abertura da tela de Despacho - ${company} - EXCEDIDO EM: ${percent}%`;
    const record = parseLine(syntheticLine, shiftDate, state.lastClockMinutes ?? null);
    addRecord(record);

    event.target.reset();
    state.lastClockMinutes = record.currentClockMinutes;
  } catch (err) {
    errorElem.textContent = err.message;
  }
}

function handleBulkAdd() {
  const errorElem = document.getElementById("error");
  errorElem.textContent = "";

  try {
    const shiftDate = ensureShiftDate();
    const raw = document.getElementById("bulk-input").value.trim();
    if (!raw) {
      throw new Error("Cole ao menos uma linha para adicionar em massa.");
    }

    const lines = raw.split(/\n+/);
    let lastClock = state.lastClockMinutes;

    lines.forEach((line) => {
      const record = parseLine(line.trim(), shiftDate, lastClock ?? null);
      addRecord(record);
      lastClock = record.currentClockMinutes;
    });

    document.getElementById("bulk-input").value = "";
  } catch (err) {
    errorElem.textContent = err.message;
  }
}

function exportFile(kind) {
  if (!state.records.length) return alert("Nenhum registro para exportar.");

  const lines = state.records.map((r) => {
    const iso = adjustedIso(r.timestamp);
    return `${iso};${r.company};${r.percent}`;
  });

  const header = "data_hora_utc;empresa;excedido_%";
  const payload = [header, ...lines].join(kind === "csv" ? "\n" : "\r\n");
  const blob = new Blob([payload], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `despacho.${kind}`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function main() {
  document.getElementById("record-form").addEventListener("submit", handleFormSubmit);
  document.getElementById("bulk-add").addEventListener("click", handleBulkAdd);
  document.getElementById("export-csv").addEventListener("click", () => exportFile("csv"));
  document.getElementById("export-txt").addEventListener("click", () => exportFile("txt"));
}

main();
