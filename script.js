const form = document.getElementById("receipt-form");
const fields = {
  buyer: document.getElementById("preview-buyer"),
  status: document.getElementById("preview-status"),
  quantity: document.getElementById("preview-quantity"),
  total: document.getElementById("preview-total"),
};
const previewNumbers = document.getElementById("preview-numbers");
const totalInput = document.getElementById("total");

const defaults = {
  buyer: "---",
  status: "---",
  quantity: "---",
  total: "---",
};

function setText(field, value) {
  fields[field].textContent = value || defaults[field];
}

function formatTotal(value) {
  if (!value) return "";
  return String(value)
    .replace(/\./g, "TEMP")
    .replace(/,/g, ".")
    .replace(/TEMP/g, ",");
}

function parseBrazilianNumber(value) {
  let text = String(value).trim();
  if (/^R\$/i.test(text)) {
    text = text.replace(/^R\$/i, "").trim();
  }
  text = text.replace(/\s/g, "");

  if (/^\d{1,3}(\.\d{3})*(,\d{1,2})?$/.test(text)) {
    text = text.replace(/\./g, "").replace(/,/g, ".");
  } else if (/^\d+(,\d{1,2})?$/.test(text)) {
    text = text.replace(/,/g, ".");
  } else if (/^\d+(\.\d{1,2})?$/.test(text)) {
    // no change, dot is decimal separator
  } else {
    const digits = text.replace(/[^0-9]/g, "");
    if (digits === "") return null;
    if (digits.length <= 2) {
      text = `0.${digits.padStart(2, "0")}`;
    } else {
      text = `${digits.slice(0, -2)}.${digits.slice(-2)}`;
    }
  }

  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

function generateRandomNumbers(count) {
  const total = Number.parseInt(count, 10);
  if (!Number.isFinite(total) || total <= 0) {
    return [];
  }
  const results = new Set();
  while (results.size < total) {
    const value = Math.floor(1000000 + Math.random() * 9000000);
    results.add(String(value));
  }
  return Array.from(results);
}

function updatePreview() {
  const data = new FormData(form);
  setText("buyer", data.get("buyer")?.trim());
  setText("status", data.get("status"));
  setText("quantity", data.get("quantity")?.trim());
  setText("total", formatTotal(data.get("total")));

  const numbersRaw = data.get("numbers") || "";
  const numbers = numbersRaw
    .split(/[\n,;]+/)
    .map((item) => item.trim())
    .filter(Boolean);
  const maxItems = 60;
  const shownNumbers = numbers.slice(0, maxItems);
  previewNumbers.innerHTML = "";
  if (shownNumbers.length === 0) {
    const empty = document.createElement("div");
    empty.className = "muted";
    empty.textContent = "Nenhum número informado.";
    previewNumbers.appendChild(empty);
  } else {
    shownNumbers.forEach((value) => {
      const pill = document.createElement("div");
      pill.className = "pill";
      pill.textContent = value;
      previewNumbers.appendChild(pill);
    });
  }
}

document.getElementById("generate-numbers").addEventListener("click", () => {
  const quantity = document.getElementById("quantity").value;
  const randomNumbers = generateRandomNumbers(quantity);
  if (randomNumbers.length > 0) {
    document.getElementById("numbers").value = randomNumbers.join(", ");
  }
  updatePreview();
});

if (totalInput) {
  totalInput.addEventListener("input", () => {
    totalInput.value = formatTotal(totalInput.value);
    updatePreview();
  });
}

form.addEventListener("input", () => {
  updatePreview();
});

document.getElementById("process-excel").addEventListener("click", () => {
  const fileInput = document.getElementById("excel-file");
  const file = fileInput.files[0];
  if (!file) {
    alert("Por favor, selecione um arquivo Excel.");
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" });

    if (jsonData.length < 2) {
      alert("O arquivo Excel deve ter pelo menos uma linha de cabeçalho e uma linha de dados.");
      return;
    }

    const headers = jsonData[0];
    const rows = jsonData.slice(1);

    const normalizeHeader = (value) =>
      String(value || "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/\p{Diacritic}/gu, "")
        .trim();

    const colMap = {};
    headers.forEach((header, index) => {
      const normalized = normalizeHeader(header);
      if (["comprador", "cliente", "nome", "nome do cliente", "buyer", "nome cliente"].some((key) => normalized.includes(key))) {
        colMap.buyer = index;
      }
      if (["situacao", "situacao", "situacao", "status", "situacao do pagamento", "situacao pagamento"].some((key) => normalized.includes(key))) {
        colMap.status = index;
      }
      if (["quantidade", "qtd", "qtde", "quantidade de itens", "quantidade itens"].some((key) => normalized.includes(key))) {
        colMap.quantity = index;
      }
      if (["total", "valor", "valor total", "montante", "amount", "price"].some((key) => normalized.includes(key))) {
        colMap.total = index;
      }
      if (["titulo", "titulos", "títulos", "títulos", "numeros", "números", "numbers", "boletos", "codigos"].some((key) => normalized.includes(key))) {
        colMap.numbers = index;
      }
    });

    // Fallback por posição quando não houver cabeçalho reconhecido
    if (rows.length > 0) {
      if (colMap.buyer === undefined) colMap.buyer = 0;
      if (colMap.status === undefined) colMap.status = 1;
      if (colMap.quantity === undefined) colMap.quantity = 2;
      if (colMap.total === undefined) colMap.total = 3;
      if (colMap.numbers === undefined) colMap.numbers = 4;
    }

    const batchPreviews = document.getElementById("batch-previews");
    batchPreviews.innerHTML = ""; // Limpar prévias existentes

    if (rows.length > 0) {
      const firstRow = rows[0];
      document.getElementById("buyer").value = firstRow[colMap.buyer] || "";
      document.getElementById("quantity").value = firstRow[colMap.quantity] || "";
      document.getElementById("status").value = firstRow[colMap.status] || "Pago";
      document.getElementById("total").value = firstRow[colMap.total] || "";
      document.getElementById("numbers").value = firstRow[colMap.numbers] || "";
      updatePreview(); // Atualizar a pré-visualização do formulário
    }

    rows.forEach((row, index) => {
      const buyer = row[colMap.buyer] || "";
      const status = row[colMap.status] || "Pago";
      const quantity = row[colMap.quantity] || "";
      const total = row[colMap.total] || "";
      const numbers = row[colMap.numbers] || "";

      const previewDiv = document.createElement("div");
      previewDiv.className = "preview";
      previewDiv.innerHTML = `
        <h3>Comprovante ${index + 1}</h3>
        <ul class="preview-list">
          <li><span class="label">Comprador:</span> <span>${buyer}</span></li>
          <li><span class="label">Situação:</span> <span>${status}</span></li>
          <li><span class="label">Quantidade:</span> <span>${quantity}</span></li>
          <li><span class="label">Total:</span> <span>${formatTotal(total)}</span></li>
        </ul>
        <div class="numbers-title">Títulos:</div>
        <div class="numbers-grid">${generateNumbersHTML(numbers)}</div>
      `;
      batchPreviews.appendChild(previewDiv);
    });
  };
  reader.readAsArrayBuffer(file);
});

function generateNumbersHTML(numbersRaw) {
  const numbers = numbersRaw
    .toString()
    .split(/[\n,;]+/)
    .map((item) => item.trim())
    .filter(Boolean);
  const maxItems = 60;
  const shownNumbers = numbers.slice(0, maxItems);
  if (shownNumbers.length === 0) {
    return '<div class="muted">Nenhum número informado.</div>';
  } else {
    return shownNumbers.map((value) => `<div class="pill">${value}</div>`).join("");
  }
}

document.getElementById("print-all").addEventListener("click", () => {
  document.body.classList.add("printing-batch");
  window.print();
  setTimeout(() => {
    document.body.classList.remove("printing-batch");
  }, 1000);
});

updatePreview();
