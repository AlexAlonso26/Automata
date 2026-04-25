import { useState } from "react";
import * as XLSX from "xlsx";
import "../styles.css";

const normalizeHeader = (value) =>
  String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .trim();

const formatTotal = (value) => {
  if (value === undefined || value === null) return "";
  return String(value)
    .replace(/\./g, "TEMP")
    .replace(/,/g, ".")
    .replace(/TEMP/g, ",");
};

const parseNumbers = (numbersRaw) => {
  const numbers = String(numbersRaw || "")
    .split(/[\n,;]+/)
    .map((item) => item.trim())
    .filter(Boolean);
  return numbers;
};

const cleanQuantityValue = (value) => {
  if (value === undefined || value === null) return "";
  const normalizedText = String(value)
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .trim();
  const cleaned = normalizedText
    .replace(/titulos?/gi, "")
    .replace(/titulo/gi, "")
    .replace(/[^0-9]/g, "")
    .trim();
  return cleaned;
};

const findColumns = (headers) => {
  const colMap = {};
  headers.forEach((header, index) => {
    const key = normalizeHeader(header);
    if (["comprador", "cliente", "nome", "nome do cliente", "buyer", "nome cliente"].some((item) => key.includes(item))) {
      colMap.buyer = index;
    }
    if (["situacao", "situacao", "status", "situacao do pagamento", "situacao pagamento"].some((item) => key.includes(item))) {
      colMap.status = index;
    }
    if (["quantidade", "qtd", "qtde", "quantidade de itens", "quantidade itens"].some((item) => key.includes(item))) {
      colMap.quantity = index;
    }
    if (["total", "valor", "valor total", "montante", "amount", "price"].some((item) => key.includes(item))) {
      colMap.total = index;
    }
    if (["titulo", "titulos", "titulos", "numeros", "numeros", "numbers", "boletos", "codigos"].some((item) => key.includes(item))) {
      colMap.numbers = index;
    }
  });

  if (headers.length > 0) {
    if (colMap.buyer === undefined) colMap.buyer = 0;
    if (colMap.status === undefined) colMap.status = 1;
    if (colMap.quantity === undefined) colMap.quantity = 2;
    if (colMap.total === undefined) colMap.total = 3;
    if (colMap.numbers === undefined) colMap.numbers = 4;
  }

  return colMap;
};

export default function App() {
  const [buyer, setBuyer] = useState("");
  const [status, setStatus] = useState("Pago");
  const [quantity, setQuantity] = useState("");
  const [total, setTotal] = useState("");
  const [numbers, setNumbers] = useState("");
  const [batchPreviews, setBatchPreviews] = useState([]);

  const randomNumberTitles = (count) => {
    const total = Number.parseInt(count, 10);
    const maxCount = Math.min(Number.isFinite(total) && total > 0 ? total : 60, 60);
    const results = new Set();
    while (results.size < maxCount) {
      const value = Math.floor(1000000 + Math.random() * 9000000);
      results.add(String(value));
    }
    return Array.from(results);
  };

  const handleExcelProcess = (file) => {
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
      const colMap = findColumns(headers);

      if (rows.length > 0) {
        const firstRow = rows[0];
        const firstQuantity = cleanQuantityValue(firstRow[colMap.quantity]);
        setBuyer(firstRow[colMap.buyer] || "");
        setStatus("Pago");
        setQuantity(firstQuantity);
        setTotal(firstRow[colMap.total] || "");
        setNumbers(randomNumberTitles(firstQuantity || 60).join(", "));
      }

      const previews = rows.map((row, index) => {
        const quantityValue = cleanQuantityValue(row[colMap.quantity]);
        return {
          id: index + 1,
          buyer: row[colMap.buyer] || "",
          status: "Pago",
          quantity: quantityValue,
          total: row[colMap.total] || "",
          numbers: randomNumberTitles(quantityValue || 60).join(", "),
        };
      });
      setBatchPreviews(previews);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleGenerateNumbers = () => {
    const results = randomNumberTitles(quantity || 60);
    setNumbers(results.join(", "));
  };

  const handlePrintAll = () => {
    document.body.classList.add("printing-batch");
    window.print();
    setTimeout(() => {
      document.body.classList.remove("printing-batch");
    }, 1000);
  };

  const renderPills = (value) => {
    const items = parseNumbers(value).slice(0, 60);
    if (items.length === 0) {
      return <div className="muted">Nenhum número informado.</div>;
    }
    return items.map((item, index) => (
      <div className="pill" key={`${item}-${index}`}>{item}</div>
    ));
  };

  return (
    <div>
      <header>
        <h1>Gerador de Comprovante</h1>
        <p>Preencha os dados, gere o comprovante e edite se precisar.</p>
      </header>
      <main>
        <section className="card">
          <h2>Dados do comprovante</h2>
          <form id="receipt-form" onSubmit={(event) => event.preventDefault()}>
            <div>
              <label htmlFor="buyer">Comprador</label>
              <input
                id="buyer"
                name="buyer"
                placeholder="Nome do cliente"
                value={buyer}
                onChange={(event) => setBuyer(event.target.value)}
              />
            </div>
            <div>
              <label htmlFor="status">Situação</label>
              <select id="status" name="status" value={status} onChange={(event) => setStatus(event.target.value)}>
                <option value="Pago">Pago</option>
                <option value="Pendente">Pendente</option>
                <option value="Cancelado">Cancelado</option>
              </select>
            </div>
            <div>
              <label htmlFor="quantity">Quantidade</label>
              <input
                id="quantity"
                name="quantity"
                placeholder="Ex: 150"
                value={quantity}
                onChange={(event) => setQuantity(cleanQuantityValue(event.target.value))}
              />
            </div>
            <div>
              <label htmlFor="total">Total</label>
              <input
                id="total"
                name="total"
                placeholder="R$ 25,99"
                value={total}
                onChange={(event) => setTotal(formatTotal(event.target.value))}
              />
            </div>
            <div>
              <label htmlFor="numbers">Títulos</label>
              <textarea
                id="numbers"
                name="numbers"
                placeholder="0007902, 0068170, 0085461..."
                value={numbers}
                onChange={(event) => setNumbers(event.target.value)}
              />
              <div className="field-actions">
                <button type="button" className="ghost small" onClick={handleGenerateNumbers}>
                  Gerar números aleatórios
                </button>
              </div>
            </div>
            <div>
              <label htmlFor="excel-file">Arquivo Excel</label>
              <input type="file" id="excel-file" name="excel-file" accept=".xlsx,.xls" onChange={(event) => handleExcelProcess(event.target.files?.[0])} />
            </div>
            <div className="actions">
              <button type="button" className="ghost" onClick={() => window.print()}>
                Imprimir / Salvar PDF
              </button>
              <button type="button" className="ghost" onClick={handlePrintAll}>
                Imprimir Todos os Comprovantes
              </button>
            </div>
            <p className="muted">Dica: use quebra de linha ou vírgula para separar os números.</p>
          </form>
        </section>

        <section className="card preview-card">
          <h2>Pré-visualização</h2>
          <div id="batch-previews">
            <div className="preview" id="preview">
              <ul className="preview-list">
                <li>
                  <span className="label">Comprador:</span>
                  <span>{buyer || "---"}</span>
                </li>
                <li>
                  <span className="label">Situação:</span>
                  <span>{status || "---"}</span>
                </li>
                <li>
                  <span className="label">Quantidade:</span>
                  <span>{quantity || "---"}</span>
                </li>
                <li>
                  <span className="label">Total:</span>
                  <span>{formatTotal(total) || "---"}</span>
                </li>
              </ul>
              <div className="numbers-title">Títulos:</div>
              <div className="numbers-grid">{renderPills(numbers)}</div>
            </div>
            {batchPreviews.map((item) => (
              <div className="preview" key={item.id}>
                <h3>Comprovante {item.id}</h3>
                <ul className="preview-list">
                  <li>
                    <span className="label">Comprador:</span> <span>{item.buyer}</span>
                  </li>
                  <li>
                    <span className="label">Situação:</span> <span>{item.status}</span>
                  </li>
                  <li>
                    <span className="label">Quantidade:</span> <span>{item.quantity}</span>
                  </li>
                  <li>
                    <span className="label">Total:</span> <span>{formatTotal(item.total)}</span>
                  </li>
                </ul>
                <div className="numbers-title">Títulos:</div>
                <div className="numbers-grid">{renderPills(item.numbers)}</div>
              </div>
            ))}
          </div>
        </section>
      </main>
    </div>
  );
}
