const sumEl = document.getElementById("sum");
const avgEl = document.getElementById("avg");
const countEl = document.getElementById("count");

const NUMBER_PATTERN = /[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?/g;

function parseNumbersFromText(text) {
  const matches = text.match(NUMBER_PATTERN) || [];
  return matches
    .map((raw) => raw.replaceAll(",", ""))
    .map((normalized) => Number(normalized))
    .filter((value) => Number.isFinite(value));
}

function formatNumber(value) {
  return Number.isFinite(value)
    ? value.toLocaleString("zh-CN", { maximumFractionDigits: 8 })
    : "0";
}

function renderStats(numbers) {
  const count = numbers.length;
  const sum = numbers.reduce((acc, n) => acc + n, 0);
  const avg = count > 0 ? sum / count : 0;

  sumEl.textContent = formatNumber(sum);
  avgEl.textContent = formatNumber(avg);
  countEl.textContent = String(count);
}

async function recalcFromSelection() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const numbers = parseNumbersFromText(selection.text || "");
    renderStats(numbers);
  });
}

Office.onReady(() => {
  recalcFromSelection().catch(() => renderStats([]));

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => {
      recalcFromSelection().catch(() => renderStats([]));
    }
  );
});
