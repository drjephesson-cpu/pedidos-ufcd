async function salvarManual(codigo, valor) {
  await fetch("/api/manual", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ codigo, valor: valor === "" ? null : Number(valor) }),
  });
  const url = new URL(window.location.href);
  window.location.href = url.toString();
}

document.querySelectorAll(".manual-input").forEach((input) => {
  input.addEventListener("change", () => {
    salvarManual(input.dataset.codigo, input.value.trim());
  });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      input.blur();
    }
  });
});

function syncDataPedido() {
  const input = document.getElementById("dataPedido");
  const hidden = document.getElementById("dataPedidoHidden");
  if (!input) return;
  const val = input.value.trim();
  if (hidden) hidden.value = val;
  document.querySelectorAll(".pdf-link").forEach((link) => {
    const u = new URL(link.getAttribute("href"), window.location.origin);
    u.searchParams.set("data", val);
    link.href = u.pathname + u.search;
  });
}

const dataInput = document.getElementById("dataPedido");
if (dataInput) {
  dataInput.addEventListener("change", syncDataPedido);
  dataInput.addEventListener("input", syncDataPedido);
  syncDataPedido();
}

const toggle = document.getElementById("sidebarToggle");
if (toggle) {
  toggle.addEventListener("click", () => {
    document.body.classList.toggle("sidebar-open");
  });
}

const pdfToggle = document.getElementById("pdfToggle");
const pdfPanel = document.querySelector("#pdfMenu .pdf-menu-panel");
if (pdfToggle && pdfPanel) {
  pdfToggle.addEventListener("click", (e) => {
    e.stopPropagation();
    pdfPanel.hidden = !pdfPanel.hidden;
  });
  document.addEventListener("click", () => {
    pdfPanel.hidden = true;
  });
  pdfPanel.addEventListener("click", (e) => e.stopPropagation());
}
