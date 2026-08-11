async function salvarManual(codigo, valor, unidade) {
  await fetch("/api/manual", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      codigo,
      valor: valor === "" ? null : Number(valor),
      unidade: unidade || undefined,
    }),
  });
  window.location.href = new URL(window.location.href).toString();
}

document.querySelectorAll(".manual-input").forEach((input) => {
  input.addEventListener("change", () => {
    salvarManual(input.dataset.codigo, input.value.trim(), input.dataset.unidade);
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
  document.querySelectorAll(".export-link").forEach((link) => {
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

document.querySelectorAll(".unit-toggle").forEach((btn) => {
  btn.addEventListener("click", () => {
    const unit = btn.dataset.unit;
    const params = new URLSearchParams(window.location.search);
    const current = params.get("unidade") || "ufcd";
    if (unit && unit !== current) {
      params.set("unidade", unit);
      params.set("aba", "todos");
      params.delete("q");
      window.location.href = `${window.location.pathname}?${params.toString()}`;
      return;
    }
    const panel = document.querySelector(`[data-unit-panel="${unit}"]`);
    const open = btn.classList.toggle("is-open");
    btn.setAttribute("aria-expanded", open ? "true" : "false");
    if (panel) panel.classList.toggle("is-open", open);
  });
});
