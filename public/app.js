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
  const pdf = document.getElementById("btnPdf");
  if (!input) return;
  const val = input.value.trim();
  if (hidden) hidden.value = val;
  if (pdf) {
    const u = new URL(pdf.href, window.location.origin);
    u.searchParams.set("data", val);
    pdf.href = u.pathname + u.search;
  }
}

const dataInput = document.getElementById("dataPedido");
if (dataInput) {
  dataInput.addEventListener("change", syncDataPedido);
  dataInput.addEventListener("input", syncDataPedido);
  syncDataPedido();
}
