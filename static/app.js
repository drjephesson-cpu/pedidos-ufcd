async function salvarManual(codigo, valor) {
  await fetch("/api/manual", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ codigo, valor: valor === "" ? null : Number(valor) }),
  });
  // recarrega para recalcular "Quanto Pedir?"
  const url = new URL(window.location.href);
  window.location.href = url.toString();
}

document.querySelectorAll(".manual-input").forEach((input) => {
  let timer;
  input.addEventListener("change", () => {
    clearTimeout(timer);
    salvarManual(input.dataset.codigo, input.value.trim());
  });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      input.blur();
    }
  });
});
