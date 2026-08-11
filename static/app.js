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
  // Sem reload completo — só atualiza a linha na tela
  const row = document.querySelector(`.item-row[data-codigo="${CSS.escape(String(codigo))}"]`);
  if (!row) {
    window.location.reload();
    return;
  }
  const qtdeCell = row.querySelector(".quanto");
  if (valor === "") {
    // volta ao cálculo do servidor no próximo refresh; marca visualmente
    return;
  }
  const n = Number(valor);
  if (qtdeCell && Number.isFinite(n)) {
    qtdeCell.textContent = String(Math.trunc(n));
    qtdeCell.classList.add("col-qtde-pedir");
  }
  const pill = row.querySelector(".pill");
  if (pill && Number.isFinite(n) && n > 0) {
    pill.className = "pill sim";
    pill.textContent = "Sim";
  }
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

function setRowEditing(row, editing) {
  row.classList.toggle("is-editing", editing);
  row.querySelectorAll(".param-view").forEach((el) => {
    el.hidden = editing;
  });
  row.querySelectorAll(".param-input").forEach((el) => {
    el.hidden = !editing;
  });
  const editBtn = row.querySelector(".btn-edit-param");
  const saveBtn = row.querySelector(".btn-save-param");
  const cancelBtn = row.querySelector(".btn-cancel-param");
  if (editBtn) editBtn.hidden = editing;
  if (saveBtn) saveBtn.hidden = !editing;
  if (cancelBtn) cancelBtn.hidden = !editing;
}

document.querySelectorAll(".item-row").forEach((row) => {
  const editBtn = row.querySelector(".btn-edit-param");
  const saveBtn = row.querySelector(".btn-save-param");
  const cancelBtn = row.querySelector(".btn-cancel-param");
  if (!editBtn) return;

  editBtn.addEventListener("click", () => setRowEditing(row, true));

  cancelBtn?.addEventListener("click", () => {
    row.querySelectorAll(".param-cell").forEach((cell) => {
      const view = cell.querySelector(".param-view");
      const input = cell.querySelector(".param-input");
      if (view && input) input.value = view.textContent.trim();
    });
    setRowEditing(row, false);
  });

  saveBtn?.addEventListener("click", async () => {
    const payload = {
      unidade: row.dataset.unidade,
      codigo: row.dataset.codigo,
    };
    row.querySelectorAll(".param-cell").forEach((cell) => {
      const field = cell.dataset.field;
      const input = cell.querySelector(".param-input");
      if (field && input) payload[field] = Number(input.value);
    });
    saveBtn.disabled = true;
    try {
      const res = await fetch("/api/parametros", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok || !data.ok) {
        alert(data.erro || "Não foi possível salvar.");
        saveBtn.disabled = false;
        return;
      }
      // Atualiza a linha sem recarregar a página inteira
      row.querySelectorAll(".param-cell").forEach((cell) => {
        const view = cell.querySelector(".param-view");
        const input = cell.querySelector(".param-input");
        if (view && input) view.textContent = String(Math.trunc(Number(input.value) || 0));
      });
      setRowEditing(row, false);
      saveBtn.disabled = false;
    } catch (err) {
      alert("Erro ao salvar parâmetros.");
      saveBtn.disabled = false;
    }
  });
});

async function gravarPedidoDiaEmBackground() {
  const params = new URLSearchParams(window.location.search);
  if (params.get("gravar_pedido") !== "1") return;

  const status = document.getElementById("pedidoDiaStatus");
  if (status) status.textContent = "Gravando pedido do dia no histórico…";

  const unidade =
    params.get("unidade") ||
    document.querySelector(".item-row")?.dataset.unidade ||
    "ufcd";
  const dataPedido =
    document.getElementById("dataPedido")?.value.trim() ||
    document.getElementById("dataPedidoHidden")?.value ||
    "";

  try {
    const res = await fetch("/api/autosave-pedido", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ unidade, data_pedido: dataPedido }),
    });
    const data = await res.json();
    if (status) {
      if (data.salvo) {
        status.textContent = `Pedido do dia salvo #${data.pedido_id} · ${data.itens} itens`;
      } else if (data.ok) {
        status.textContent = "Estoque ok — nada a pedir para gravar no histórico";
      } else {
        status.textContent = data.erro || "Falha ao gravar pedido do dia";
      }
    }
  } catch (err) {
    if (status) status.textContent = "Estoque ok — histórico pode ser salvo pelo botão";
  }

  params.delete("gravar_pedido");
  const qs = params.toString();
  const clean = `${window.location.pathname}${qs ? `?${qs}` : ""}`;
  window.history.replaceState({}, "", clean);
}

gravarPedidoDiaEmBackground();
