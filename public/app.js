/* —— Barra de carregamento Neon —— */
const NeonLoad = (() => {
  let depth = 0;
  let hideTimer = null;

  function el() {
    return document.getElementById("neonLoader");
  }

  function labelEl() {
    return document.getElementById("neonLoaderLabel");
  }

  function show(message) {
    const node = el();
    if (!node) return;
    if (hideTimer) {
      clearTimeout(hideTimer);
      hideTimer = null;
    }
    depth += 1;
    if (message && labelEl()) labelEl().textContent = message;
    node.hidden = false;
    node.classList.remove("is-done");
    node.classList.add("is-on");
    document.documentElement.classList.add("neon-loading");
    try {
      sessionStorage.setItem("neon_loading", "1");
    } catch (e) {}
  }

  function hide(force) {
    depth = force ? 0 : Math.max(0, depth - 1);
    if (depth > 0) return;
    const node = el();
    if (!node) {
      document.documentElement.classList.remove("neon-loading");
      try {
        sessionStorage.removeItem("neon_loading");
      } catch (e) {}
      return;
    }
    node.classList.add("is-done");
    hideTimer = setTimeout(() => {
      node.classList.remove("is-on", "is-done");
      node.hidden = true;
      document.documentElement.classList.remove("neon-loading");
      try {
        sessionStorage.removeItem("neon_loading");
      } catch (e) {}
      hideTimer = null;
    }, 220);
  }

  function wrap(promise, message) {
    show(message || "Carregando Neon…");
    return Promise.resolve(promise).finally(() => hide());
  }

  // Página terminou de carregar → some a barra da navegação anterior
  function ready() {
    depth = 0;
    hide(true);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", ready);
  } else {
    ready();
  }
  window.addEventListener("pageshow", (ev) => {
    if (ev.persisted) ready();
  });

  return { show, hide, wrap };
})();

async function neonFetch(url, options, message) {
  return NeonLoad.wrap(fetch(url, options), message || "Salvando no Neon…");
}

function isInternalNav(href) {
  if (!href || href.startsWith("#") || href.startsWith("javascript:")) return false;
  try {
    const u = new URL(href, window.location.origin);
    return u.origin === window.location.origin;
  } catch (e) {
    return false;
  }
}

// Links e formulários que batem no servidor/Neon
document.addEventListener("click", (e) => {
  const a = e.target.closest("a[href]");
  if (!a || e.defaultPrevented || e.button !== 0 || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) {
    return;
  }
  if (a.hasAttribute("download") || a.target === "_blank") return;
  if (!isInternalNav(a.getAttribute("href"))) return;
  const path = a.pathname || "";
  // Export PDF/Excel também passa pelo servidor
  NeonLoad.show(
    path.includes("historico")
      ? "Carregando histórico…"
      : path.includes("pdf") || path.includes("export")
        ? "Gerando arquivo…"
        : "Carregando Neon…"
  );
});

document.addEventListener("submit", (e) => {
  const form = e.target;
  if (!(form instanceof HTMLFormElement)) return;
  if (form.dataset.noLoader === "1") return;
  const action = (form.getAttribute("action") || "").toLowerCase();
  let msg = "Carregando Neon…";
  if (action.includes("estoque") || form.querySelector('input[type="file"]')) {
    msg = "Enviando estoque ao Neon…";
  } else if (action.includes("pedido") || form.id === "formSalvar") {
    msg = "Salvando pedido no Neon…";
  } else if (action.includes("login") || form.classList.contains("login-form")) {
    msg = "Entrando…";
  }
  NeonLoad.show(msg);
});

async function salvarManual(codigo, valor, unidade) {
  const res = await neonFetch(
    "/api/manual",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        codigo,
        valor: valor === "" ? null : Number(valor),
        unidade: unidade || undefined,
      }),
    },
    "Salvando ajuste no Neon…"
  );
  let data = {};
  try {
    data = await res.json();
  } catch (e) {
    data = { ok: false };
  }

  const row = document.querySelector(
    `.item-row[data-codigo="${CSS.escape(String(codigo))}"]`
  );
  if (!row) {
    NeonLoad.show("Atualizando tela…");
    window.location.reload();
    return;
  }

  const pedir = Boolean(data.pedir);
  const quanto = data.quanto_pedir;
  const qtdeCell = row.querySelector(".quanto");
  const pill = row.querySelector(".pill");

  if (qtdeCell) {
    if (quanto != null && Number.isFinite(Number(quanto))) {
      qtdeCell.textContent = String(Math.trunc(Number(quanto)));
      qtdeCell.classList.toggle("col-qtde-pedir", pedir);
    } else {
      qtdeCell.innerHTML = '<span class="muted">—</span>';
      qtdeCell.classList.remove("col-qtde-pedir");
    }
  }
  if (pill) {
    if (data.estoque_aghu == null && data.manual == null) {
      pill.className = "pill neutro";
      pill.textContent = "?";
    } else if (pedir) {
      pill.className = "pill sim";
      pill.textContent = "Sim";
    } else {
      pill.className = "pill nao";
      pill.textContent = "Não";
    }
  }

  // Atualiza pedido do dia em segundo plano (não trava a linha)
  const dataPedido =
    document.getElementById("dataPedido")?.value.trim() ||
    document.getElementById("dataPedidoHidden")?.value ||
    "";
  neonFetch(
    "/api/autosave-pedido",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        unidade: unidade || row.dataset.unidade,
        data_pedido: dataPedido,
      }),
    },
    "Atualizando pedido do dia…"
  ).catch(() => {});
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
      NeonLoad.show("Carregando unidade…");
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
      const res = await neonFetch(
        "/api/parametros",
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        },
        "Salvando parâmetros no Neon…"
      );
      const data = await res.json();
      if (!res.ok || !data.ok) {
        alert(data.erro || "Não foi possível salvar.");
        saveBtn.disabled = false;
        return;
      }
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
    const res = await neonFetch(
      "/api/autosave-pedido",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ unidade, data_pedido: dataPedido }),
      },
      "Gravando pedido do dia no Neon…"
    );
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
