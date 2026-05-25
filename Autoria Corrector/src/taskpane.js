(function () {
  const demoText = [
    "La ciudad dormia con un ojo abierto. Martina cruzo la avenida sin mirar atras, porque mirar atras era darle permiso al miedo.",
    "En el bolsillo llevaba una carta doblada tres veces, manchada de cafe, y en la carta habia una promesa que ya no sabia si era suya.",
    "El viento golpeaba los letreros y repetia su nombre. Martina penso en Julian, en la estacion, en la ultima vez que alguien le dijo quedate.",
    "Entonces corrio. Corrio como si la noche pudiera romperse. Corrio hasta que el silencio tuvo dientes."
  ].join(" ");

  const stopWords = new Set([
    "a", "al", "algo", "ante", "aqui", "asi", "cada", "como", "con", "de", "del", "desde",
    "el", "ella", "en", "entre", "era", "es", "esa", "ese", "esta", "este", "ha", "habia",
    "la", "las", "le", "lo", "los", "mas", "me", "mi", "no", "o", "para", "pero", "por",
    "que", "se", "si", "sin", "su", "sus", "te", "un", "una", "y", "ya"
  ]);

  const state = {
    text: demoText,
    initialWordCount: 0,
    currentWordCount: 0,
    editedBlocks: 0,
    activeSeconds: 0,
    lastActivity: Date.now(),
    timer: null
  };

  const $ = (selector) => document.querySelector(selector);
  const $$ = (selector) => Array.from(document.querySelectorAll(selector));

  function words(text) {
    return normalize(text).match(/[a-z0-9áéíóúñü]+/gi) || [];
  }

  function normalize(text) {
    return (text || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  function sentences(text) {
    return (text || "")
      .replace(/\s+/g, " ")
      .split(/(?<=[.!?;:])\s+/)
      .map((sentence) => sentence.trim())
      .filter(Boolean);
  }

  function getTopWords(text) {
    const counts = new Map();
    words(text).forEach((word) => {
      if (word.length > 3 && !stopWords.has(word)) {
        counts.set(word, (counts.get(word) || 0) + 1);
      }
    });
    return Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([word, count]) => ({ word, count }));
  }

  function analyzeText(text) {
    const sample = words(text).slice(0, 500).join(" ");
    const allSentences = sentences(text);
    const lengths = allSentences.map((sentence) => words(sentence).length).filter(Boolean);
    const average = lengths.length ? lengths.reduce((sum, value) => sum + value, 0) / lengths.length : 0;
    const variance = lengths.length
      ? lengths.reduce((sum, value) => sum + Math.pow(value - average, 2), 0) / lengths.length
      : 0;
    const rhythm = Math.min(100, Math.round(Math.sqrt(variance) * 9));
    const intensityWords = words(text).filter((word) =>
      ["miedo", "romperse", "corrio", "silencio", "dientes", "golpeaba", "promesa", "ultima"].includes(word)
    ).length;

    return {
      sample,
      sentences: allSentences,
      lengths,
      average,
      rhythm,
      topWords: getTopWords(text),
      intensity: intensityWords
    };
  }

  async function getDocumentText() {
    if (!window.Word || !window.Office || !Office.context || !Office.context.document) {
      return state.text;
    }

    return Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      return body.text || state.text;
    }).catch(() => state.text);
  }

  async function replaceSelection(text) {
    if (!window.Word) {
      $("#draftInput").value = text;
      markActivity();
      return;
    }

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
    markActivity();
  }

  async function protectSelection() {
    if (!window.Word) {
      $("#voiceState").textContent = "Fragmento protegido en demo";
      return;
    }

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const control = selection.insertContentControl();
      control.title = "AutorIA: No tocar";
      control.tag = "autoria-no-tocar";
      control.appearance = "Tags";
      await context.sync();
    });
    $("#voiceState").textContent = "Fragmento protegido";
  }

  function drawRhythm(lengths) {
    const canvas = $("#rhythmCanvas");
    const context = canvas.getContext("2d");
    const width = canvas.width;
    const height = canvas.height;
    const max = Math.max(18, ...lengths);

    context.clearRect(0, 0, width, height);
    context.fillStyle = "#241b0d";
    context.fillRect(0, 0, width, height);
    context.strokeStyle = "rgba(200, 148, 74, 0.2)";
    context.lineWidth = 1;

    for (let i = 0; i < 5; i += 1) {
      const y = 18 + i * 38;
      context.beginPath();
      context.moveTo(0, y);
      context.lineTo(width, y);
      context.stroke();
    }

    lengths.forEach((length, index) => {
      const barWidth = Math.max(12, width / Math.max(lengths.length, 1) - 7);
      const x = 14 + index * (barWidth + 7);
      const barHeight = Math.max(12, (length / max) * (height - 40));
      const y = height - barHeight - 18;
      context.fillStyle = length > max * 0.82 ? "#b64032" : "#c8944a";
      context.fillRect(x, y, barWidth, barHeight);
    });
  }

  function renderHeatmap(topWords) {
    const max = Math.max(1, ...topWords.map((item) => item.count));
    $("#heatmap").innerHTML = topWords.length
      ? topWords.map((item) => `
        <div class="heat-word">
          <span>${item.word}</span>
          <span class="bar"><span style="width:${Math.max(18, (item.count / max) * 100)}%"></span></span>
          <span>${item.count}</span>
        </div>
      `).join("")
      : "<p class=\"ghost\">No hay palabras repetidas suficientes para construir mapa de calor.</p>";
  }

  function buildSuggestions(analysis) {
    const suggestions = [];
    const duplicate = analysis.topWords.find((item) => item.count > 1);
    const flatZone = analysis.lengths.length > 3 && analysis.rhythm < 18;
    const longSentence = analysis.sentences.find((sentence) => words(sentence).length > Math.max(28, analysis.average * 1.8));

    if (duplicate) {
      suggestions.push({
        title: `Repeticion sensible: "${duplicate.word}"`,
        body: "Aparece varias veces cerca del nucleo del texto. Si es una marca de voz, mantenla; si busca pasar inadvertida, prueba una sustitucion contextual."
      });
    }

    if (flatZone) {
      suggestions.push({
        title: "Zona de ritmo plano",
        body: "Las oraciones tienen una longitud muy similar. Puedes romper la cadencia con una frase mas seca o una imagen mas extendida."
      });
    }

    if (longSentence) {
      suggestions.push({
        title: "Oracion de respiracion larga",
        body: longSentence.slice(0, 170) + (longSentence.length > 170 ? "..." : "")
      });
    }

    if (!suggestions.length) {
      suggestions.push({
        title: "Voz consistente",
        body: "No hay alertas fuertes. AutorIA sugiere observar el cierre de escena y conservar las decisiones de estilo que sostienen la tension."
      });
    }

    $("#suggestions").innerHTML = suggestions.map((item) => `
      <article class="suggestion">
        <strong>${item.title}</strong>
        <p>${item.body}</p>
        <div class="actions">
          <button class="mini accept" type="button">Aceptar</button>
          <button class="mini" type="button">Ignorar</button>
          <button class="mini reject" type="button">Descartar</button>
        </div>
      </article>
    `).join("");
  }

  function renderAnalysis(text) {
    const analysis = analyzeText(text);
    $("#avgSentence").textContent = analysis.average ? `${analysis.average.toFixed(1)} pal.` : "--";
    $("#rhythmScore").textContent = analysis.rhythm ? `${analysis.rhythm}/100` : "--";
    $("#toneLabel").textContent = analysis.intensity > 4 ? "Alta tension" : "Tension moderada";
    $("#voiceState").textContent = "Perfil construido";
    drawRhythm(analysis.lengths);
    renderHeatmap(analysis.topWords);
    buildSuggestions(analysis);
    updateGhost(analysis);
  }

  function updateGhost(analysis) {
    const input = $("#draftInput").value.trim();
    const source = input || analysis.sentences[analysis.sentences.length - 1] || "";
    const lastWord = words(source).slice(-1)[0] || "escena";
    $("#ghostText").textContent = `Texto fantasma: y entonces ${lastWord} vuelve a cargar la escena con una decision que todavia no quiere nombrar.`;
  }

  function rewriteDraft() {
    const text = $("#draftInput").value.trim();
    const base = text || "El personaje mira la habitacion y comprende que algo cambio.";
    const compact = base.replace(/\s+/g, " ");
    const softer = compact.replace(/\.$/, "") + ", aunque todavia no sabe como decirlo.";
    const sharper = compact.split(",")[0].replace(/\.$/, "") + ". Esta vez no retrocede.";
    const sensory = compact.replace(/\.$/, "") + "; el aire, de pronto, parece mas pesado.";

    $("#variants").innerHTML = [softer, sharper, sensory].map((variant, index) => `
      <article class="variant">
        <strong>Variante ${index + 1}</strong>
        <p>${variant}</p>
        <div class="actions">
          <button class="mini accept use-variant" data-text="${encodeURIComponent(variant)}" type="button">Usar</button>
          <button class="mini" type="button">Guardar</button>
        </div>
      </article>
    `).join("");
  }

  function switchTab(tabName) {
    $$(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.tab === tabName));
    $$(".panel").forEach((panel) => panel.classList.toggle("active", panel.id === tabName));
  }

  function countWords(text) {
    return words(text).length;
  }

  function markActivity() {
    state.lastActivity = Date.now();
    state.editedBlocks += 1;
    $("#editedBlocks").textContent = String(state.editedBlocks);
  }

  function startSessionTimer() {
    if (state.timer) {
      clearInterval(state.timer);
    }

    state.timer = setInterval(() => {
      const paused = Date.now() - state.lastActivity > 120000;
      if (!paused) {
        state.activeSeconds += 1;
      }
      const minutes = Math.floor(state.activeSeconds / 60).toString().padStart(2, "0");
      const seconds = (state.activeSeconds % 60).toString().padStart(2, "0");
      $("#activeTime").textContent = `${minutes}:${seconds}`;
    }, 1000);
  }

  async function scan() {
    $("#voiceState").textContent = "Analizando...";
    const text = await getDocumentText();
    state.text = text || demoText;
    state.currentWordCount = countWords(state.text);
    if (!state.initialWordCount) {
      state.initialWordCount = state.currentWordCount;
    }
    $("#sessionWords").textContent = String(Math.max(0, state.currentWordCount - state.initialWordCount));
    renderAnalysis(state.text);
    markActivity();
  }

  function closeSession() {
    const newWords = Math.max(0, state.currentWordCount - state.initialWordCount);
    const minutes = Math.floor(state.activeSeconds / 60);
    $("#sessionReport").textContent = `Hoy escribiste ${newWords} palabras nuevas, editaste ${state.editedBlocks} fragmentos y mantuviste ${minutes} minuto(s) de escritura activa.`;
  }

  function resetSession() {
    state.initialWordCount = state.currentWordCount;
    state.editedBlocks = 0;
    state.activeSeconds = 0;
    state.lastActivity = Date.now();
    $("#sessionWords").textContent = "0";
    $("#editedBlocks").textContent = "0";
    $("#activeTime").textContent = "00:00";
    $("#sessionReport").textContent = "Sesion reiniciada. AutorIA volvera a medir desde este punto.";
  }

  function bindEvents() {
    $$(".tab").forEach((tab) => {
      tab.addEventListener("click", () => switchTab(tab.dataset.tab));
    });
    $("#scanButton").addEventListener("click", scan);
    $("#protectButton").addEventListener("click", protectSelection);
    $("#rewriteButton").addEventListener("click", rewriteDraft);
    $("#closeSessionButton").addEventListener("click", closeSession);
    $("#resetSessionButton").addEventListener("click", resetSession);
    $("#draftInput").addEventListener("input", markActivity);
    document.addEventListener("click", (event) => {
      if (event.target.classList.contains("use-variant")) {
        replaceSelection(decodeURIComponent(event.target.dataset.text));
      }
    });
  }

  function boot() {
    bindEvents();
    startSessionTimer();
    scan();
  }

  if (window.Office) {
    Office.onReady(boot);
  } else {
    boot();
  }
}());
