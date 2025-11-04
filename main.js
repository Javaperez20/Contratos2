/* main.js - Integrado con UI de Ejecutivo (engranaje + modal) y sin alterar la lógica existente
   Comentarios en tercera persona impersonal, pensados para un lector con nivel básico-intermedio.
   Ajustes realizados ahora:
   - Porta en Hogar funciona para Trio, Duo, Uno (se muestra si el plan seleccionado contiene "FIJO").
   - Generación de párrafos para Movil mejorada para 0/1/2 descuentos (incluye vigencia permanente).
   - Precios y Facturación en Movil presentan descripción legible de promociones.
   - Aplicada la misma lógica de precios y facturación para Hogar.
   - Añadido el TAG <<Hogar>> (campo 'Hogar') en la plantilla contrato_template2.docx con texto según reglas solicitadas.
   - Cambio solicitado: todas las ocurrencias de "meses2 + 1" ahora se calculan como (meses2 * 2) + 1 usando computeOffsetMonths(..., 2).
   - Corrección: el resumen ALL ahora omite la frase de "y con descuento..." cuando no hay descuentos.
   - Mantiene el resto de la lógica existente (Movil, generación .docx/.pdf, Ejecutivo, etc.).
*/

/* ------------ Estado global ------------
*/
let catalog = [];
let structure = [];
const state = { sections: {} };

/* ------------ Inicialización ------------
*/
document.addEventListener('DOMContentLoaded', () => {
  loadFromServer('data.xlsx').catch(err => {
    console.error('Carga inicial falló:', err);
    showMessage('No se encontró data.xlsx en la raíz o hubo un error al leerlo. Asegúrate de servir la app vía HTTP y de que data.xlsx exista en la raíz.', true, 0);
  });
  inicializarContrato();
  initEjecutivoUI();
});

/* ------------ Mensajes visibles ------------
*/
function showMessage(msg, isError = true, timeout = 8000) {
  const el = document.getElementById('messages');
  if (!el) return;
  el.textContent = msg;
  el.style.color = isError ? '#b71c1c' : '#1b5e20';
  if (timeout) setTimeout(() => { if (el.textContent === msg) el.textContent = ''; }, timeout);
}

/* ------------ Cargar data.xlsx desde raíz ------------
*/
async function loadFromServer(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const ab = await res.arrayBuffer();
  parseWorkbook(ab);
}

/* ------------ Parseo y normalización ------------
*/
function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });

  const catalogSheetName = workbook.SheetNames.find(n => n.toLowerCase() === 'catalog') || workbook.SheetNames[0];
  const rawCatalogRows = XLSX.utils.sheet_to_json(workbook.Sheets[catalogSheetName], { defval: '' });

  catalog = rawCatalogRows.map(row => {
    const mapped = {
      Código: row['Código'] || row['Codigo'] || row['Code'] || '',
      Plan: row['Plan'] || row['Name'] || '',
      Valor: row['Valor'] || row['Value'] || row['Price'] || '',
      Promo1: row['Promo1'] || row['Promo_1'] || '',
      Meses1: row['Meses1'] || row['Meses_1'] || '',
      Promo2: row['Promo2'] || row['Promo_2'] || '',
      Meses2: row['Meses2'] || row['Meses_2'] || '',
      Detalles: row['Detalles'] || row['Details'] || '',
      Section: row['Section'] || row['Sección'] || row['Seccion'] || '',
      Subsection: row['Subsection'] || row['Subsección'] || row['Subseccion'] || '',
      ExtraFor: row['ExtraFor'] || row['Extra_for'] || ''
    };
    mapped.Valor = normalizeNumber(mapped.Valor);
    mapped.Promo1 = normalizeNumber(mapped.Promo1);
    mapped.Promo2 = normalizeNumber(mapped.Promo2);
    mapped.Meses1 = normalizeNumber(mapped.Meses1, true, true);
    mapped.Meses2 = normalizeNumber(mapped.Meses2, true, true);
    return mapped;
  });

  const firstRaw = rawCatalogRows[0] || {};
  const headerKeys = Object.keys(firstRaw).map(k => String(k).toLowerCase());
  const requiredCols = ['código','plan','valor','promo1','meses1','promo2','meses2','detalles'];
  const missing = requiredCols.filter(c => !headerKeys.includes(c));
  if (missing.length) {
    console.warn('Encabezados faltantes detectados en catalog:', missing);
    showMessage(`Advertencia: faltan encabezados obligatorios en catalog: ${missing.join(', ')}. El parser intentará mapear columnas alternativas.`, true, 12000);
  }

  const structureSheetName = workbook.SheetNames.find(n => n.toLowerCase() === 'structure');
  if (structureSheetName) {
    const rawStruct = XLSX.utils.sheet_to_json(workbook.Sheets[structureSheetName], { defval: '' });
    structure = rawStruct.map(r => ({
      Section: r['Section'] || r['Sección'] || r['Seccion'] || '',
      Subsection: r['Subsection'] || r['Subsección'] || r['Subseccion'] || '',
      ComponentType: r['ComponentType'] || r['Tipo'] || '',
      Prefixes: r['Prefixes'] || '',
      ToggleOptions: r['ToggleOptions'] || '',
      MultiPrefixes: r['MultiPrefixes'] || '',
      MaxAdditional: Number(r['MaxAdditional'] || r['MaxAdicional'] || 4),
      ExtraMapping: r['ExtraMapping'] || ''
    }));
  } else {
    // Estructura por defecto (Hogar: nuevo/cartera con inner Trio/Duo/Uno)
    structure = [
      { Section: 'Hogar', Subsection: 'nuevo', ComponentType: 'home_group', Prefixes: '', MaxAdditional: 0 },
      { Section: 'Hogar', Subsection: 'cartera', ComponentType: 'home_group', Prefixes: '', MaxAdditional: 0 },
      { Section: 'Movil', Subsection: 'nuevo', ComponentType: 'movil_group', MultiPrefixes: 'multi:NM,datos:ND,voz:NV', MaxAdditional: 4, ExtraMapping: 'NM02:NM02S;NM03:NM03S' },
      { Section: 'Movil', Subsection: 'cartera', ComponentType: 'movil_group', MultiPrefixes: 'multi:CM,datos:CD,voz:CV', MaxAdditional: 4, ExtraMapping: 'CM02:CM02S;CM03:CM03S' }
    ];
    // showMessage('Hoja "structure" no encontrada: uso estructura por defecto con jerarquía Hogar (Nuevo/Cartera).', false, 6000);
  }

  console.info('Catalog rows (muestra 10):', catalog.slice(0,10));
  buildUIFromStructure();
}

/* ------------ Utilidades de número y prefijos ------------
*/
function normalizeNumber(val, integer = false, allowText = false) {
  if (val === null || val === undefined || val === '') return allowText ? '' : '';
  if (typeof val === 'number') return integer ? Math.floor(val) : val;

  let s = String(val).trim();
  if (!s) return allowText ? '' : '';

  const low = s.toLowerCase();
  if (['no aplica','noaplica','n/a','na','-'].includes(low)) return '';

  const match = s.match(/-?\d[\d\.\,]*/);
  if (match) {
    let numStr = match[0];
    if (numStr.indexOf('.') !== -1 && numStr.indexOf(',') !== -1) {
      numStr = numStr.replace(/\./g, '').replace(',', '.');
    } else if (numStr.indexOf(',') !== -1 && numStr.indexOf('.') === -1) {
      numStr = numStr.replace(',', '.');
    } else {
      const dots = (numStr.match(/\./g) || []).length;
      if (dots > 1) numStr = numStr.replace(/\./g, '');
    }

    const num = Number(numStr);
    if (!Number.isNaN(num)) return integer ? Math.floor(num) : num;
  }

  if (allowText) return s;
  return '';
}

function matchesPrefix(code = '', prefix = '') {
  if (!code || !prefix) return false;
  code = String(code);
  prefix = String(prefix);
  if (code === prefix) return true;
  if (!code.startsWith(prefix)) return false;
  const next = code.charAt(prefix.length);
  if (!next) return true;
  return !(/[A-Za-z]/.test(next));
}

/* ------------ Utilidad solicitada: computeOffsetMonths ------------
   Permite calcular (meses * multiplier) + 1 de forma segura.
   La petición era: cambiar las expresiones `meses2 + 1` para que sean (meses2 * 2) + 1.
*/
function computeOffsetMonths(m, multiplier = 1) {
  const n = Number(m);
  if (Number.isNaN(n)) return '';
  return n * multiplier + 1;
}

/* ------------ Reset / limpieza ------------
*/
function resetSectionSelections(sectionName) {
  const sec = state.sections[sectionName];
  if (!sec || !sec.subsections) return;
  Object.keys(sec.subsections).forEach(subName => {
    const st = sec.subsections[subName];
    if (!st || !st.elementos) return;
    if (st.elementos.mainSelect) {
      try { st.elementos.mainSelect.value = ''; } catch (e) {}
    }
    if (Array.isArray(st.elementos.lines)) {
      st.elementos.lines.forEach(line => {
        try {
          if (line.select) line.select.value = '';
          if (line.portaCheckbox) {
            line.portaCheckbox.checked = false;
            if (line.portaFields) line.portaFields.style.display = 'none';
          }
          if (line.portaNumeroInput) line.portaNumeroInput.value = '';
          if (line.portaDonanteInput) line.portaDonanteInput.value = '';
        } catch (e) {}
      });
    }
    if (st.elementos && st.elementos.innerSelect) {
      try { st.elementos.innerSelect.value = ''; } catch (e) {}
    }
    // Limpieza de porta en Home si existe
    if (st.elementos && st.elementos.portaCheckbox) {
      try {
        st.elementos.portaCheckbox.checked = false;
        if (st.elementos.portaFields) st.elementos.portaFields.style.display = 'none';
        if (st.elementos.portaNumeroInput) st.elementos.portaNumeroInput.value = '';
        if (st.elementos.portaDonanteInput) st.elementos.portaDonanteInput.value = '';
      } catch (e) {}
    }
  });
}

function clearSelectionsExceptSection(keepSection) {
  Object.keys(state.sections).forEach(secName => {
    if (secName !== keepSection) resetSectionSelections(secName);
  });
}

function resetSubsectionsExcept(sectionName, keepSub) {
  const sec = state.sections[sectionName];
  if (!sec || !sec.subsections) return;
  Object.keys(sec.subsections).forEach(subName => {
    if (subName === keepSub) return;
    const st = sec.subsections[subName];
    if (!st || !st.elementos) return;
    if (st.elementos.mainSelect) {
      try { st.elementos.mainSelect.value = ''; } catch (e) {}
    }
    if (Array.isArray(st.elementos.lines)) {
      st.elementos.lines.forEach(line => {
        try {
          if (line.select) line.select.value = '';
          if (line.portaCheckbox) {
            line.portaCheckbox.checked = false;
            if (line.portaFields) line.portaFields.style.display = 'none';
          }
          if (line.portaNumeroInput) line.portaNumeroInput.value = '';
          if (line.portaDonanteInput) line.portaDonanteInput.value = '';
        } catch (e) {}
      });
    }
    if (st.elementos && st.elementos.innerSelect) {
      try { st.elementos.innerSelect.value = ''; } catch (e) {}
    }
    if (st.elementos && st.elementos.portaCheckbox) {
      try {
        st.elementos.portaCheckbox.checked = false;
        if (st.elementos.portaFields) st.elementos.portaFields.style.display = 'none';
        if (st.elementos.portaNumeroInput) st.elementos.portaNumeroInput.value = '';
        if (st.elementos.portaDonanteInput) st.elementos.portaDonanteInput.value = '';
      } catch (e) {}
    }
  });
}

/* ------------ Construcción dinámica de UI ------------
*/
function buildUIFromStructure() {
  const bySection = {};
  structure.forEach(s => { if (!bySection[s.Section]) bySection[s.Section] = []; bySection[s.Section].push(s); });

  const root = document.getElementById('tarifario-root');
  root.innerHTML = '';

  const tabs = document.createElement('div'); tabs.className = 'tabs';
  const sectionNames = Object.keys(bySection);
  sectionNames.forEach((sec, i) => {
    const btn = document.createElement('button');
    btn.className = 'tab-btn' + (i === 0 ? ' active' : '');
    btn.textContent = sec;
    btn.dataset.section = sec;
    tabs.appendChild(btn);
    btn.addEventListener('click', () => {
      clearSelectionsExceptSection(sec);
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderSection(sec, bySection[sec]);
    });
    state.sections[sec] = { subsections: {}, activeSub: null };
  });
  root.appendChild(tabs);

  const panel = document.createElement('div'); panel.className = 'section-panel';
  root.appendChild(panel);

  if (sectionNames.length) renderSection(sectionNames[0], bySection[sectionNames[0]]);
}

function renderSection(sectionName, subsections) {
  const panel = document.querySelector('#tarifario-root .section-panel');
  panel.innerHTML = '';

  const subtabs = document.createElement('div'); subtabs.className = 'subtabs';
  panel.appendChild(subtabs);

  const contentArea = document.createElement('div');
  panel.appendChild(contentArea);

  subsections.forEach((sub, idx) => {
    const sbtn = document.createElement('button');
    sbtn.className = 'subtab-btn' + (idx === 0 ? ' active' : '');
    sbtn.textContent = sub.Subsection;
    sbtn.dataset.sub = sub.Subsection;
    subtabs.appendChild(sbtn);

    state.sections[sectionName].subsections[sub.Subsection] = { config: sub, elementos: {}, main: {} };

    sbtn.addEventListener('click', () => {
      resetSubsectionsExcept(sectionName, sub.Subsection);
      subtabs.querySelectorAll('.subtab-btn').forEach(b => b.classList.remove('active'));
      sbtn.classList.add('active');
      renderSubsectionContent(contentArea, sectionName, sub);
      state.sections[sectionName].activeSub = sub.Subsection;
    });

    if (idx === 0) {
      renderSubsectionContent(contentArea, sectionName, sub);
      state.sections[sectionName].activeSub = sub.Subsection;
    }
  });
}

function renderSubsectionContent(container, sectionName, subCfg) {
  container.innerHTML = '';
  const row = document.createElement('div'); row.className = 'row';
  const colLeft = document.createElement('div'); colLeft.className = 'col';
  const colRight = document.createElement('div'); colRight.className = 'col';
  row.appendChild(colLeft); row.appendChild(colRight);
  container.appendChild(row);

  const stateSub = state.sections[sectionName].subsections[subCfg.Subsection];

  const optionsFromPrefixes = (prefixStr) => {
    if (!prefixStr) return [];
    const prefs = prefixStr.split(',').map(s => s.trim()).filter(Boolean);
    return catalog.filter(r => prefs.some(p => matchesPrefix(String(r.Código || ''), p)));
  };

  if (subCfg.ComponentType === 'home_group') {
    const innerTabs = document.createElement('div'); innerTabs.className = 'subtabs';
    const innerTypes = [
      { key: 'trio', label: 'Trio' },
      { key: 'duo', label: 'Duo' },
      { key: 'uno', label: 'Uno' }
    ];
    innerTypes.forEach((it, i) => {
      const ibtn = document.createElement('button');
      ibtn.className = 'subtab-btn' + (i === 0 ? ' active' : '');
      ibtn.textContent = it.label;
      innerTabs.appendChild(ibtn);
      ibtn.addEventListener('click', () => {
        innerTabs.querySelectorAll('.subtab-btn').forEach(b => b.classList.remove('active'));
        ibtn.classList.add('active');
        renderHomeInner(it.key, colLeft, colRight, sectionName, subCfg);
      });
    });
    colLeft.appendChild(innerTabs);
    renderHomeInner(innerTypes[0].key, colLeft, colRight, sectionName, subCfg);
    stateSub.elementos = stateSub.elementos || {};
  } else if (subCfg.ComponentType === 'trio') {
    colLeft.appendChild(createLabel(`Plan ${subCfg.Subsection}`));
    const sel = document.createElement('select');
    sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    const opts = optionsFromPrefixes(subCfg.Prefixes || 'T');
    opts.forEach(o => sel.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${o.Plan}</option>`));
    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(colRight); else updateDetalleYPrecios(colRight, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
    });
    colLeft.appendChild(sel);
    stateSub.elementos.mainSelect = sel;
    colRight.appendChild(createOfferBox('Selecciona un plan para ver detalles.'));
  } else if (subCfg.ComponentType === 'duo' || subCfg.ComponentType === 'uno') {
    colLeft.appendChild(createLabel(`${subCfg.Subsection} - Opciones`));
    const segDiv = document.createElement('div'); segDiv.className = 'segmented-toggle';
    colLeft.appendChild(segDiv);

    const sel = document.createElement('select'); sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    colLeft.appendChild(sel);
    stateSub.elementos.mainSelect = sel;

    const pairs = parseToggleOptions(subCfg.ToggleOptions || '');
    if (pairs.length === 0) {
      const note = document.createElement('div'); note.textContent = 'No hay opciones de toggle definidas en structure';
      colLeft.appendChild(note);
    } else {
      pairs.forEach((p, i) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        const labelText = String(p.key).replace(/[_\-]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
        btn.textContent = labelText;
        if (i === 0) btn.classList.add('active');
        btn.addEventListener('click', () => {
          segDiv.querySelectorAll('button').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          populateSelectFromPrefixes(sel, p.prefixes);
          sel.value = '';
          updateDetalleYPreciosDefault(colRight);
        });
        segDiv.appendChild(btn);
      });
      // carga primera opción por defecto
      populateSelectFromPrefixes(sel, pairs[0].prefixes);
    }

    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(colRight); else updateDetalleYPrecios(colRight, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
    });
    colRight.appendChild(createOfferBox('Selecciona un plan para ver detalles.'));
  } else if (subCfg.ComponentType === 'movil_group') {
    const wrapper = document.createElement('div'); wrapper.className = 'movil-lines';
    colLeft.appendChild(wrapper);

    const principal = createMovilLine(sectionName, subCfg, 0, false);
    wrapper.appendChild(principal.lineElement);
    stateSub.elementos.lines = [principal];
    stateSub.main = { lines: [{ idx: 0, codigo: '' }] };

    const addBtn = document.createElement('button'); addBtn.type = 'button'; addBtn.className = 'btn btn-add';
    addBtn.textContent = '+ Añadir línea';
    addBtn.addEventListener('click', () => {
      const max = Math.max(0, subCfg.MaxAdditional || 4);
      const currentAdditional = stateSub.elementos.lines.length - 1;
      if (currentAdditional >= max) { alert(`Máximo ${max} líneas adicionales`); return; }
      const newIdx = stateSub.elementos.lines.length;
      const newLine = createMovilLine(sectionName, subCfg, newIdx, true);
      stateSub.elementos.lines.push(newLine);
      wrapper.appendChild(newLine.lineElement);
      actualizarMovilSection(sectionName, subCfg.Subsection);
    });
    colLeft.appendChild(addBtn);

    const detallesBox = document.createElement('div'); detallesBox.className = 'offer-details';
    detallesBox.innerHTML = 'Selecciona las líneas para ver detalles.';
    const preciosBox = document.createElement('div'); preciosBox.className = 'precios';
    const factBox = document.createElement('div'); factBox.className = 'facturacion';
    colRight.appendChild(detallesBox); colRight.appendChild(preciosBox); colRight.appendChild(factBox);

    stateSub.elementos.detallesBox = detallesBox;
    stateSub.elementos.preciosBox = preciosBox;
    stateSub.elementos.facturacionBox = factBox;
  } else {
    colLeft.appendChild(createLabel(`Plan ${subCfg.Subsection}`));
    const sel = document.createElement('select');
    sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    const opts = optionsFromPrefixes(subCfg.Prefixes || '');
    opts.forEach(o => sel.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${o.Plan}</option>`));
    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(colRight); else updateDetalleYPrecios(colRight, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
    });
    colLeft.appendChild(sel);
    colRight.appendChild(createOfferBox('Selecciona un plan para ver detalles.'));
    stateSub.elementos.mainSelect = sel;
  }
}

/* ------------ Render interno para Hogar ------------
*/
function renderHomeInner(type, colLeft, colRight, sectionName, subCfg) {
  // Mantener la pestaña interna si existe
  const innerTabs = colLeft.querySelector('.subtabs');
  colLeft.innerHTML = '';
  if (innerTabs) colLeft.appendChild(innerTabs);
  colRight.innerHTML = '';

  const stateSub = state.sections[sectionName].subsections[subCfg.Subsection];
  stateSub.elementos = stateSub.elementos || {};

  // Crea (o reusa) el contenedor de porta y lo inserta justo después del select provisto.
  function createOrAttachPorta(selectEl) {
    try {
      if (!stateSub.elementos) stateSub.elementos = {};
      // Si ya existe un portaContainer, reubicarlo junto al select actual
      if (stateSub.elementos.portaContainer) {
        // si está en otro lugar del DOM, moverlo
        if (stateSub.elementos.portaContainer.parentElement !== selectEl.parentElement) {
          // quitar de su padre actual (si sigue en DOM)
          if (stateSub.elementos.portaContainer.parentElement) {
            try { stateSub.elementos.portaContainer.parentElement.removeChild(stateSub.elementos.portaContainer); } catch (e) {}
          }
          // insertarlo después del select
          selectEl.insertAdjacentElement('afterend', stateSub.elementos.portaContainer);
        } else {
          // asegurarnos que esté inmediatamente después del select
          if (selectEl.nextSibling !== stateSub.elementos.portaContainer) {
            selectEl.insertAdjacentElement('afterend', stateSub.elementos.portaContainer);
          }
        }
        return stateSub.elementos;
      }

      // Crear estructura nueva
      const portaContainer = document.createElement('div');
      portaContainer.className = 'porta-container';
      portaContainer.style.display = 'flex';
      portaContainer.style.flexDirection = 'column';
      portaContainer.style.alignItems = 'flex-end';
      portaContainer.style.marginLeft = '8px';

      const portaLabel = document.createElement('label');
      portaLabel.className = 'porta-label';
      const portaCheckbox = document.createElement('input');
      portaCheckbox.type = 'checkbox';
      portaCheckbox.className = 'porta-checkbox';
      portaLabel.appendChild(portaCheckbox);
      const portaText = document.createElement('span'); portaText.textContent = ' Porta';
      portaLabel.appendChild(portaText);
      portaContainer.appendChild(portaLabel);

      const portaFields = document.createElement('div');
      portaFields.className = 'porta-fields';
      portaFields.style.display = 'none';
      portaFields.style.marginTop = '6px';
      portaFields.style.flexDirection = 'column';
      portaFields.style.gap = '6px';
      const inputNumero = document.createElement('input'); inputNumero.type = 'text'; inputNumero.placeholder = 'Número a portar';
      inputNumero.className = 'porta-numero';
      const inputDonante = document.createElement('input'); inputDonante.type = 'text'; inputDonante.placeholder = 'Compañía donante';
      inputDonante.className = 'porta-donante';
      portaFields.appendChild(inputNumero);
      portaFields.appendChild(inputDonante);

      portaContainer.appendChild(portaFields);

      // Listener para checkbox: mostrar/ocultar campos
      portaCheckbox.addEventListener('change', () => {
        portaFields.style.display = portaCheckbox.checked ? 'block' : 'none';
      });

      // Insertar después del select
      selectEl.insertAdjacentElement('afterend', portaContainer);

      // Guardar referencias
      stateSub.elementos.portaContainer = portaContainer;
      stateSub.elementos.portaCheckbox = portaCheckbox;
      stateSub.elementos.portaFields = portaFields;
      stateSub.elementos.portaNumeroInput = inputNumero;
      stateSub.elementos.portaDonanteInput = inputDonante;

      return stateSub.elementos;
    } catch (err) {
      console.error('Error creando/adjuntando porta controls:', err);
      return {};
    }
  }

  // Mostrar u ocultar el portaContainer según el plan (si contiene 'FIJO')
  function evaluatePortaVisibility(selectEl) {
    try {
      if (!selectEl) return;
      const elems = createOrAttachPorta(selectEl);
      if (!elems || !elems.portaContainer) return;
      const code = selectEl.value || '';
      const plan = findByCode(code);
      const hasFijo = plan && typeof plan.Plan === 'string' && plan.Plan.toUpperCase().includes('FIJO');
      if (hasFijo) {
        elems.portaContainer.style.display = 'flex';
      } else {
        // ocultar y limpiar
        elems.portaContainer.style.display = 'none';
        try {
          if (elems.portaCheckbox) elems.portaCheckbox.checked = false;
          if (elems.portaFields) elems.portaFields.style.display = 'none';
          if (elems.portaNumeroInput) elems.portaNumeroInput.value = '';
          if (elems.portaDonanteInput) elems.portaDonanteInput.value = '';
        } catch (e) {}
      }
    } catch (err) {
      console.error('Error evaluando visibilidad de porta:', err);
    }
  }

  // Crear cajas de detalles/precios/facturación para Hogar (si aún no existen)
  function ensureRightBoxes() {
    if (!stateSub.elementos.detallesBox) {
      const detallesBox = document.createElement('div'); detallesBox.className = 'offer-details';
      detallesBox.innerHTML = 'Selecciona un plan para ver detalles.';
      stateSub.elementos.detallesBox = detallesBox;
    }
    if (!stateSub.elementos.preciosBox) {
      const preciosBox = document.createElement('div'); preciosBox.className = 'precios';
      stateSub.elementos.preciosBox = preciosBox;
    }
    if (!stateSub.elementos.facturacionBox) {
      const factBox = document.createElement('div'); factBox.className = 'facturacion';
      stateSub.elementos.facturacionBox = factBox;
    }
    // Limpia y agrega a colRight (reemplaza)
    colRight.innerHTML = '';
    colRight.appendChild(stateSub.elementos.detallesBox);
    colRight.appendChild(stateSub.elementos.preciosBox);
    colRight.appendChild(stateSub.elementos.facturacionBox);
  }

  // Actualizar precios/facturación para Hogar (usa buildPromoDescription)
  function updateHomePricing() {
    try {
      const sel = stateSub.elementos && stateSub.elementos.mainSelect ? stateSub.elementos.mainSelect : null;
      if (!sel) {
        if (stateSub.elementos && stateSub.elementos.preciosBox) stateSub.elementos.preciosBox.innerHTML = 'Selecciona un plan para ver precios.';
        if (stateSub.elementos && stateSub.elementos.facturacionBox) stateSub.elementos.facturacionBox.innerHTML = '';
        return;
      }
      const code = sel.value || '';
      const plan = findByCode(code);
      if (!plan) {
        if (stateSub.elementos && stateSub.elementos.preciosBox) stateSub.elementos.preciosBox.innerHTML = 'Selecciona un plan para ver precios.';
        if (stateSub.elementos && stateSub.elementos.facturacionBox) stateSub.elementos.facturacionBox.innerHTML = '';
        return;
      }
      const info = buildPromoDescription(plan);
      if (stateSub.elementos && stateSub.elementos.preciosBox) {
        stateSub.elementos.preciosBox.innerHTML = `<div>${escapeHtml(info.preciosText)}</div>`;
      }
      if (stateSub.elementos && stateSub.elementos.facturacionBox) {
        stateSub.elementos.facturacionBox.innerHTML = `<div>${escapeHtml(info.factText)}</div>`;
      }
    } catch (err) {
      console.error('Error actualizando precios Hogar:', err);
    }
  }

  // Renderizar cada tipo interno: trio / duo / uno
  if (type === 'trio') {
    colLeft.appendChild(createLabel('Plan Trio'));
    const sel = document.createElement('select');
    sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    const opts = catalog.filter(r => matchesPrefix(String(r.Código || ''), 'T'));
    opts.forEach(o => sel.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${o.Plan}</option>`));
    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(stateSub.elementos && stateSub.elementos.detallesBox ? stateSub.elementos.detallesBox : colRight); else updateDetalleYPrecios(stateSub.elementos && stateSub.elementos.detallesBox ? stateSub.elementos.detallesBox : colRight, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
      evaluatePortaVisibility(sel);
      ensureRightBoxes();
      updateHomePricing();
    });
    colLeft.appendChild(sel);
    // crear/adjuntar pero oculto inicialmente
    createOrAttachPorta(sel);
    if (stateSub.elementos && stateSub.elementos.portaContainer) stateSub.elementos.portaContainer.style.display = 'none';
    stateSub.elementos.mainSelect = sel;

    // crear cajas en el derecho
    ensureRightBoxes();
    stateSub.elementos.detallesBox.innerHTML = createOfferBox('Selecciona un plan Trio para ver detalles.').innerHTML;
    updateHomePricing();
  } else if (type === 'duo') {
    colLeft.appendChild(createLabel('Plan Duo - Opciones'));
    const segDiv = document.createElement('div'); segDiv.className = 'segmented-toggle';
    colLeft.appendChild(segDiv);

    const sel = document.createElement('select'); sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    colLeft.appendChild(sel);
    stateSub.elementos.mainSelect = sel;

    const pairs = parseToggleOptions('fibra_tv:DT,fibra_fijo:DF,tv_fijo:DTF');
    pairs.forEach((p, i) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      const labelText = String(p.key).replace(/[_\-]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
      btn.textContent = labelText;
      if (i === 0) btn.classList.add('active');
      btn.addEventListener('click', () => {
        segDiv.querySelectorAll('button').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        populateSelectFromPrefixes(sel, p.prefixes);
        sel.value = '';
        // actualizar detalles y precios
        ensureRightBoxes();
        updateDetalleYPreciosDefault(stateSub.elementos.detallesBox);
        updateHomePricing();
        // re-evaluar visibilidad de porta según nuevas opciones (mantener oculto hasta selección)
        createOrAttachPorta(sel);
        evaluatePortaVisibility(sel);
      });
      segDiv.appendChild(btn);
    });
    if (pairs[0]) {
      populateSelectFromPrefixes(sel, pairs[0].prefixes);
      // re-evaluar (por si la lista inicial contiene un plan seleccionado por defecto)
      createOrAttachPorta(sel);
      evaluatePortaVisibility(sel);
    }

    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(stateSub.elementos.detallesBox); else updateDetalleYPrecios(stateSub.elementos.detallesBox, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
      evaluatePortaVisibility(sel);
      ensureRightBoxes();
      updateHomePricing();
    });

    // Crear controles de porta (ocultos por defecto)
    createOrAttachPorta(sel);
    if (stateSub.elementos && stateSub.elementos.portaContainer) stateSub.elementos.portaContainer.style.display = 'none';

    // crear cajas derecho
    ensureRightBoxes();
    stateSub.elementos.detallesBox.innerHTML = createOfferBox('Selecciona un plan Duo para ver detalles.').innerHTML;
    updateHomePricing();
  } else if (type === 'uno') {
    colLeft.appendChild(createLabel('Plan Uno - Opciones'));
    const segDiv = document.createElement('div'); segDiv.className = 'segmented-toggle';
    colLeft.appendChild(segDiv);

    const sel = document.createElement('select'); sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    colLeft.appendChild(sel);
    stateSub.elementos.mainSelect = sel;

    const pairs = parseToggleOptions('fibra:F,tv:TV,fijo:FI');
    pairs.forEach((p, i) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      const labelText = String(p.key).replace(/[_\-]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
      btn.textContent = labelText;
      if (i === 0) btn.classList.add('active');
      btn.addEventListener('click', () => {
        segDiv.querySelectorAll('button').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        populateSelectFromPrefixes(sel, p.prefixes);
        sel.value = '';
        updateDetalleYPreciosDefault(stateSub.elementos.detallesBox);
        ensureRightBoxes();
        updateHomePricing();
        evaluatePortaVisibility(sel);
      });
      segDiv.appendChild(btn);
    });
    if (pairs[0]) {
      populateSelectFromPrefixes(sel, pairs[0].prefixes);
      evaluatePortaVisibility(sel);
    }

    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(stateSub.elementos.detallesBox); else updateDetalleYPrecios(stateSub.elementos.detallesBox, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
      evaluatePortaVisibility(sel);
      ensureRightBoxes();
      updateHomePricing();
    });

    // Crear controles de porta (ocultos por defecto)
    createOrAttachPorta(sel);
    if (stateSub.elementos && stateSub.elementos.portaContainer) stateSub.elementos.portaContainer.style.display = 'none';

    // crear cajas derecho
    ensureRightBoxes();
    stateSub.elementos.detallesBox.innerHTML = createOfferBox('Selecciona un plan Uno para ver detalles.').innerHTML;
    updateHomePricing();
  }
}

/* ------------ Helpers y movil (con buildPromoDescription reutilizable) ------------
*/
function createLabel(text) { const l = document.createElement('label'); l.textContent = text; l.style.display = 'block'; l.style.marginTop = '6px'; return l; }
function createOfferBox(initialText) {
  const box = document.createElement('div');
  box.className = 'offer-details';
  const txt = document.createElement('div');
  txt.className = 'offer-details-text';
  txt.textContent = initialText;
  box.appendChild(txt);
  return box;
}
function findByCode(code) { return catalog.find(r => r.Código === code) || null; }
function updateDetalleYPrecios(containerColRight, plan) {
  containerColRight.innerHTML = '';
  const box = document.createElement('div'); box.className = 'offer-details';
  const title = document.createElement('div'); title.className = 'offer-details-title';
  title.innerHTML = plan ? `<strong>${escapeHtml(plan.Plan)}</strong>` : '<strong>Plan</strong>';
  const details = document.createElement('div'); details.className = 'offer-details-text';
  details.textContent = plan ? (plan.Detalles || '') : 'Selecciona un plan para ver detalles.';
  box.appendChild(title);
  box.appendChild(details);

  if (plan) {
    const precios = document.createElement('div'); precios.className = 'precios';
    const lines = [];
    if (plan.Promo1 !== '' && plan.Promo1 !== null && plan.Promo1 !== undefined) {
      const months1Text = (typeof plan.Meses1 === 'number') ? `${plan.Meses1} meses` : (plan.Meses1 || '-');
      lines.push(`Promo 1: $${plan.Promo1} (${months1Text})`);
    }
    if (plan.Promo2 !== '' && plan.Promo2 !== null && plan.Promo2 !== undefined) {
      const months2Text = (typeof plan.Meses2 === 'number') ? `${plan.Meses2} meses` : (plan.Meses2 || '-');
      lines.push(`Promo 2: $${plan.Promo2} (${months2Text})`);
    }
    if (plan.Valor !== '' && plan.Valor !== null && plan.Valor !== undefined) {
      lines.push(`Sin descuento: $${plan.Valor}`);
    }
    precios.innerHTML = lines.map(l => `<div>${escapeHtml(l)}</div>`).join('');
    box.appendChild(precios);
  } else {
    const defaultBox = createOfferBox('Selecciona un plan para ver detalles.');
    containerColRight.appendChild(defaultBox);
    return;
  }
  containerColRight.appendChild(box);
}
function updateDetalleYPreciosDefault(containerColRight) {
  containerColRight.innerHTML = '';
  const box = createOfferBox('Selecciona un plan para ver detalles.');
  containerColRight.appendChild(box);
}
function parseToggleOptions(str) {
  if (!str) return [];
  return str.split(',').map(s => {
    const [k, pref] = s.split(':').map(x => x && x.trim());
    return { key: k || '', prefixes: (pref || '').split('|').map(p => p.trim()).filter(Boolean) };
  }).filter(x => x.key);
}
function populateSelectFromPrefixes(selectEl, prefixes) {
  selectEl.innerHTML = `<option value="">-- Selecciona --</option>`;
  if (!prefixes || prefixes.length === 0) return;
  const opts = catalog.filter(r => prefixes.some(p => matchesPrefix(String(r.Código || ''), String(p || '').trim())));
  opts.forEach(o => selectEl.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${escapeHtml(o.Plan)}</option>`));
}

/* Helper reutilizable para construir textos de precios/facturación */
function buildPromoDescription(plan) {
  const valorPlan = plan.Valor || '';
  const promo1 = (plan.Promo1 === '' || plan.Promo1 === null || plan.Promo1 === undefined) ? '' : String(plan.Promo1);
  const promo2 = (plan.Promo2 === '' || plan.Promo2 === null || plan.Promo2 === undefined) ? '' : String(plan.Promo2);
  const rawM1 = plan.Meses1;
  const rawM2 = plan.Meses2;
  const meses1 = (rawM1 === '' || rawM1 === null || rawM1 === undefined) ? null : (Number(rawM1) && !Number.isNaN(Number(rawM1)) ? Number(rawM1) : (String(rawM1).trim() || null));
  const meses2 = (rawM2 === '' || rawM2 === null || rawM2 === undefined) ? null : (Number(rawM2) && !Number.isNaN(Number(rawM2)) ? Number(rawM2) : (String(rawM2).trim() || null));

  // texto detallado para "precios"
  let preciosText = '';
  if (!promo1) {
    preciosText = `Valor total de $${valorPlan}.`;
  } else if (promo1 && !promo2) {
    if (meses1 && typeof meses1 === 'number') {
      preciosText = `Promo de $${promo1} por ${meses1} meses, a partir del mes ${meses1 + 1} valor normal de $${valorPlan}.`;
    } else {
      preciosText = `Promo de $${promo1} por vigencia permanente.`;
    }
  } else {
    // dos promos
    if (meses1 && typeof meses1 === 'number' && meses2 && typeof meses2 === 'number') {
      preciosText = `Promo de $${promo1} por ${meses1} meses, a partir del mes ${meses1 + 1} promo de $${promo2} por ${meses2} meses, a partir del mes ${computeOffsetMonths(meses2, 2)} valor normal de $${valorPlan}.`;
    } else if (meses1 && typeof meses1 === 'number' && (!meses2 || meses2 === null)) {
      preciosText = `Promo de $${promo1} por ${meses1} meses, a partir del mes ${meses1 + 1} promo de $${promo2} por vigencia permanente.`;
    } else {
      preciosText = `Promociones: $${promo1}${meses1 ? ` (${meses1} meses)` : ''} y $${promo2}${meses2 ? ` (${meses2} meses)` : ''}. Valor normal $${valorPlan}.`;
    }
  }

  // texto resumido para "facturación"
  let factText = '';
  if (!promo1) {
    factText = `${plan.Plan}: $${valorPlan} (Sin descuento)`;
  } else if (promo1 && !promo2) {
    if (meses1 && typeof meses1 === 'number') {
      factText = `${plan.Plan}: $${promo1} (Promo) / $${valorPlan} (Sin descuento) — ${meses1} meses`;
    } else {
      factText = `${plan.Plan}: $${promo1} (Promo permanente) / $${valorPlan} (Sin descuento)`;
    }
  } else {
    if (meses1 && typeof meses1 === 'number' && meses2 && typeof meses2 === 'number') {
      factText = `${plan.Plan}: $${promo1} (Promo1) / $${promo2} (Promo2) / $${valorPlan} (Sin descuento) — ${meses1}m + ${meses2}m`;
    } else if (meses1 && typeof meses1 === 'number') {
      factText = `${plan.Plan}: $${promo1} (Promo1) / $${promo2} (Promo2 permanente) / $${valorPlan} (Sin descuento) — ${meses1}m`;
    } else {
      factText = `${plan.Plan}: $${promo1} / $${promo2} / $${valorPlan}`;
    }
  }

  return { preciosText, factText };
}

/* Movil functions (sin cambios funcionales respecto a tu original salvo la actualización de precios/facturación) */
function createMovilLine(sectionName, cfg, idx, esSecundaria) {
  const card = document.createElement('div');
  card.className = 'movil-line-card';

  const header = document.createElement('div');
  header.className = 'movil-line-header';

  const title = document.createElement('div');
  title.className = 'movil-line-title';
  title.textContent = esSecundaria ? `Adicional ${idx}` : 'Línea Principal';

  const headerRight = document.createElement('div');
  headerRight.className = 'movil-line-header-right';

  const segSmall = document.createElement('div'); segSmall.className = 'segmented-toggle small';
  const estados = ['multi','datos','voz'];
  estados.forEach((e,i) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = e.charAt(0).toUpperCase() + e.slice(1);
    if (i === 0) btn.classList.add('active');
    btn.addEventListener('click', () => {
      segSmall.querySelectorAll('button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      cargarOpcionesMovilDesdeCfg(sectionName, cfg, select, e, idx, esSecundaria);
      if (select.value) select.dispatchEvent(new Event('change'));
      else {
        const stateSub = findStateSub(sectionName, cfg.Subsection);
        if (stateSub && stateSub.elementos && stateSub.elementos.detallesBox) {
          stateSub.elementos.detallesBox.innerHTML = `<div><strong>Línea ${idx === 0 ? 'Principal' : 'Adicional ' + idx}</strong><br>Selecciona un plan para ver detalles.</div>`;
        }
        actualizarMovilSection(sectionName, cfg.Subsection);
      }
    });
    segSmall.appendChild(btn);
  });

  headerRight.appendChild(segSmall);

  header.appendChild(title);
  header.appendChild(headerRight);
  card.appendChild(header);

  const content = document.createElement('div'); content.className = 'movil-line-content';

  const select = document.createElement('select');
  select.id = `${sectionName}-${cfg.Subsection}-select-${idx}`;
  select.innerHTML = `<option value="">-- Selecciona --</option>`;

  const portaContainer = document.createElement('div'); portaContainer.className = 'porta-container';
  const portaLabel = document.createElement('label'); portaLabel.className = 'porta-label';
  const portaCheckbox = document.createElement('input');
  portaCheckbox.type = 'checkbox';
  portaCheckbox.className = 'porta-checkbox';
  portaLabel.appendChild(portaCheckbox);
  const portaText = document.createElement('span'); portaText.textContent = ' Porta';
  portaLabel.appendChild(portaText);
  portaContainer.appendChild(portaLabel);

  const portaFields = document.createElement('div'); portaFields.classList.add('porta-fields');
  portaFields.style.display = 'none';
  const inputNumero = document.createElement('input'); inputNumero.type = 'text'; inputNumero.placeholder = 'Número a portar';
  inputNumero.className = 'porta-numero';
  const inputDonante = document.createElement('input'); inputDonante.type = 'text'; inputDonante.placeholder = 'Compañía donante';
  inputDonante.className = 'porta-donante';
  portaFields.appendChild(inputNumero);
  portaFields.appendChild(inputDonante);

  portaContainer.appendChild(portaFields);

  portaCheckbox.addEventListener('change', () => {
    portaFields.style.display = portaCheckbox.checked ? 'block' : 'none';
    actualizarMovilSection(sectionName, cfg.Subsection);
  });

  select.addEventListener('change', () => {
    if (!select.value) {
      const stateSub = findStateSub(sectionName, cfg.Subsection);
      if (stateSub && stateSub.elementos && stateSub.elementos.detallesBox) {
        stateSub.elementos.detallesBox.innerHTML = `<div><strong>Línea ${idx === 0 ? 'Principal' : 'Adicional ' + idx}</strong><br>Selecciona un plan para ver detalles.</div>`;
      }
      actualizarMovilSection(sectionName, cfg.Subsection);
    } else {
      actualizarMovilSection(sectionName, cfg.Subsection);
    }
  });

  content.appendChild(select);
  content.appendChild(portaContainer);
  card.appendChild(content);

  cargarOpcionesMovilDesdeCfg(sectionName, cfg, select, 'multi', idx, esSecundaria);

  let removeBtn = null;
  if (esSecundaria) {
    removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn btn-remove';
    removeBtn.textContent = '🗑';
    removeBtn.title = 'Eliminar línea';
    removeBtn.addEventListener('click', () => {
      const stateSub = findStateSub(sectionName, cfg.Subsection);
      if (!stateSub) return;
      const arr = stateSub.elementos.lines;
      const pos = arr.findIndex(l => l.lineElement === card);
      if (pos >= 0) arr.splice(pos, 1);
      card.remove();
      arr.slice(1).forEach((l, i) => {
        const newIdx = i+1;
        l.lineElement.dataset.idx = newIdx;
        const lbl = l.lineElement.querySelector('.movil-line-title');
        if (lbl) lbl.textContent = `Adicional ${newIdx}`;
      });
      actualizarMovilSection(sectionName, cfg.Subsection);
    });
    card.appendChild(removeBtn);
  }

  return {
    idx,
    esSecundaria,
    select,
    seg: segSmall,
    portaCheckbox,
    portaNumeroInput: inputNumero,
    portaDonanteInput: inputDonante,
    portaFields,
    lineElement: card
  };
}

function findStateSub(sectionName, subName) { if (!state.sections[sectionName]) return null; return state.sections[sectionName].subsections[subName] || null; }

function cargarOpcionesMovilDesdeCfg(sectionName, cfg, selectEl, estado, idx, esSecundaria) {
  selectEl.innerHTML = `<option value="">-- Selecciona --</option>`;
  const mp = cfg.MultiPrefixes || '';
  const parts = mp.split(',').map(s => s.trim()).filter(Boolean);
  const map = {};
  parts.forEach(p => { const [k, pref] = p.split(':').map(x => x && x.trim()); if (k) map[k] = pref; });
  const pref = map[estado] || '';
  if (!pref) return;
  const base = catalog.filter(r => matchesPrefix(String(r.Código || ''), pref));
  base.forEach(o => selectEl.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${escapeHtml(o.Plan)}</option>`));

  if (esSecundaria && estado === 'multi' && cfg.ExtraMapping) {
    const extraMap = {};
    cfg.ExtraMapping.split(';').map(x => x.trim()).filter(Boolean).forEach(p => { const [k,v] = p.split(':').map(x => x && x.trim()); if (k && v) extraMap[k] = v; });
    const stateSub = findStateSub(sectionName, cfg.Subsection);
    if (stateSub && stateSub.elementos && stateSub.elementos.lines && stateSub.elementos.lines[0]) {
      const principalSel = stateSub.elementos.lines[0].select;
      const principalVal = principalSel ? principalSel.value : '';
      if (principalVal && extraMap[principalVal]) {
        const extraPlan = catalog.find(r => r.Código === extraMap[principalVal]);
        if (extraPlan) selectEl.insertAdjacentHTML('beforeend', `<option value="${extraPlan.Código}">${escapeHtml(extraPlan.Plan)}</option>`);
      }
    }
  }

  if (!Array.from(selectEl.options).some(o => o.value === selectEl.value)) selectEl.value = '';
}

/* ------------ ACTUALIZADO: nueva lógica para construir párrafos en submit (MOVIL) y para mostrar precios/facturación ------------ */

function actualizarMovilSection(sectionName, subName) {
  const stateSub = findStateSub(sectionName, subName);
  if (!stateSub) return;
  const lines = stateSub.elementos.lines || [];
  const detalles = [], preciosInfo = [];

  lines.forEach((ln, i) => {
    const code = ln.select ? ln.select.value : '';
    const plan = findByCode(code);
    detalles.push(plan ? { name: plan.Plan, details: plan.Detalles || '' } : null);
    preciosInfo.push(plan ? buildPromoDescription(plan) : null);
  });

  if (stateSub.elementos.detallesBox) {
    stateSub.elementos.detallesBox.innerHTML = '';
    detalles.forEach((d, i) => {
      const wrapper = document.createElement('div');
      wrapper.style.marginBottom = '8px';
      const hdr = document.createElement('strong');
      hdr.textContent = i === 0 ? 'Línea Principal' : `Adicional ${i}`;
      const txt = document.createElement('div');
      txt.className = 'offer-details-text';
      if (!d) {
        txt.textContent = 'Selecciona un plan para ver detalles.';
      } else {
        const safeName = escapeHtml(d.name);
        const safeDetails = escapeHtml(d.details || '');
        txt.innerHTML = `<b>${safeName}</b><br><br>${safeDetails}`;
      }
      wrapper.appendChild(hdr);
      wrapper.appendChild(txt);
      stateSub.elementos.detallesBox.appendChild(wrapper);
    });
  }

  if (stateSub.elementos.preciosBox) {
    stateSub.elementos.preciosBox.innerHTML = '';
    preciosInfo.forEach((pInfo, i) => {
      if (!pInfo) {
        const line = document.createElement('div');
        line.textContent = `Línea ${i === 0 ? 'Principal' : `Adicional ${i}`}: Selecciona un plan para ver precios.`;
        stateSub.elementos.preciosBox.appendChild(line);
      } else {
        const line = document.createElement('div');
        line.innerHTML = `Línea ${i === 0 ? 'Principal' : `Adicional ${i}`}: ${escapeHtml(pInfo.preciosText)}`;
        stateSub.elementos.preciosBox.appendChild(line);
      }
    });
  }

  if (stateSub.elementos.facturacionBox) {
    let totalDesc = 0, totalSin = 0;
    const rows = preciosInfo.map((pInfo, i) => {
      if (!pInfo) return '';
      const ln = lines[i];
      const code = ln && ln.select ? ln.select.value : '';
      const plan = findByCode(code);
      if (plan) {
        totalDesc += Number(plan.Promo1 || 0);
        totalSin += Number(plan.Valor || 0);
      }
      return `<div>${escapeHtml(pInfo.factText)}</div>`;
    }).join('');
    stateSub.elementos.facturacionBox.innerHTML = rows + `<hr><b>Total con descuento: $${totalDesc}</b><br><b>Total sin descuento: $${totalSin}</b>`;
  }
}

/* ------------ Contrato / SUBMIT (aquí se actualizó la construcción de movilParagraphs y agregado Hogar tag) ------------
*/
function inicializarContrato() {
  (function initPickupToggleInJS() {
    const toggle = document.getElementById('pickupToggle');
    if (!toggle) return;
    toggle.querySelectorAll('button').forEach(btn => {
      btn.addEventListener('click', () => {
        toggle.querySelectorAll('button').forEach(b => {
          b.classList.remove('active');
          b.setAttribute('aria-pressed', 'false');
        });
        btn.classList.add('active');
        btn.setAttribute('aria-pressed', 'true');
      });
    });
  })();

  document.getElementById('contractForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const titular = document.getElementById('nombre').value || '';
    const activeSectionBtn = document.querySelector('.tab-btn.active');
    if (!activeSectionBtn) { alert('Selecciona una sección'); return; }
    const sectionName = activeSectionBtn.dataset.section;
    const activeSubName = state.sections[sectionName].activeSub;
    const stateSub = findStateSub(sectionName, activeSubName);

    let planObj = null;
    if (stateSub) {
      if (stateSub.elementos && stateSub.elementos.lines) {
        for (const ln of stateSub.elementos.lines) {
          const code = ln.select ? ln.select.value : '';
          if (code) { planObj = findByCode(code); break; }
        }
      } else if (stateSub.elementos && stateSub.elementos.mainSelect) {
        planObj = findByCode(stateSub.elementos.mainSelect.value);
      }
    }

    let movilParagraphs = [];
    const planCounts = {};
    let totalSinDescuento = 0;
    let totalConDescuento = 0;
    let anyPromo = false; // bandera para detectar si hubo al menos un Promo1 real

    // NUEVA LÓGICA: Construcción de párrafos para MOVIL (0/1/2 descuentos y portabilidad)
    if (state.sections['Movil']) {
      const subsecs = state.sections['Movil'].subsections || {};
      const allLines = [];
      Object.keys(subsecs).forEach(subName => {
        const st = subsecs[subName];
        const lines = st.elementos && st.elementos.lines ? st.elementos.lines : [];
        lines.forEach(ln => { allLines.push(ln); });
      });

      let selectedIdx = 0;
      for (const ln of allLines) {
        const code = ln.select ? ln.select.value : '';
        const plan = findByCode(code);
        if (!plan) continue;

        const planName = plan.Plan || 'Sin nombre';
        planCounts[planName] = (planCounts[planName] || 0) + 1;
        // total sin descuento siempre suma el valor normal
        totalSinDescuento += Number(plan.Valor || 0);
        // total con descuento suma Promo1 si existe; si no existe suma el valor normal
        if (plan.Promo1 !== '' && plan.Promo1 !== null && plan.Promo1 !== undefined && String(plan.Promo1).trim() !== '') {
          totalConDescuento += Number(plan.Promo1 || 0);
          anyPromo = true;
        } else {
          totalConDescuento += Number(plan.Valor || 0);
        }

        const portaChecked = ln.portaCheckbox ? ln.portaCheckbox.checked : false;
        const numeroPorta = ln.portaNumeroInput ? (ln.portaNumeroInput.value || '').trim() : '';
        const donante = ln.portaDonanteInput ? (ln.portaDonanteInput.value || '').trim() : '';

        const valorPlan = plan.Valor || '';
        const promo1 = (plan.Promo1 === '' || plan.Promo1 === null || plan.Promo1 === undefined) ? '' : String(plan.Promo1);
        const promo2 = (plan.Promo2 === '' || plan.Promo2 === null || plan.Promo2 === undefined) ? '' : String(plan.Promo2);
        const rawM1 = plan.Meses1;
        const rawM2 = plan.Meses2;
        const meses1 = (rawM1 === '' || rawM1 === null || rawM1 === undefined) ? null : (Number(rawM1) && !Number.isNaN(Number(rawM1)) ? Number(rawM1) : (String(rawM1).trim() || null));
        const meses2 = (rawM2 === '' || rawM2 === null || rawM2 === undefined) ? null : (Number(rawM2) && !Number.isNaN(Number(rawM2)) ? Number(rawM2) : (String(rawM2).trim() || null));

        // Construcción del párrafo según cantidad de descuentos
        let prefix = '';
        if (selectedIdx === 0) {
          prefix = `Sr./Sra. ${titular}, Confirmamos el ${plan.Plan},`;
        } else {
          prefix = `Confirmamos siguiente plan, el ${plan.Plan},`;
        }

        let paragraph = '';
        // Sin descuentos
        if (!promo1) {
          paragraph = `${prefix} con valor total de $${valorPlan}.`;
        } else if (promo1 && !promo2) {
          // Un solo descuento
          if (meses1 && typeof meses1 === 'number') {
            // Descuento por N meses
            paragraph = `${prefix} con valor total de $${valorPlan}. Al que se aplicará un valor promocional inicial de $${promo1} durante ${meses1} meses. Finalizado este periodo, a partir del mes ${meses1 + 1}, se aplicará el valor completo sin descuento: $${valorPlan}.`;
          } else {
            // Promo sin meses -> vigencia permanente (o desconocida)
            paragraph = `${prefix} con valor total de $${valorPlan}. Al que se aplicará un valor promocional de $${promo1} por vigencia permanente.`;
          }
        } else {
          // Dos descuentos presentes (promo1 y promo2)
          if (meses1 && typeof meses1 === 'number' && meses2 && typeof meses2 === 'number') {
            paragraph = `${prefix} con valor total de $${valorPlan}. Al que se aplicará un valor promocional inicial de $${promo1} durante ${meses1} meses. A partir del mes ${meses1 + 1}, el valor mensual será de $${promo2} durante ${meses2} meses. Finalizado este periodo, a partir del mes ${computeOffsetMonths(meses2, 2)}, se aplicará el valor completo sin descuento: $${valorPlan}.`;
          } else if (meses1 && typeof meses1 === 'number' && (!meses2 || meses2 === null)) {
            // promo2 sin meses -> promo2 por vigencia permanente tras promo1
            paragraph = `${prefix} con valor total de $${valorPlan}. Al que se aplicará un valor promocional inicial de $${promo1} durante ${meses1} meses. A partir del mes ${meses1 + 1}, se aplicará un valor promocional de $${promo2} por vigencia permanente.`;
          } else {
            // casuística de meses no numéricos: fallback explicativo
            paragraph = `${prefix} con valor total de $${valorPlan}. Existen promociones aplicables: $${promo1}${meses1 ? ` por ${meses1} meses` : ''}${promo2 ? ` y luego $${promo2}${meses2 ? ` por ${meses2} meses` : ' (vigencia permanente)'}` : ''}.`;
          }
        }

        // Agregar portabilidad si corresponde (la inclusión en el texto principal)
        if (portaChecked && numeroPorta && donante) {
          // añadir con conjunción adecuada
          // Si paragraph ya termina con punto, lo usamos antes del punto final
          if (paragraph.endsWith('.')) {
            paragraph = paragraph.slice(0, -1) + `, con portabilidad del número ${numeroPorta} desde la compañía ${donante}.`;
          } else {
            paragraph += ` Con portabilidad del número ${numeroPorta} desde la compañía ${donante}.`;
          }
        }

        movilParagraphs.push(paragraph);
        selectedIdx++;
      }
    }

    const movilText = movilParagraphs.join('\n\n');

    // Generar texto Hogar para la plantilla (TAG <<Hogar>>) usando reglas del usuario
    let hogarTagText = '';
    if (sectionName && String(sectionName).toLowerCase() === 'hogar') {
      // tomar plan desde planObj que ya calculamos arriba (es el plan seleccionado en Hogar)
      const plan = planObj || {};
      const valorPlan = plan.Valor || '';
      const promo1 = (plan.Promo1 === '' || plan.Promo1 === null || plan.Promo1 === undefined) ? '' : String(plan.Promo1);
      const promo2 = (plan.Promo2 === '' || plan.Promo2 === null || plan.Promo2 === undefined) ? '' : String(plan.Promo2);
      const rawM1 = plan.Meses1;
      const rawM2 = plan.Meses2;
      const meses1 = (rawM1 === '' || rawM1 === null || rawM1 === undefined) ? null : (Number(rawM1) && !Number.isNaN(Number(rawM1)) ? Number(rawM1) : (String(rawM1).trim() || null));
      const meses2 = (rawM2 === '' || rawM2 === null || rawM2 === undefined) ? null : (Number(rawM2) && !Number.isNaN(Number(rawM2)) ? Number(rawM2) : (String(rawM2).trim() || null));

      if (!promo1) {
        hogarTagText = `El valor total del plan es de $${valorPlan}`;
      } else if (promo1 && !promo2) {
        if (meses1 && typeof meses1 === 'number') {
          hogarTagText = `El valor total del plan es de $${valorPlan}, al que se aplicará un valor promocional inicial de $${promo1} durante ${meses1} meses. Finalizado este periodo, a partir del mes ${meses1 + 1}, se aplicará el valor completo sin descuento: $${valorPlan}.`;
        } else {
          hogarTagText = `El valor total del plan es de $${valorPlan}, al que se aplicará un valor promocional de $${promo1} por vigencia permanente.`;
        }
      } else {
        // dos descuentos
        if (meses1 && typeof meses1 === 'number' && meses2 && typeof meses2 === 'number') {
          hogarTagText = `El valor total del plan es de $${valorPlan}, al que se aplicará un valor promocional inicial de $${promo1} durante ${meses1} meses. A partir del mes ${meses1 + 1}, el valor mensual será de $${promo2} durante ${meses2} meses. Finalizado este periodo, a partir del mes ${computeOffsetMonths(meses2, 2)}, se aplicará el valor completo sin descuento: $${valorPlan}.`;
        } else if (meses1 && typeof meses1 === 'number' && (!meses2 || meses2 === null)) {
          hogarTagText = `El valor total del plan es de $${valorPlan}, al que se aplicará un valor promocional inicial de $${promo1} durante ${meses1} meses. A partir del mes ${meses1 + 1}, el valor mensual será de $${promo2} por vigencia permanente.`;
        } else {
          hogarTagText = `El valor total del plan es de $${valorPlan}, al que se aplicarán promociones: $${promo1}${meses1 ? ` durante ${meses1} meses` : ''} y $${promo2}${meses2 ? ` durante ${meses2} meses` : ' (vigencia permanente)'}; luego aplicará el valor normal $${valorPlan}.`;
        }
      }
    }

    let hasAnyPortability = false;
    if (state.sections['Movil']) {
      const subsecs = state.sections['Movil'].subsections || {};
      outerLoop:
      for (const subName of Object.keys(subsecs)) {
        const st = subsecs[subName];
        const lines = st.elementos && st.elementos.lines ? st.elementos.lines : [];
        for (const ln of lines) {
          const code = ln.select ? ln.select.value : '';
          if (!code) continue;
          const portaChecked = ln.portaCheckbox ? ln.portaCheckbox.checked : false;
          if (portaChecked) { hasAnyPortability = true; break outerLoop; }
        }
      }
    }

    let condicionText = ' ';
    if (hasAnyPortability) {
      condicionText = `¿Autoriza usted mediante esta grabación a Pacífico Cable SPA a solicitar al OAP toda información necesaria para activar el proceso? Necesito que me indique su número telefónico actual, la compañía donante, su RUT y su nombre completo.\n\nLa portabilidad solo aplica al número telefónico. Su compañía actual podría cobrar por servicios pendientes. El cambio se realiza entre 03:00 y 05:00 AM, con posible breve interrupción. En caso de retracto, puede realizarlo hasta las 20:00 horas del día en que se active el servicio.\n`;
    }

    // NOC en Movil (se mantiene)
    let nocTextMovil = '';
    if (state.sections['Movil'] && state.sections['Movil'].activeSub) {
      const activeMovilSub = state.sections['Movil'].activeSub;
      if (String(activeMovilSub).toLowerCase() === 'nuevo') {
        nocTextMovil = `En Mundo, nuestros servicios tienen el cobro por mes adelantado con seis ciclos de facturación distintos con fecha de inicio 1, 5, 10, 15, 20 y 25 de cada mes. La primera boleta se emitirá en el ciclo más cercano a la activación de los servicios, con 20 días continuos de plazo para pagar. Si no se paga 5 días después, el servicio se suspende y la reposición cuesta $2.500.`;
      } else if (String(activeMovilSub).toLowerCase() === 'cartera') {
        const cicloVal = document.getElementById('ciclo') ? (document.getElementById('ciclo').value || '') : '';
        nocTextMovil = `Nuestros servicios se facturan por mes adelantado y se acoplan a su actual ciclo de facturación ${cicloVal}. Puede aplicarse un cobro proporcional el día de la activación si corresponde.`;
      } else {
        nocTextMovil = '';
      }
    }

    // OBTEN: texto según modo de pickup (Sucursal / Domicilio)
    let obtenText = '';
    const pickupToggle = document.getElementById('pickupToggle');
    let pickupMode = 'Sucursal';
    if (pickupToggle) {
      const activeBtn = pickupToggle.querySelector('button.active');
      if (activeBtn && activeBtn.dataset && activeBtn.dataset.val) pickupMode = activeBtn.dataset.val;
    }
    if (pickupMode === 'Sucursal') {
      const suc = document.getElementById('sucursal') ? (document.getElementById('sucursal').value || '') : '';
      obtenText = `En la sucursal seleccionada por usted ${suc}. El retiro y activación de su Sim Card puede realizarlo a partir del día hábil siguiente (24 horas).`;
    } else if (pickupMode === 'Domicilio') {
      const dir = document.getElementById('direccion') ? (document.getElementById('direccion').value || '') : '';
      obtenText = `La tarjeta SIM será enviada a su dirección ${dir}, en un plazo de 2 a 5 días hábiles, una vez recibida debe activarla siguiendo las indicaciones entregadas junto con su Sim Card. Si tiene dudas o consultas puede realizarlas al 6009100100 o al 442160800 opción móvil. (Activación Opción 5)`;
    }

    // ALL: resumen por plan y totales calculados
    let allText = '';
    const planEntries = Object.keys(planCounts).map(planName => {
      const count = planCounts[planName];
      const plural = count === 1 ? 'linea' : 'lineas';
      return `${count} ${plural} con el ${planName}`;
    });
    if (planEntries.length > 0) {
      const listaPlanes = planEntries.join(', ');
      // mostrar la parte de "con descuento" sólo si hubo al menos un promo aplicado
      if (anyPromo && totalConDescuento !== totalSinDescuento) {
        allText = `Usted está contratando ${listaPlanes}, con valor total de $${totalSinDescuento} y con descuento quedaría en $${totalConDescuento}.`;
      } else {
        allText = `Usted está contratando ${listaPlanes}, con valor total de $${totalSinDescuento}`;
      }
    } else {
      allText = '';
    }

    // PREPARAR DATOS Y ELEGIR PLANTILLA
    let templateFile = 'contrato_template.docx';
    const sectionNameLower = sectionName ? String(sectionName).toLowerCase() : '';
    let data = {};

    const ejecutivoNameForTemplate = (typeof window !== 'undefined' && window.Ejecutivo) ? String(window.Ejecutivo) : '';

    if (sectionNameLower === 'hogar') {
      // Plantilla específica para Hogar (contrato_template2.docx)
      templateFile = 'contrato_template2.docx';
      const plan = planObj || {};

      // NOC para Hogar: debe comportarse igual que Movil cuando la subsección activa en Hogar es nuevo/cartera
      let nocTextHogar = '';
      if (state.sections['Hogar'] && state.sections['Hogar'].activeSub) {
        const activeHogarSub = state.sections['Hogar'].activeSub;
        if (String(activeHogarSub).toLowerCase() === 'nuevo') {
          nocTextHogar = `En Mundo, nuestros servicios tienen el cobro por mes adelantado con seis ciclos de facturación distintos con fecha de inicio 1, 5, 10, 15, 20 y 25 de cada mes. La primera boleta se emitirá en el ciclo más cercano a la activación de los servicios, con 20 días continuos de plazo para pagar. Si no se paga 5 días después, el servicio se suspende y la reposición cuesta $2.500.`;
        } else if (String(activeHogarSub).toLowerCase() === 'cartera') {
          const cicloVal = document.getElementById('ciclo') ? (document.getElementById('ciclo').value || '') : '';
          nocTextHogar = `Nuestros servicios se facturan por mes adelantado y se acoplan a su actual ciclo de facturación ${cicloVal}. Puede aplicarse un cobro proporcional el día de la activación si corresponde.`;
        } else {
          nocTextHogar = '';
        }
      }

      // PORTA / PORTA2: tomar desde el stateSub actual en Hogar (si existe)
      let portaText = '';
      let porta2Text = '';
      try {
        if (stateSub && stateSub.elementos && stateSub.elementos.portaCheckbox && stateSub.elementos.portaCheckbox.checked) {
          // Texto exacto solicitado para <<PORTA>>
          portaText = `¿autoriza usted mediante esta grabación a Pacífico Cable SPA a solicitar al OAP toda información necesaria para activar el proceso? Necesito que me indique su número telefónico actual, la compañía donante, su RUT y su nombre completo.\nLe informo que las llamadas a números internacionales y líneas 700 están bloqueadas, aunque usted puede realizarlas sin costo usando plataformas como Skype, WhatsApp, ZOOM o Meet. En caso de corte de luz o suspensión por no pago, el servicio telefónico quedará inhabilitado.`;
          const num = (stateSub.elementos.portaNumeroInput && stateSub.elementos.portaNumeroInput.value) ? stateSub.elementos.portaNumeroInput.value.trim() : '';
          const comp = (stateSub.elementos.portaDonanteInput && stateSub.elementos.portaDonanteInput.value) ? stateSub.elementos.portaDonanteInput.value.trim() : '';
          porta2Text = `Portabilidad para el número fijo ${num} actualmente con la compañía ${comp}.`;
        } else {
          portaText = '';
          porta2Text = '';
        }
      } catch (e) {
        portaText = '';
        porta2Text = '';
      }

      // Se calculan variantes de meses para plantillas que lo requieran
      let meses1Minus1 = '';
      const rawM1 = plan.Meses1;
      if (typeof rawM1 === 'number') {
        meses1Minus1 = rawM1 + 1;
      } else if (typeof rawM1 === 'string' && rawM1.trim() !== '' && !Number.isNaN(Number(rawM1))) {
        meses1Minus1 = Number(rawM1) + 1;
      } else {
        meses1Minus1 = '';
      }

      let meses2Plus1 = '';
      const rawM2 = plan.Meses2;
      // Aquí aplicamos la regla solicitada: (meses2 * 2) + 1
      if (typeof rawM2 === 'number') {
        meses2Plus1 = computeOffsetMonths(rawM2, 2);
      } else if (typeof rawM2 === 'string' && rawM2.trim() !== '' && !Number.isNaN(Number(rawM2))) {
        meses2Plus1 = computeOffsetMonths(Number(rawM2), 2);
      } else {
        meses2Plus1 = '';
      }

      data = {
        'NOMBRE': titular,
        'PLAN': plan.Plan || '',
        'DIRECCION': document.getElementById('direccion').value || '',
        'VALOR': plan.Valor || '',
        'PROMO1': plan.Promo1 || '',
        'MESES1': plan.Meses1 || '',
        'MESES1-1': meses1Minus1,
        'MESES2+1': meses2Plus1,
        'PROMO2': plan.Promo2 || '',
        'MESES2': plan.Meses2 || '',
        'DETALLES': plan.Detalles || '',
        'FECHA': document.getElementById('fecha').value || '',
        'EJECUTIVO': ejecutivoNameForTemplate,
        // Campos nuevos solicitados
        'PORTA': portaText,
        'PORTA2': porta2Text,
        'NOC': nocTextHogar,
        // TAG solicitado para Hogar
        'Hogar': hogarTagText
      };
    } else {
      templateFile = 'contrato_template.docx';
      data = {
        NOMBRE: titular,
        DIRECCION: document.getElementById('direccion').value,
        SUCURSAL: document.getElementById('sucursal').value,
        PLAN: planObj ? planObj.Plan : '',
        VALOR_PLAN: planObj ? planObj.Valor : '',
        VALOR_PROMO: planObj ? planObj.Promo1 : '',
        VALOR_PROMO2: planObj ? planObj.Promo2 : '',
        DURACION: planObj ? planObj.Meses1 : '',
        CICLO: document.getElementById('ciclo').value,
        FECHA: document.getElementById('fecha').value,
        MOVIL: movilText,
        CONDICION: condicionText,
        NOC: nocTextMovil,
        OBTEN: obtenText,
        ALL: allText,
        'EJECUTIVO': ejecutivoNameForTemplate
      };
    }

    try {
      const content = await loadFile(templateFile);
      const zip = new PizZip(content);
      const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true, delimiters: { start: '<<', end: '>>' } });
      doc.render(data);
      const blob = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

      const safeTitular = (titular || 'Contrato').replace(/[\\\/:\*\?"<>\|]+/g, '').trim() || 'Contrato';
      const filename = `${safeTitular}.docx`;

      if (typeof saveContrato === 'function') {
        await saveContrato(blob, filename);
      } else {
        try {
          const tmp = indexedDB.open('ContratoDB', 1);
          tmp.onsuccess = () => {
            const db = tmp.result;
            const tx = db.transaction('Contratos', 'readwrite');
            const store = tx.objectStore('Contratos');
            store.put(blob, 'ultimoContrato');
          };
        } catch (err) {
          console.warn('No se pudo usar saveContrato, y fallback falló:', err);
        }
      }

        // Envío silencioso a Google Forms con ejecutivo y nombre del cliente
      // Se envía de forma asíncrona y no bloquea la generación ni la UI.
      try {
        const ejecutivoToSend = (typeof window !== 'undefined' && window.Ejecutivo) ? String(window.Ejecutivo) : '';
        const clienteToSend = titular || '';
        // Fire-and-forget, no-cors para evitar bloqueos por CORS en navegadores
        sendToGoogleForm(ejecutivoToSend, clienteToSend).catch(err => {
          // no mostrar error al usuario; solo log
          console.warn('Error enviando a Google Forms (silencioso):', err);
        });
      } catch (err) {
        console.warn('Error preparando envío a Google Forms:', err);
      }

      document.getElementById('preview').innerHTML = '<p>Contrato generado y guardado. Pulsa “Visualizar contrato”.</p>';
    } catch (err) {
      console.error('Error generando contrato:', err);
      showMessage('Error generando contrato. Revisa la consola.');
    }
  });



  document.getElementById('visualizarButton').addEventListener('click', async () => {
    try {
      const stored = await getContrato();
      if (!stored || !stored.blob) { alert('No hay contrato generado.'); return; }
      const blob = stored.blob;
      const filenameDocx = stored.filename || 'Contrato.docx';
      const archivo = new File([blob], filenameDocx, { type: blob.type });
      const container = document.getElementById('preview');
      container.innerHTML = '';
      await window.docx.renderAsync(archivo, container);
      const imgs = container.querySelectorAll('img'); if (imgs.length > 1) imgs[1].remove();
      const hdr = container.querySelector('div'); if (hdr) Object.assign(hdr.style, { margin: '0', padding: '0', float: 'none', display: 'block' });
      const first = container.firstElementChild; if (first) Object.assign(first.style, { margin: '0', padding: '0' });
      Object.assign(container.style, { margin: '0', padding: '0' });
      await new Promise(requestAnimationFrame);

      const capture = document.getElementById('pdf-capture');
      const allImgs = capture.querySelectorAll('img'); allImgs.forEach(img => (img.crossOrigin = 'anonymous'));
      await Promise.all(Array.from(allImgs).map(img => new Promise(resolve => { if (img.complete) return resolve(); img.onload = resolve; img.onerror = resolve; })));
      const canvas = await html2canvas(capture, { scale: 2, useCORS: true, allowTaint: false, scrollX: 0, scrollY: -window.scrollY, width: capture.offsetWidth, height: capture.scrollHeight });
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ unit: 'mm', format: 'letter', orientation: 'portrait' });
      const pageW = pdf.internal.pageSize.getWidth(); const margin = 5;
      const pdfW = pageW - margin * 2; const pxPerMm = canvas.width / pdfW;
      const pageH = pdf.internal.pageSize.getHeight() - margin * 2; const pagePxH = Math.floor(pageH * pxPerMm);
      let renderedH = 0, pageCount = 0;
      while (renderedH < canvas.height) {
        const fragH = Math.min(pagePxH, canvas.height - renderedH);
        const pageCanvas = document.createElement('canvas'); pageCanvas.width = canvas.width; pageCanvas.height = fragH;
        pageCanvas.getContext('2d').drawImage(canvas, 0, renderedH, canvas.width, fragH, 0, 0, canvas.width, fragH);
        const fragImg = pageCanvas.toDataURL('image/jpeg', 1.0);
        if (pageCount > 0) pdf.addPage();
        pdf.addImage(fragImg, 'JPEG', margin, margin, pdfW, (fragH / canvas.width) * pdfW);
        renderedH += fragH; pageCount++;
      }
      const baseName = (filenameDocx || 'Contrato.docx').replace(/\.docx$/i, '');
      pdf.save(`${baseName}.pdf`);
    } catch (err) {
      console.error('Error exportando PDF:', err);
      showMessage('Error exportando PDF. Revisa la consola.');
    }
  });
}

/* ------------ Nuevo módulo: envío silencioso a Google Forms ------------
   Requisitos del usuario:
   - Al apretar "Generar contrato" (ya integrado arriba), tomar el nombre del Ejecutivo y el nombre del cliente (input 'nombre')
     y enviarlos a este Google Form de forma oculta:
     https://docs.google.com/forms/d/e/1FAIpQLSdQL8_CkZwBVZq6pbWizrCnBoRhNoWOleuWwNH_kM4QR5SMuQ/viewform?usp=dialog
   - Campos a enviar:
     entry.775437783 -> nombre del ejecutivo
     entry.315589411 -> nombre del cliente
   - El form no debe mostrarse nunca. El envío debe ser silencioso, no interferir con la UI ni con la generación del .docx.
   Implementación:
   - Se realiza un POST a la endpoint "formResponse" del form. Se usa fetch con mode: 'no-cors' para evitar bloqueos CORS.
   - La función es "fire-and-forget": se ejecuta asíncronamente y no muestra errores al usuario en caso de fallo.
*/
async function sendToGoogleForm(ejecutivoName, clienteName) {
  // URL base de envío (formResponse)
  const formBase = 'https://docs.google.com/forms/d/e/1FAIpQLSdQL8_CkZwBVZq6pbWizrCnBoRhNoWOleuWwNH_kM4QR5SMuQ/formResponse';
  // Construir FormData con las entradas solicitadas
  const fd = new FormData();
  // Campos provistos por el usuario:
  // entry.775437783 -> nombre del ejecutivo
  // entry.315589411 -> nombre del cliente
  fd.append('entry.775437783', ejecutivoName || '');
  fd.append('entry.315589411', clienteName || '');
  // Opcional: agregar un timestamp (no obligatorio)
  fd.append('timestamp', new Date().toISOString());

  // Intentar enviar por fetch. Usar no-cors para no depender de CORS del destino.
  // Nota: cuando mode:'no-cors' la respuesta será opaque y no se podrá inspeccionar el resultado.
  try {
    await fetch(formBase, {
      method: 'POST',
      mode: 'no-cors',
      body: fd,
      // cache: 'no-cache' // opcional
    });
    // No hacemos nada con la respuesta.
    return true;
  } catch (err) {
    // En algunos navegadores el fetch puede fallar; lo registramos pero no interrumpimos la UX.
    console.warn('sendToGoogleForm failed:', err);
    return false;
  }
}


/* ------------ Util ------------
*/
function loadFile(url) { return new Promise((resolve, reject) => { window.PizZipUtils.getBinaryContent(url, (err, data) => err ? reject(err) : resolve(data)); }); }
function escapeHtml(text) { if (text === null || text === undefined) return ''; return String(text).replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch])); }

/* ------------ Ejecutivo UI (sin cambios funcionales respecto a tu original) ------------
*/
function initEjecutivoUI() {
  const gearBtn = document.getElementById('gearBtn');
  const modal = document.getElementById('ejecutivoModal');
  const ejecutivoInput = document.getElementById('ejecutivoInput');
  const editBtn = document.getElementById('editEjecutivoBtn');
  const deleteBtn = document.getElementById('deleteEjecutivoBtn');
  const cancelBtn = document.getElementById('cancelEjecutivoBtn');
  const acceptBtn = document.getElementById('acceptEjecutivoBtn');
  const nameSpan = document.getElementById('ejecutivoName');

  window.Ejecutivo = '';

  (async function loadAndRender() {
    try {
      const stored = (typeof getEjecutivo === 'function') ? await getEjecutivo() : '';
      window.Ejecutivo = stored || '';
      renderEjecutivoName();
    } catch (err) {
      console.error('No se pudo leer Ejecutivo desde IndexedDB', err);
    }
  })();

  function renderEjecutivoName() {
    if (!nameSpan) return;
    if (window.Ejecutivo && String(window.Ejecutivo).trim() !== '') {
      nameSpan.textContent = String(window.Ejecutivo);
      nameSpan.title = `Ejecutivo: ${window.Ejecutivo}`;
    } else {
      nameSpan.textContent = '';
      nameSpan.title = '';
    }
  }

  function openModal() {
    if (!modal) return;
    ejecutivoInput.value = window.Ejecutivo || '';
    ejecutivoInput.setAttribute('readonly', 'readonly');
    if (!window.Ejecutivo) {
      ejecutivoInput.removeAttribute('readonly');
      ejecutivoInput.focus();
    }
    modal.hidden = false;
    acceptBtn.focus();
  }

  function closeModal() {
    if (!modal) return;
    modal.hidden = true;
  }

  gearBtn && gearBtn.addEventListener('click', (e) => {
    openModal();
  });

  editBtn && editBtn.addEventListener('click', () => {
    ejecutivoInput.removeAttribute('readonly');
    ejecutivoInput.focus();
    const val = ejecutivoInput.value;
    ejecutivoInput.value = '';
    ejecutivoInput.value = val;
  });

  deleteBtn && deleteBtn.addEventListener('click', async () => {
    ejecutivoInput.removeAttribute('readonly');
    ejecutivoInput.value = '';
    ejecutivoInput.focus();
  });

  cancelBtn && cancelBtn.addEventListener('click', (e) => {
    e.preventDefault();
    ejecutivoInput.value = window.Ejecutivo || '';
    ejecutivoInput.setAttribute('readonly', 'readonly');
    closeModal();
  });

  acceptBtn && acceptBtn.addEventListener('click', async (e) => {
    e.preventDefault();
    const newName = (ejecutivoInput.value || '').trim();
    try {
      if (!newName) {
        if (typeof deleteEjecutivo === 'function') await deleteEjecutivo();
        window.Ejecutivo = '';
      } else {
        if (typeof saveEjecutivo === 'function') await saveEjecutivo(newName);
        window.Ejecutivo = newName;
      }
      renderEjecutivoName();
      ejecutivoInput.setAttribute('readonly', 'readonly');
      closeModal();
    } catch (err) {
      console.error('Error guardando/eliminando Ejecutivo', err);
      showMessage('No se pudo guardar el nombre del Ejecutivo. Revisa la consola.', true, 6000);
    }
  });

  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape') {
      const modalVisible = modal && !modal.hidden;
      if (modalVisible) {
        cancelBtn && cancelBtn.click();
      }
    }
  });

  modal && modal.addEventListener('click', (ev) => {
    if (ev.target === modal) {
      cancelBtn && cancelBtn.click();
    }
  });
}