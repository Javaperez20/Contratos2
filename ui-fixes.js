// ui-fixes.js
// - Mueve el nodo .porta-fields debajo del <select> dentro de la tarjeta cuando el checkbox 'porta' se activa.
// - Añade/retira clase .porta-visible en la tarjeta para que CSS muestre los campos.
// - Asegura que al cerrar la porta los campos se oculten y vuelvan a su contenedor original.
// - Detecta tarjetas ya existentes y tarjetas añadidas dinámicamente.
//
// Instrucciones: incluye este script justo después de main.js en tu HTML:
// <script src="db.js"></script>
// <script src="main.js"></script>
// <script src="ui-fixes.js"></script>

(function () {
  'use strict';

  // Observa #tarifario-root para inicializar tarjetas recién creadas
  const root = document.getElementById('tarifario-root');

  if (!root) {
    console.warn('ui-fixes: #tarifario-root no encontrado — asegúrate de incluir este script después de main.js');
  }

  // Inicializa una tarjeta: localiza elementos y enlaza comportamiento
  function initCard(card) {
    if (!card || card._uiFixed) return;
    card._uiFixed = true;

    const select = card.querySelector('select');
    const portaContainer = card.querySelector('.porta-container');
    const portaFields = card.querySelector('.porta-fields');
    const portaCheckbox = card.querySelector('.porta-checkbox');
    const deleteBtn = card.querySelector('.btn-remove');

    // If no portaFields or select, nothing to do
    if (!select || !portaFields || !portaContainer || !portaCheckbox) {
      // still move delete button if present
      if (deleteBtn) {
        deleteBtn.style.left = '12px';
        deleteBtn.style.right = 'auto';
        deleteBtn.style.top = 'auto';
        deleteBtn.style.bottom = '12px';
      }
      return;
    }

    // Ensure delete button style: bottom-left
    if (deleteBtn) {
      deleteBtn.style.left = '12px';
      deleteBtn.style.right = 'auto';
      deleteBtn.style.top = 'auto';
      deleteBtn.style.bottom = '12px';
    }

    // Keep a reference where portaFields originally live in case we want to restore
    portaFields._originalParent = portaFields.parentNode;
    portaFields._originalNext = portaFields.nextSibling;

    // helper to show porta fields under select
    function showPortaFields() {
      // move portaFields directly after select if not already there
      if (select.parentNode && portaFields.parentNode !== select.parentNode) {
        select.parentNode.insertBefore(portaFields, select.nextSibling);
      }
      // mark visible class on card for CSS selector
      card.classList.add('porta-visible');
      if (portaContainer) portaContainer.classList.add('porta-visible');
      portaFields.style.display = 'flex';
    }

    function hidePortaFields() {
      // hide visually
      portaFields.style.display = 'none';
      card.classList.remove('porta-visible');
      if (portaContainer) portaContainer.classList.remove('porta-visible');
      // move back to original parent (defensive)
      if (portaFields._originalParent && portaFields.parentNode !== portaFields._originalParent) {
        if (portaFields._originalNext) portaFields._originalParent.insertBefore(portaFields, portaFields._originalNext);
        else portaFields._originalParent.appendChild(portaFields);
      }
    }

    // Set initial visibility according to checkbox state
    if (portaCheckbox.checked) showPortaFields();
    else hidePortaFields();

    // Toggle on change
    portaCheckbox.addEventListener('change', () => {
      if (portaCheckbox.checked) showPortaFields();
      else hidePortaFields();
    });

    // Defensive: if user types into inputs and then delete card, values remain local; no further action required.

    // If select changes (user picks plan) keep porta fields positioned under select (no extra action needed)
  }

  // Initialize existing cards
  function initExisting() {
    const cards = root ? root.querySelectorAll('.movil-line, .movil-line-card') : [];
    cards.forEach(initCard);
  }

  // Observe for dynamic additions
  if (root) {
    const mo = new MutationObserver(muts => {
      for (const m of muts) {
        if (m.type === 'childList' && m.addedNodes.length) {
          m.addedNodes.forEach(node => {
            if (!(node instanceof HTMLElement)) return;
            // if a subtree added, search for cards
            if (node.matches && (node.matches('.movil-line') || node.matches('.movil-line-card'))) {
              initCard(node);
            } else {
              node.querySelectorAll && node.querySelectorAll('.movil-line, .movil-line-card').forEach(initCard);
            }
          });
        }
      }
    });
    mo.observe(root, { childList: true, subtree: true });
  }

  // Run initial
  // Wait a tick to let main.js render initial UI
  window.requestAnimationFrame(() => {
    initExisting();
    // also init after a small delay to catch asynchronous renders
    setTimeout(initExisting, 300);
  });

})();