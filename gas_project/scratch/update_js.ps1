$mockJson = Get-Content 'gas_project\scratch\mock_catalog.json' -Raw -Encoding UTF8

$jsContent = Get-Content 'gas_project\JavaScript.html' -Raw -Encoding UTF8

# 1. Update State & DOM ready
$oldState = @'
  // Global Application State
  const state = {
    items: [],
    filteredItems: [],
    masterCatalog: [], // Master DCI catalog from Sheet1
    logs: [],
    selectedRow: null,
    loading: false,
    activeTab: 'tab-inventory'
  };

  // On DOM Ready
  document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    setupAutocomplete();
    loadInventory();
  });
'@

$newState = @'
  // Global Application State
  const state = {
    items: [],
    filteredItems: [],
    masterCatalog: [], // Master DCI catalog from Sheet1
    logs: [],
    selectedRow: null,
    loading: false,
    activeTab: 'tab-inventory',
    activeDropdown: null, // 'code' or 'desc' or null
    focusedIndex: -1
  };

  // On DOM Ready
  document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    setupComboboxEvents();
    loadInventory();
  });
'@

if ($jsContent.Contains($oldState)) {
    $jsContent = $jsContent.Replace($oldState, $newState)
    Write-Output "State updated."
} else {
    Write-Warning "Old state block not found."
}

# 2. Add DEFAULT_MOCK_DCI_CATALOG and update loadInventory mock
$catalogDecl = "  // Default Master DCI Catalog (291 items from Sheet1)`r`n  const DEFAULT_MOCK_DCI_CATALOG = " + $mockJson.Trim() + ";`r`n`r`n  // Load Inventory from Google Apps Script Backend"

$oldLoadHeader = "  // Load Inventory from Google Apps Script Backend"
if ($jsContent.Contains($oldLoadHeader)) {
    $jsContent = $jsContent.Replace($oldLoadHeader, $catalogDecl)
    Write-Output "DEFAULT_MOCK_DCI_CATALOG inserted."
} else {
    Write-Warning "Load header not found."
}

# Replace the tiny 18 mock items in loadInventory with DEFAULT_MOCK_DCI_CATALOG
$patternOldMock = '(?s)catalog:\s*\[\s*\{ code: ''DMR000002''.*?\{ code: ''DMDON001151'', desc: ''\(DON\) CARBAMAZEPINE 200MG TABLET'' \}\s*\]'
if ($jsContent -match $patternOldMock) {
    $jsContent = $jsContent -replace $patternOldMock, 'catalog: DEFAULT_MOCK_DCI_CATALOG'
    Write-Output "Mock catalog reference updated."
} else {
    Write-Warning "Old mock catalog array not matched."
}

# 3. Update onInventoryLoaded
$oldLoaded = @'
    state.items = response.items || [];
    if (response.catalog && Array.isArray(response.catalog)) {
      state.masterCatalog = response.catalog;
    }
    renderStats();
    applyFilter();
    populateCatalogDatalists();
    setupAutocomplete();
'@

$newLoaded = @'
    state.items = response.items || [];
    if (response.catalog && Array.isArray(response.catalog) && response.catalog.length > 0) {
      state.masterCatalog = response.catalog;
    } else if (state.masterCatalog.length === 0) {
      state.masterCatalog = DEFAULT_MOCK_DCI_CATALOG;
    }
    renderStats();
    applyFilter();
    populateCatalogDatalists();
    setupComboboxEvents();
'@

if ($jsContent.Contains($oldLoaded)) {
    $jsContent = $jsContent.Replace($oldLoaded, $newLoaded)
    Write-Output "onInventoryLoaded updated."
} else {
    Write-Warning "onInventoryLoaded block not found."
}

# 4. Update submitAddNew clear section
$oldClearAddNew = @'
            // Clear form
            document.getElementById('add-code').value = '';
            document.getElementById('add-desc').value = '';
            document.getElementById('add-brand').value = '';
            document.getElementById('add-qty').value = '';
            document.getElementById('add-exp').value = '';
            document.getElementById('add-price').value = '';
            document.getElementById('add-batch-preview').style.display = 'none';
'@

$newClearAddNew = @'
            // Clear form
            resetDciSelection();
            document.getElementById('add-brand').value = '';
            document.getElementById('add-qty').value = '';
            document.getElementById('add-exp').value = '';
            document.getElementById('add-price').value = '';
            document.getElementById('add-batch-preview').style.display = 'none';
'@

if ($jsContent.Contains($oldClearAddNew)) {
    $jsContent = $jsContent.Replace($oldClearAddNew, $newClearAddNew)
    Write-Output "submitAddNew clear updated."
} else {
    Write-Warning "Old clear block in submitAddNew not found."
}

# 5. Replace setupAutocomplete through end of script with the complete Combobox system
$oldTail = @'
  function setupAutocomplete() {
    const codeInput = document.getElementById('add-code');
    const descInput = document.getElementById('add-desc');
    if (!codeInput || !descInput) return;

    codeInput.removeEventListener('input', handleCodeChange);
    codeInput.removeEventListener('change', handleCodeChange);
    descInput.removeEventListener('input', handleDescChange);
    descInput.removeEventListener('change', handleDescChange);

    codeInput.addEventListener('input', handleCodeChange);
    codeInput.addEventListener('change', handleCodeChange);
    descInput.addEventListener('input', handleDescChange);
    descInput.addEventListener('change', handleDescChange);
  }

  function handleCodeChange() {
    const codeInput = document.getElementById('add-code');
    const descInput = document.getElementById('add-desc');
    if (!codeInput || !descInput) return;

    const val = codeInput.value.trim().toUpperCase();
    if (!val) {
      onAddCodeOrDescChange();
      return;
    }

    // Auto-fill generic description from Sheet1 master catalog if matched
    if (state.masterCatalog && state.masterCatalog.length > 0) {
      const match = state.masterCatalog.find(item => item.code.toUpperCase() === val);
      if (match && match.desc) {
        descInput.value = match.desc;
      }
    }

    onAddCodeOrDescChange();
  }

  function handleDescChange() {
    const codeInput = document.getElementById('add-code');
    const descInput = document.getElementById('add-desc');
    if (!codeInput || !descInput) return;

    const val = descInput.value.trim().toUpperCase();
    if (!val) {
      onAddCodeOrDescChange();
      return;
    }

    // Auto-fill DCI code from Sheet1 master catalog if matched
    if (state.masterCatalog && state.masterCatalog.length > 0) {
      const match = state.masterCatalog.find(item => item.desc.toUpperCase() === val);
      if (match && match.code) {
        codeInput.value = match.code;
      }
    }

    onAddCodeOrDescChange();
  }

  function clearAddInput(inputId) {
    const input = document.getElementById(inputId);
    if (input) {
      input.value = '';
      input.focus();
      onAddCodeOrDescChange();
    }
  }

  function populateCatalogDatalists() {
    const codeList = document.getElementById('dci-codes-datalist');
    const descList = document.getElementById('generic-names-datalist');
    const badge = document.getElementById('catalog-count-badge');

    if (!state.masterCatalog || state.masterCatalog.length === 0) return;

    if (codeList) {
      codeList.innerHTML = state.masterCatalog.map(item => 
        `<option value="${escapeHtml(item.code)}">${escapeHtml(item.desc)}</option>`
      ).join('');
    }

    if (descList) {
      descList.innerHTML = state.masterCatalog.map(item => 
        `<option value="${escapeHtml(item.desc)}">${escapeHtml(item.code)}</option>`
      ).join('');
    }

    if (badge) {
      badge.textContent = `📋 ${state.masterCatalog.length} in Sheet1`;
      badge.style.display = 'inline-block';
    }
  }

  function escapeHtml(text) {
    if (!text) return '';
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
</script>
'@

$newTail = @'
  // --- MASTER DCI FORMULARY & SEARCHABLE COMBOBOX SYSTEM ---
  function setupComboboxEvents() {
    const codeInput = document.getElementById('add-code');
    const descInput = document.getElementById('add-desc');
    if (!codeInput || !descInput) return;

    // Reset previous listeners
    codeInput.oninput = handleCodeInput;
    codeInput.onfocus = () => {
      renderDropdownItems('code', codeInput.value.trim());
      openDropdown('code');
    };
    codeInput.onkeydown = (e) => handleComboboxKeydown(e, 'code');

    descInput.oninput = handleDescInput;
    descInput.onfocus = () => {
      renderDropdownItems('desc', descInput.value.trim());
      openDropdown('desc');
    };
    descInput.onkeydown = (e) => handleComboboxKeydown(e, 'desc');

    // Global outside click listener to close dropdowns
    document.addEventListener('click', (e) => {
      const wrapperCode = document.getElementById('combobox-code-wrapper');
      const wrapperDesc = document.getElementById('combobox-desc-wrapper');
      const inCode = wrapperCode && wrapperCode.contains(e.target);
      const inDesc = wrapperDesc && wrapperDesc.contains(e.target);

      if (!inCode && !inDesc) {
        closeAllDropdowns();
      }
    });
  }

  function handleCodeInput(e) {
    const codeInput = document.getElementById('add-code');
    const descInput = document.getElementById('add-desc');
    const q = codeInput.value.trim();
    updateClearBtn('add-code', q);
    renderDropdownItems('code', q);
    openDropdown('code');

    // Bi-directional Auto-Sync: Exact match on DCI Code immediately sets Generic Description
    if (q && state.masterCatalog && state.masterCatalog.length > 0) {
      const match = state.masterCatalog.find(item => item.code.toUpperCase() === q.toUpperCase());
      if (match && match.desc) {
        descInput.value = match.desc;
        updateClearBtn('add-desc', match.desc);
      }
    }

    onAddCodeOrDescChange();
  }

  function handleDescInput(e) {
    const codeInput = document.getElementById('add-code');
    const descInput = document.getElementById('add-desc');
    const q = descInput.value.trim();
    updateClearBtn('add-desc', q);
    renderDropdownItems('desc', q);
    openDropdown('desc');

    // Bi-directional Auto-Sync: Exact match on Generic Description immediately sets DCI Code
    if (q && state.masterCatalog && state.masterCatalog.length > 0) {
      const match = state.masterCatalog.find(item => item.desc.toUpperCase() === q.toUpperCase());
      if (match && match.code) {
        codeInput.value = match.code;
        updateClearBtn('add-code', match.code);
      }
    }

    onAddCodeOrDescChange();
  }

  function toggleComboboxDropdown(type, event) {
    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }
    const dropdown = document.getElementById(`dropdown-${type}`);
    if (!dropdown) return;

    if (state.activeDropdown === type) {
      closeAllDropdowns();
    } else {
      closeAllDropdowns();
      const input = document.getElementById(type === 'code' ? 'add-code' : 'add-desc');
      renderDropdownItems(type, input ? input.value.trim() : '');
      openDropdown(type);
      if (input) input.focus();
    }
  }

  function openDropdown(type) {
    closeAllDropdowns();
    const dropdown = document.getElementById(`dropdown-${type}`);
    if (dropdown) {
      dropdown.style.display = 'block';
      state.activeDropdown = type;
      state.focusedIndex = -1;
    }
  }

  function closeAllDropdowns() {
    const dropCode = document.getElementById('dropdown-code');
    const dropDesc = document.getElementById('dropdown-desc');
    if (dropCode) dropCode.style.display = 'none';
    if (dropDesc) dropDesc.style.display = 'none';
    state.activeDropdown = null;
    state.focusedIndex = -1;
  }

  function updateClearBtn(inputId, value) {
    const btn = document.getElementById(inputId === 'add-code' ? 'btn-clear-code' : 'btn-clear-desc');
    if (btn) {
      btn.style.display = value ? 'flex' : 'none';
    }
  }

  function renderDropdownItems(type, query) {
    const listEl = document.getElementById(`list-${type}`);
    const countEl = document.getElementById(`count-${type}-matches`);
    if (!listEl) return;

    const catalog = state.masterCatalog || [];
    const qLower = (query || '').toLowerCase().trim();

    let matches = [];
    if (!qLower) {
      matches = catalog;
    } else {
      matches = catalog.filter(item => {
        const c = item.code.toLowerCase();
        const d = item.desc.toLowerCase();
        return c.includes(qLower) || d.includes(qLower);
      }).sort((a, b) => {
        const aCodeMatch = a.code.toLowerCase().startsWith(qLower);
        const bCodeMatch = b.code.toLowerCase().startsWith(qLower);
        const aDescMatch = a.desc.toLowerCase().startsWith(qLower);
        const bDescMatch = b.desc.toLowerCase().startsWith(qLower);

        if (type === 'code') {
          if (aCodeMatch && !bCodeMatch) return -1;
          if (!aCodeMatch && bCodeMatch) return 1;
        } else {
          if (aDescMatch && !bDescMatch) return -1;
          if (!aDescMatch && bDescMatch) return 1;
        }
        return a.desc.localeCompare(b.desc);
      });
    }

    if (countEl) {
      countEl.textContent = `${matches.length} item${matches.length !== 1 ? 's' : ''}`;
    }

    if (matches.length === 0) {
      listEl.innerHTML = `
        <div class="dropdown-empty">
          <div>🔍 No exact matches in Sheet1 Master Catalog.</div>
          <div class="dropdown-empty-hint">You can still use this as a brand new medicine entry.</div>
        </div>
      `;
      return;
    }

    // Render top 60 matches for optimal scrolling performance
    const displayList = matches.slice(0, 60);
    listEl.innerHTML = displayList.map((item, idx) => {
      const codeHtml = highlightText(item.code, qLower);
      const descHtml = highlightText(item.desc, qLower);

      if (type === 'code') {
        return `
          <div class="dropdown-item" data-index="${idx}" onclick="selectDciItem('${escapeJs(item.code)}', '${escapeJs(item.desc)}')">
            <span class="dropdown-item-code">${codeHtml}</span>
            <span class="dropdown-item-desc">${descHtml}</span>
          </div>
        `;
      } else {
        return `
          <div class="dropdown-item" data-index="${idx}" onclick="selectDciItem('${escapeJs(item.code)}', '${escapeJs(item.desc)}')">
            <span class="dropdown-item-desc">${descHtml}</span>
            <span class="dropdown-item-code">${codeHtml}</span>
          </div>
        `;
      }
    }).join('');
  }

  function handleComboboxKeydown(e, type) {
    const dropdown = document.getElementById(`dropdown-${type}`);
    if (!dropdown || dropdown.style.display === 'none') {
      if (e.key === 'ArrowDown' || e.key === 'Enter') {
        openDropdown(type);
        renderDropdownItems(type, e.target.value.trim());
      }
      return;
    }

    const items = dropdown.querySelectorAll('.dropdown-item');
    if (items.length === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      state.focusedIndex = (state.focusedIndex + 1) % items.length;
      updateDropdownFocus(items);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      state.focusedIndex = (state.focusedIndex - 1 + items.length) % items.length;
      updateDropdownFocus(items);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (state.focusedIndex >= 0 && state.focusedIndex < items.length) {
        items[state.focusedIndex].click();
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      closeAllDropdowns();
    }
  }

  function updateDropdownFocus(items) {
    items.forEach((item, i) => {
      if (i === state.focusedIndex) {
        item.classList.add('focused');
        item.scrollIntoView({ block: 'nearest' });
      } else {
        item.classList.remove('focused');
      }
    });
  }

  function selectDciItem(code, desc) {
    const codeInput = document.getElementById('add-code');
    const descInput = document.getElementById('add-desc');
    const brandInput = document.getElementById('add-brand');

    if (codeInput) {
      codeInput.value = code;
      updateClearBtn('add-code', code);
    }
    if (descInput) {
      descInput.value = desc;
      updateClearBtn('add-desc', desc);
    }

    closeAllDropdowns();
    onAddCodeOrDescChange();

    // Smooth workflow: focus brand name for incoming delivery
    if (brandInput && !brandInput.value) {
      setTimeout(() => brandInput.focus(), 60);
    }
  }

  function clearAddInput(inputId) {
    const input = document.getElementById(inputId);
    if (input) {
      input.value = '';
      updateClearBtn(inputId, '');
      input.focus();
      closeAllDropdowns();
      onAddCodeOrDescChange();
    }
  }

  function resetDciSelection() {
    const codeInput = document.getElementById('add-code');
    const descInput = document.getElementById('add-desc');
    const previewBox = document.getElementById('add-batch-preview');

    if (codeInput) {
      codeInput.value = '';
      updateClearBtn('add-code', '');
    }
    if (descInput) {
      descInput.value = '';
      updateClearBtn('add-desc', '');
    }
    if (previewBox) {
      previewBox.style.display = 'none';
    }

    closeAllDropdowns();
    if (codeInput) codeInput.focus();
  }

  function highlightText(text, query) {
    if (!text) return '';
    if (!query) return escapeHtml(text);
    const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`(${escapedQuery})`, 'gi');
    return escapeHtml(text).replace(regex, '<mark class="dropdown-highlight">$1</mark>');
  }

  function escapeJs(str) {
    if (!str) return '';
    return String(str).replace(/'/g, "\\'").replace(/"/g, '&quot;');
  }

  function populateCatalogDatalists() {
    const codeList = document.getElementById('dci-codes-datalist');
    const descList = document.getElementById('generic-names-datalist');
    const badge = document.getElementById('catalog-count-badge');

    if (!state.masterCatalog || state.masterCatalog.length === 0) return;

    if (codeList) {
      codeList.innerHTML = state.masterCatalog.map(item => 
        `<option value="${escapeHtml(item.code)}">${escapeHtml(item.desc)}</option>`
      ).join('');
    }

    if (descList) {
      descList.innerHTML = state.masterCatalog.map(item => 
        `<option value="${escapeHtml(item.desc)}">${escapeHtml(item.code)}</option>`
      ).join('');
    }

    if (badge) {
      badge.textContent = `📋 ${state.masterCatalog.length} in Sheet1`;
      badge.style.display = 'inline-block';
    }
  }

  function escapeHtml(text) {
    if (!text) return '';
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
</script>
'@

if ($jsContent.Contains($oldTail)) {
    $jsContent = $jsContent.Replace($oldTail, $newTail)
    Write-Output "Tail functions updated."
} else {
    Write-Warning "oldTail block not found directly, checking normalized line breaks..."
    $normalizedOld = $oldTail.Replace("`r`n", "`n")
    $normalizedContent = $jsContent.Replace("`r`n", "`n")
    if ($normalizedContent.Contains($normalizedOld)) {
        $normalizedContent = $normalizedContent.Replace($normalizedOld, $newTail.Replace("`r`n", "`n"))
        $jsContent = $normalizedContent.Replace("`n", "`r`n")
        Write-Output "Tail functions updated with normalized line breaks."
    } else {
        Write-Error "Failed to match tail functions."
    }
}

Set-Content -Path 'gas_project\JavaScript.html' -Value $jsContent -Encoding UTF8
Write-Output "gas_project\JavaScript.html saved successfully."
