// Wiresock Auto Launcher - Popup Script

let currentDomains = [];

// Инициализация при загрузке popup
document.addEventListener('DOMContentLoaded', async () => {
  await loadDomains();
  await loadConfig();
  await loadStatus();

  // Назначаем обработчики событий
  document.getElementById('add-domain').addEventListener('click', addDomain);
  document.getElementById('save-config').addEventListener('click', saveConfig);
  document.getElementById('test-connection').addEventListener('click', testConnection);

  // Добавление домена по Enter
  document.getElementById('new-domain').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      addDomain();
    }
  });
});

// Загрузка списка доменов
async function loadDomains() {
  try {
    const response = await chrome.runtime.sendMessage({ action: 'getDomains' });
    currentDomains = response.domains || [];
    renderDomains();
  } catch (error) {
    console.error('Ошибка загрузки доменов:', error);
    showMessage('Ошибка загрузки доменов', 'error');
  }
}

// Отображение списка доменов
function renderDomains() {
  const domainsList = document.getElementById('domains-list');

  if (currentDomains.length === 0) {
    domainsList.innerHTML = '<div class="empty-list">Нет отслеживаемых доменов</div>';
    return;
  }

  domainsList.innerHTML = currentDomains.map((domain, index) => `
    <div class="domain-item">
      <span class="domain-name">${escapeHtml(domain)}</span>
      <button class="remove-btn" data-index="${index}">Удалить</button>
    </div>
  `).join('');

  // Добавляем обработчики для кнопок удаления
  domainsList.querySelectorAll('.remove-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const index = parseInt(e.target.getAttribute('data-index'));
      removeDomain(index);
    });
  });
}

// Добавление нового домена
async function addDomain() {
  const input = document.getElementById('new-domain');
  const domain = input.value.trim();

  if (!domain) {
    showMessage('Введите домен', 'error');
    return;
  }

  // Проверка формата домена (базовая)
  if (!/^[a-zA-Z0-9][a-zA-Z0-9-]{0,61}[a-zA-Z0-9]?(\.[a-zA-Z]{2,})+$/.test(domain)) {
    showMessage('Неверный формат домена', 'error');
    return;
  }

  // Проверка на дубликаты
  if (currentDomains.includes(domain)) {
    showMessage('Домен уже добавлен', 'error');
    return;
  }

  currentDomains.push(domain);
  await saveDomains();
  input.value = '';
  renderDomains();
  showMessage('Домен добавлен', 'success');
}

// Удаление домена
async function removeDomain(index) {
  currentDomains.splice(index, 1);
  await saveDomains();
  renderDomains();
  showMessage('Домен удален', 'success');
}

// Сохранение списка доменов
async function saveDomains() {
  try {
    await chrome.runtime.sendMessage({
      action: 'updateDomains',
      domains: currentDomains
    });
  } catch (error) {
    console.error('Ошибка сохранения доменов:', error);
    showMessage('Ошибка сохранения доменов', 'error');
  }
}

// Загрузка конфигурации
async function loadConfig() {
  try {
    const result = await chrome.storage.sync.get(['configPath']);
    const configPath = result.configPath || '';
    document.getElementById('config-path').value = configPath;
  } catch (error) {
    console.error('Ошибка загрузки конфигурации:', error);
  }
}

// Сохранение пути к конфигурации
async function saveConfig() {
  const configPath = document.getElementById('config-path').value.trim();

  if (!configPath) {
    showMessage('Введите путь к файлу конфигурации', 'error');
    return;
  }

  try {
    await chrome.storage.sync.set({ configPath: configPath });
    showMessage('Путь к конфигурации сохранен', 'success');
  } catch (error) {
    console.error('Ошибка сохранения конфигурации:', error);
    showMessage('Ошибка сохранения', 'error');
  }
}

// Загрузка статуса активных подключений
async function loadStatus() {
  try {
    const response = await chrome.runtime.sendMessage({ action: 'getStatus' });
    const activeConnections = response.activeConnections || [];

    document.getElementById('active-count').textContent = activeConnections.length;

    const activeDomainsDiv = document.getElementById('active-domains');
    if (activeConnections.length > 0) {
      activeDomainsDiv.innerHTML = activeConnections.map(domain =>
        `<div class="active-domain">${escapeHtml(domain)}</div>`
      ).join('');
    } else {
      activeDomainsDiv.innerHTML = '<div style="font-size: 12px; color: #999;">Нет активных подключений</div>';
    }
  } catch (error) {
    console.error('Ошибка загрузки статуса:', error);
  }
}

// Тестирование подключения
async function testConnection() {
  const btn = document.getElementById('test-connection');
  const originalText = btn.textContent;

  btn.textContent = 'Тестирование...';
  btn.disabled = true;

  try {
    const response = await chrome.runtime.sendMessage({ action: 'testConnection' });

    if (response.success) {
      showMessage('Тестовое подключение успешно', 'success');
    } else {
      showMessage(`Ошибка: ${response.error}`, 'error');
    }
  } catch (error) {
    showMessage(`Ошибка тестирования: ${error.message}`, 'error');
  } finally {
    btn.textContent = originalText;
    btn.disabled = false;
  }
}

// Показ сообщений
function showMessage(text, type) {
  const messageDiv = document.getElementById('message');
  messageDiv.textContent = text;
  messageDiv.className = `message ${type}`;

  setTimeout(() => {
    messageDiv.className = 'message';
  }, 3000);
}

// Экранирование HTML для безопасности
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}
