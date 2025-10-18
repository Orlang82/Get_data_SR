// Wiresock Auto Launcher - Background Script

// Имя native messaging host (должно совпадать с именем в манифесте хоста)
const NATIVE_HOST_NAME = 'com.wiresock.launcher';

// Хранилище для списка доменов и активных соединений
let monitoredDomains = [];
let activeConnections = new Set();

// Загрузка списка доменов из storage при инициализации
chrome.runtime.onInstalled.addListener(async () => {
  console.log('Wiresock Auto Launcher установлен');

  // Загружаем сохраненный список доменов или используем пример
  const result = await chrome.storage.sync.get(['domains']);
  monitoredDomains = result.domains || ['mail.ru', 'example.com'];

  // Сохраняем дефолтный список, если его не было
  if (!result.domains) {
    await chrome.storage.sync.set({ domains: monitoredDomains });
  }

  console.log('Загружены домены:', monitoredDomains);
});

// Загрузка доменов при старте расширения
chrome.runtime.onStartup.addListener(async () => {
  const result = await chrome.storage.sync.get(['domains']);
  monitoredDomains = result.domains || [];
  console.log('Расширение запущено, домены:', monitoredDomains);
});

// Функция для проверки, соответствует ли URL одному из отслеживаемых доменов
function isDomainMonitored(url) {
  try {
    const urlObj = new URL(url);
    const hostname = urlObj.hostname;

    return monitoredDomains.some(domain => {
      // Убираем возможные пробелы
      const cleanDomain = domain.trim().toLowerCase();
      const cleanHostname = hostname.toLowerCase();

      // Проверка на wildcard-паттерны
      if (cleanDomain.includes('*')) {
        // Преобразуем паттерн с wildcard в регулярное выражение
        // Экранируем специальные символы regex, кроме *
        const regexPattern = cleanDomain
          .replace(/[.+?^${}()|[\]\\]/g, '\\$&')  // Экранируем спецсимволы
          .replace(/\*/g, '.*');                    // Заменяем * на .*

        const regex = new RegExp('^' + regexPattern + '$', 'i');
        return regex.test(cleanHostname);
      }

      // Точное совпадение или поддомен (для обычных доменов без *)
      return cleanHostname === cleanDomain || cleanHostname.endsWith('.' + cleanDomain);
    });
  } catch (error) {
    console.error('Ошибка парсинга URL:', error);
    return false;
  }
}

// Функция для запуска wiresock-client через native messaging
async function launchWiresock(domain) {
  try {
    console.log(`Попытка запуска wiresock для домена: ${domain}`);

    // Получаем настройки из storage
    const settings = await chrome.storage.sync.get(['configPath']);
    const configPath = settings.configPath || 'C:\\path\\to\\config.conf';

    // Отправляем сообщение native host приложению
    const response = await chrome.runtime.sendNativeMessage(
      NATIVE_HOST_NAME,
      {
        action: 'launch',
        domain: domain,
        configPath: configPath
      }
    );

    if (response && response.success) {
      console.log('Wiresock успешно запущен:', response.message);
      activeConnections.add(domain);

      // Отправляем уведомление пользователю
      chrome.notifications.create({
        type: 'basic',
        iconUrl: 'icons/icon48.png',
        title: 'Wiresock запущен',
        message: `Подключение активировано для ${domain}`
      });
    } else {
      console.error('Ошибка запуска wiresock:', response?.error || 'Неизвестная ошибка');
    }
  } catch (error) {
    console.error('Ошибка при запуске wiresock:', error);

    // Показываем ошибку пользователю
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icons/icon48.png',
      title: 'Ошибка Wiresock',
      message: `Не удалось запустить wiresock: ${error.message}`
    });
  }
}

// Слушаем навигацию по вкладкам
chrome.webNavigation.onBeforeNavigate.addListener(async (details) => {
  // Проверяем только главные фреймы (не iframe)
  if (details.frameId === 0) {
    const url = details.url;

    if (isDomainMonitored(url)) {
      const urlObj = new URL(url);
      const domain = urlObj.hostname;

      console.log(`Обнаружен переход на отслеживаемый домен: ${domain}`);

      // Проверяем, не запущено ли уже соединение для этого домена
      if (!activeConnections.has(domain)) {
        await launchWiresock(domain);
      } else {
        console.log(`Соединение для ${domain} уже активно`);
      }
    }
  }
});

// Обработка сообщений из popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'getDomains') {
    sendResponse({ domains: monitoredDomains });
  } else if (request.action === 'updateDomains') {
    monitoredDomains = request.domains;
    chrome.storage.sync.set({ domains: monitoredDomains });
    console.log('Обновлен список доменов:', monitoredDomains);
    sendResponse({ success: true });
  } else if (request.action === 'getStatus') {
    sendResponse({
      activeConnections: Array.from(activeConnections),
      monitoredDomains: monitoredDomains
    });
  } else if (request.action === 'testConnection') {
    // Тестовый запуск wiresock
    launchWiresock('test-domain.com')
      .then(() => sendResponse({ success: true }))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true; // Асинхронный ответ
  }

  return false;
});

// Обновление списка доменов при изменении storage
chrome.storage.onChanged.addListener((changes, namespace) => {
  if (namespace === 'sync' && changes.domains) {
    monitoredDomains = changes.domains.newValue || [];
    console.log('Список доменов обновлен из storage:', monitoredDomains);
  }
});

console.log('Background script загружен');
