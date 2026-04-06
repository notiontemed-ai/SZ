/**
 * Добавляет пользовательское меню.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TEMED')
    .addItem('Подготовить реестр заданий', 'prepareTaskRegistry')
    .addToUi();
}

/**
 * Основная функция подготовки реестра заданий.
 */
function prepareTaskRegistry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const prepSheet = getSheetByNameOrThrow(ss, 'Подготовка реестра заданий');
  const executorsSheet = getSheetByNameOrThrow(ss, 'Список исполнителей');
  const servicesSheet = getSheetByNameOrThrow(ss, 'Реестр услуг');
  const registrySheet = getSheetByNameOrThrow(ss, 'Реестр');

  const prepData = getSheetDataWithHeaders(prepSheet, ['Месяц', 'Исполнитель', 'Сумма']);
  const executorsData = getSheetDataWithHeaders(executorsSheet, ['Исполнитель', 'ИНН']);
  const servicesData = getSheetDataWithHeaders(servicesSheet, [
    'Исполнитель',
    'Наименование, описание и результат оказания услуги',
    'Мин стоимость',
    'Макс стоимость',
    'Ед. изм.',
    'Категория услуги',
  ]);
  const registryHeaders = getHeaderMapOrThrow(registrySheet, [
    'ИНН',
    'Период ОТ',
    'Период ДО',
    'Описание услуги',
    'Категория услуги',
    'Кол-во',
    'Цена',
    'Стоимость',
  ]);

  clearRegistryData(registrySheet);

  const innMap = buildInnMap(executorsData.rows, executorsData.headerMap);
  const services = buildServices(servicesData.rows, servicesData.headerMap);

  const outputRows = [];
  for (let i = 0; i < prepData.rows.length; i++) {
    const sourceRowNumber = i + 2;
    const row = prepData.rows[i];

    const monthCodeRaw = toTrimmedString(row[prepData.headerMap['Месяц']]);
    const executor = toTrimmedString(row[prepData.headerMap['Исполнитель']]);
    const amountRaw = row[prepData.headerMap['Сумма']];

    if (!monthCodeRaw && !executor && (amountRaw === '' || amountRaw === null)) {
      continue;
    }

    if (!monthCodeRaw || !executor || amountRaw === '' || amountRaw === null) {
      throw new Error('Неполные данные в строке ' + sourceRowNumber + ' листа "Подготовка реестра заданий".');
    }

    const amount = toIntegerOrThrow(amountRaw, 'Некорректная сумма в строке ' + sourceRowNumber + ': ' + amountRaw);
    if (amount <= 0) {
      throw new Error('Сумма должна быть больше 0 в строке ' + sourceRowNumber + '.');
    }

    const period = parseMonthCode(monthCodeRaw);
    const inn = findInnByExecutor(innMap, executor);
    const servicesForExecutor = filterServicesForExecutor(services, executor);

    if (servicesForExecutor.length === 0) {
      throw new Error('Для исполнителя ' + executor + ' не найдено ни одной подходящей услуги.');
    }

    const taskSet = buildTaskSetForAmount(servicesForExecutor, amount);
    if (!taskSet) {
      throw new Error('Не удалось подобрать услуги на сумму ' + amount + ' для исполнителя ' + executor + '.');
    }

    for (let j = 0; j < taskSet.length; j++) {
      const item = taskSet[j];
      outputRows.push(buildRegistryRow(registrySheet.getLastColumn(), registryHeaders, {
        inn: inn,
        periodFrom: period.from,
        periodTo: period.to,
        description: item.description,
        category: item.category,
        qty: item.qty,
        price: item.price,
        total: item.total,
      }));
    }
  }

  if (outputRows.length > 0) {
    registrySheet.getRange(2, 1, outputRows.length, registrySheet.getLastColumn()).setValues(outputRows);
  }
}

function getSheetByNameOrThrow(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Не найден обязательный лист: "' + name + '".');
  }
  return sheet;
}

function getSheetDataWithHeaders(sheet, requiredHeaders) {
  const headerMap = getHeaderMapOrThrow(sheet, requiredHeaders);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { headerMap: headerMap, rows: rows, lastCol: lastCol };
}

function getHeaderMapOrThrow(sheet, requiredHeaders) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    throw new Error('Лист "' + sheet.getName() + '" не содержит заголовков.');
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const name = toTrimmedString(headers[i]);
    if (name) {
      map[name] = i;
    }
  }

  for (let i = 0; i < requiredHeaders.length; i++) {
    const header = requiredHeaders[i];
    if (map[header] === undefined) {
      throw new Error('На листе "' + sheet.getName() + '" отсутствует колонка "' + header + '".');
    }
  }

  return map;
}

function clearRegistryData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}

function parseMonthCode(monthCode) {
  const value = toTrimmedString(monthCode);
  if (!/^\d{4}$/.test(value)) {
    throw new Error('Некорректный код месяца: ' + monthCode);
  }

  const yy = Number(value.slice(0, 2));
  const mm = Number(value.slice(2, 4));
  if (mm < 1 || mm > 12) {
    throw new Error('Некорректный код месяца: ' + monthCode);
  }

  const year = 2000 + yy;
  const firstDate = new Date(year, mm - 1, 1);
  const lastDate = new Date(year, mm, 0);

  return {
    from: formatDateDDMMYY(firstDate),
    to: formatDateDDMMYY(lastDate),
  };
}

function formatDateDDMMYY(date) {
  const dd = pad2(date.getDate());
  const mm = pad2(date.getMonth() + 1);
  const yy = pad2(date.getFullYear() % 100);
  return dd + '/' + mm + '/' + yy;
}

function pad2(n) {
  return n < 10 ? '0' + n : String(n);
}

function buildInnMap(rows, headerMap) {
  const map = {};
  for (let i = 0; i < rows.length; i++) {
    const executor = toTrimmedString(rows[i][headerMap['Исполнитель']]);
    const inn = toTrimmedString(rows[i][headerMap['ИНН']]);
    if (executor && inn) {
      map[executor] = inn;
    }
  }
  return map;
}

function findInnByExecutor(innMap, executor) {
  if (!innMap[executor]) {
    throw new Error('Не найден ИНН для исполнителя: ' + executor);
  }
  return innMap[executor];
}

function buildServices(rows, headerMap) {
  const result = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const executorsRaw = toTrimmedString(row[headerMap['Исполнитель']]);
    const description = toTrimmedString(row[headerMap['Наименование, описание и результат оказания услуги']]);
    const minCostRaw = row[headerMap['Мин стоимость']];
    const maxCostRaw = row[headerMap['Макс стоимость']];
    const unit = toTrimmedString(row[headerMap['Ед. изм.']]);
    const category = toTrimmedString(row[headerMap['Категория услуги']]);

    if (!executorsRaw || !description || minCostRaw === '' || maxCostRaw === '') {
      continue;
    }

    const minCost = toIntegerOrThrow(minCostRaw, 'Некорректное значение "Мин стоимость" в листе "Реестр услуг", строка ' + (i + 2) + '.');
    const maxCost = toIntegerOrThrow(maxCostRaw, 'Некорректное значение "Макс стоимость" в листе "Реестр услуг", строка ' + (i + 2) + '.');

    if (minCost <= 0 || maxCost <= 0 || minCost > maxCost) {
      throw new Error('Некорректный диапазон стоимости в листе "Реестр услуг", строка ' + (i + 2) + '.');
    }

    const allowedExecutors = executorsRaw
      .split(',')
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s.length > 0; });

    result.push({
      id: i,
      allowedExecutors: allowedExecutors,
      description: description,
      minCost: minCost,
      maxCost: maxCost,
      unit: unit,
      category: category,
    });
  }
  return result;
}

function filterServicesForExecutor(services, executor) {
  return services.filter(function (s) {
    return s.allowedExecutors.indexOf(executor) !== -1;
  });
}

function buildTaskSetForAmount(servicesForExecutor, targetAmount) {
  const minServices = 3;
  const maxServices = 8;

  if (servicesForExecutor.length < minServices) {
    return null;
  }

  const shuffledServices = shuffleArray(servicesForExecutor.slice());
  const suffixBounds = buildSuffixBounds(shuffledServices);

  function dfs(index, chosenCount, currentSum, chosenItems) {
    if (currentSum > targetAmount) {
      return null;
    }
    if (chosenCount > maxServices) {
      return null;
    }

    const remainingServices = shuffledServices.length - index;
    if (chosenCount + remainingServices < minServices) {
      return null;
    }

    if (chosenCount >= minServices && currentSum === targetAmount) {
      return chosenItems.slice();
    }

    if (index >= shuffledServices.length) {
      return null;
    }

    const needMin = Math.max(0, minServices - chosenCount);
    const canTakeMax = Math.min(maxServices - chosenCount, remainingServices);

    const suffix = suffixBounds[index];
    if (needMin > 0) {
      const minPossible = currentSum + suffix.minForK[needMin];
      if (!isFinite(minPossible) || minPossible > targetAmount) {
        return null;
      }
    }

    const maxPossible = currentSum + suffix.maxForK[canTakeMax];
    if (!isFinite(maxPossible) || maxPossible < targetAmount) {
      return null;
    }

    const service = shuffledServices[index];
    const variants = buildServiceVariants(service);
    shuffleArray(variants);

    for (let v = 0; v < variants.length; v++) {
      const variant = variants[v];
      const nextSum = currentSum + variant.total;
      if (nextSum > targetAmount) {
        continue;
      }

      const remainingAfterTake = shuffledServices.length - (index + 1);
      if (chosenCount + 1 + remainingAfterTake < minServices) {
        continue;
      }

      const suffixAfter = suffixBounds[index + 1];
      const needAfter = Math.max(0, minServices - (chosenCount + 1));
      const canAfter = Math.min(maxServices - (chosenCount + 1), remainingAfterTake);

      if (needAfter <= canAfter) {
        const minAfter = nextSum + suffixAfter.minForK[needAfter];
        const maxAfter = nextSum + suffixAfter.maxForK[canAfter];
        if (minAfter <= targetAmount && maxAfter >= targetAmount) {
          chosenItems.push(variant);
          const found = dfs(index + 1, chosenCount + 1, nextSum, chosenItems);
          if (found) {
            return found;
          }
          chosenItems.pop();
        }
      }
    }

    const skipFound = dfs(index + 1, chosenCount, currentSum, chosenItems);
    if (skipFound) {
      return skipFound;
    }

    return null;
  }

  return dfs(0, 0, 0, []);
}

function buildServiceVariants(service) {
  const variants = [];
  const isPiece = service.unit === 'Штука';
  const minQty = isPiece ? 1 : 1;
  const maxQty = isPiece ? 5 : 1;

  for (let qty = minQty; qty <= maxQty; qty++) {
    for (let price = service.minCost; price <= service.maxCost; price++) {
      variants.push({
        serviceId: service.id,
        description: service.description,
        category: service.category,
        qty: qty,
        price: price,
        total: qty * price,
      });
    }
  }

  return variants;
}

function buildSuffixBounds(services) {
  const n = services.length;
  const suffix = new Array(n + 1);
  suffix[n] = {
    minForK: [0],
    maxForK: [0],
  };

  for (let i = n - 1; i >= 0; i--) {
    const service = services[i];
    const minOne = service.minCost;
    const maxOne = service.unit === 'Штука' ? service.maxCost * 5 : service.maxCost;
    const next = suffix[i + 1];

    const maxK = next.minForK.length;
    const minForK = new Array(maxK + 1);
    const maxForK = new Array(maxK + 1);

    minForK[0] = 0;
    maxForK[0] = 0;

    for (let k = 1; k <= maxK; k++) {
      const skipMin = next.minForK[k] !== undefined ? next.minForK[k] : Infinity;
      const takeMin = next.minForK[k - 1] !== undefined ? minOne + next.minForK[k - 1] : Infinity;
      minForK[k] = Math.min(skipMin, takeMin);

      const skipMax = next.maxForK[k] !== undefined ? next.maxForK[k] : -Infinity;
      const takeMax = next.maxForK[k - 1] !== undefined ? maxOne + next.maxForK[k - 1] : -Infinity;
      maxForK[k] = Math.max(skipMax, takeMax);
    }

    suffix[i] = {
      minForK: minForK,
      maxForK: maxForK,
    };
  }

  return suffix;
}

function buildRegistryRow(lastCol, registryHeaders, payload) {
  const row = new Array(lastCol).fill('');
  row[registryHeaders['ИНН']] = payload.inn;
  row[registryHeaders['Период ОТ']] = payload.periodFrom;
  row[registryHeaders['Период ДО']] = payload.periodTo;
  row[registryHeaders['Описание услуги']] = payload.description;
  row[registryHeaders['Категория услуги']] = payload.category;
  row[registryHeaders['Кол-во']] = payload.qty;
  row[registryHeaders['Цена']] = payload.price;
  row[registryHeaders['Стоимость']] = payload.total;
  return row;
}

function toIntegerOrThrow(value, message) {
  const num = Number(value);
  if (!Number.isFinite(num) || Math.floor(num) !== num) {
    throw new Error(message);
  }
  return num;
}

function toTrimmedString(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = arr[i];
    arr[i] = arr[j];
    arr[j] = tmp;
  }
  return arr;
}
