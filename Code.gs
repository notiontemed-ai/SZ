const HISTORY_SPREADSHEET_ID = '1A6eXDmV5VCTCctYNoilVDPIdc257ZipqGwetiRIWRaI';

/**
 * Добавляет пользовательское меню.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TEMED')
    .addItem('Подготовить реестр заданий', 'prepareTaskRegistry')
    .addItem('Утвердить задание', 'approveTaskRegistry')
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
  const enterprisesSheet = getSheetByNameOrThrow(ss, 'Предприятия');
  const historySheet = getHistorySheetOrThrow();
  const registrySheet = getSheetByNameOrThrow(ss, 'Реестр');

  const prepData = getSheetDataWithHeaders(prepSheet, ['Месяц', 'Исполнитель', 'Сумма']);
  const executorsData = getSheetDataWithHeaders(executorsSheet, [
    'Исполнитель',
    'ИНН',
    'Название',
    'Запрещенные Заказчики',
  ]);
  const servicesData = getSheetDataWithHeaders(servicesSheet, [
    'Исполнитель',
    'Наименование, описание и результат оказания услуги',
    'Мин стоимость',
    'Макс стоимость',
    'Ед. изм.',
    'Категория услуги',
  ]);
  const enterprisesData = getSheetDataWithHeaders(enterprisesSheet, ['Заказчик']);
  const historyData = getSheetDataWithHeaders(historySheet, ['Месяц', 'ИНН', 'Заказчик']);
  const registryHeaders = getHeaderMapOrThrow(registrySheet, [
    'Месяц',
    'ИНН',
    'ФИО',
    'Название',
    'Период ОТ',
    'Период ДО',
    'Заказчик',
    'Описание услуги',
    'Категория услуги',
    'Единица',
    'Кол-во',
    'Цена',
    'Стоимость',
  ]);

  clearRegistryData(registrySheet);

  const executorMap = buildExecutorMap(executorsData.rows, executorsData.headerMap);
  const services = buildServices(servicesData.rows, servicesData.headerMap);
  const customerPool = buildCustomerPool(enterprisesData.rows, enterprisesData.headerMap);

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
    const executorInfo = findExecutorInfoOrThrow(executorMap, executor);
    const inn = executorInfo.inn;
    const servicesForExecutor = filterServicesForExecutor(services, executor);

    if (servicesForExecutor.length === 0) {
      throw new Error('Для исполнителя ' + executor + ' не найдено ни одной подходящей услуги.');
    }

    const taskSet = buildTaskSetForAmount(servicesForExecutor, amount);
    if (!taskSet) {
      throw new Error('Не удалось подобрать услуги на сумму ' + amount + ' для исполнителя ' + executor + '.');
    }
    const bannedCustomers = buildRecentCustomerSetByInn(
      historyData.rows,
      historyData.headerMap,
      inn,
      monthCodeRaw
    );
    mergeBannedCustomers(bannedCustomers, executorInfo.bannedCustomers);
    const pickedCustomers = pickTwoCustomers(customerPool, bannedCustomers, executor, monthCodeRaw);
    const assignedCustomers = assignCustomersToTaskItems(taskSet, pickedCustomers[0], pickedCustomers[1]);

    for (let j = 0; j < taskSet.length; j++) {
      const item = taskSet[j];
      outputRows.push(buildRegistryRow(registrySheet.getLastColumn(), registryHeaders, {
        month: monthCodeRaw,
        inn: inn,
        fullName: executor,
        title: executorInfo.title,
        periodFrom: period.from,
        periodTo: period.to,
        customer: assignedCustomers[j],
        description: item.description,
        category: item.category,
        unit: item.unit,
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

function getHistorySheetOrThrow() {
  const historySpreadsheet = SpreadsheetApp.openById(HISTORY_SPREADSHEET_ID);
  return getSheetByNameOrThrow(historySpreadsheet, 'История');
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

function buildExecutorMap(rows, headerMap) {
  const map = {};
  for (let i = 0; i < rows.length; i++) {
    const executor = toTrimmedString(rows[i][headerMap['Исполнитель']]);
    const inn = toTrimmedString(rows[i][headerMap['ИНН']]);
    const title = toTrimmedString(rows[i][headerMap['Название']]);
    const bannedCustomers = parseCommaSeparatedValues(rows[i][headerMap['Запрещенные Заказчики']]);
    if (executor && inn && title) {
      map[executor] = {
        inn: inn,
        title: title,
        bannedCustomers: bannedCustomers,
      };
    }
  }
  return map;
}

function buildCustomerPool(rows, headerMap) {
  const uniqueMap = {};
  const customers = [];

  for (let i = 0; i < rows.length; i++) {
    const customer = toTrimmedString(rows[i][headerMap['Заказчик']]);
    if (!customer || uniqueMap[customer]) {
      continue;
    }
    uniqueMap[customer] = true;
    customers.push(customer);
  }

  if (customers.length < 2) {
    throw new Error('На листе "Предприятия" должно быть минимум 2 уникальных непустых заказчика в колонке "Заказчик".');
  }

  return customers;
}

function getPreviousMonthCodes(monthCode, count) {
  const value = toTrimmedString(monthCode);
  if (!/^\d{4}$/.test(value)) {
    throw new Error('Некорректный код месяца: ' + monthCode);
  }

  let yy = Number(value.slice(0, 2));
  let mm = Number(value.slice(2, 4));
  if (mm < 1 || mm > 12) {
    throw new Error('Некорректный код месяца: ' + monthCode);
  }

  const result = [];
  for (let i = 0; i < count; i++) {
    mm -= 1;
    if (mm < 1) {
      mm = 12;
      yy -= 1;
      if (yy < 0) {
        yy = 99;
      }
    }
    result.push(pad2(yy) + pad2(mm));
  }
  return result;
}

function buildRecentCustomerSetByInn(historyRows, historyHeaderMap, inn, monthCode) {
  const previousMonths = getPreviousMonthCodes(monthCode, 2);
  const monthMap = {};
  for (let i = 0; i < previousMonths.length; i++) {
    monthMap[previousMonths[i]] = true;
  }

  const recentCustomers = {};
  for (let i = 0; i < historyRows.length; i++) {
    const row = historyRows[i];
    const rowInn = toTrimmedString(row[historyHeaderMap['ИНН']]);
    const rowMonth = toTrimmedString(row[historyHeaderMap['Месяц']]);
    const rowCustomer = toTrimmedString(row[historyHeaderMap['Заказчик']]);

    if (!rowInn || !rowMonth || !rowCustomer) {
      continue;
    }
    if (rowInn !== inn) {
      continue;
    }
    if (!monthMap[rowMonth]) {
      continue;
    }

    recentCustomers[normalizeCustomerNameKey(rowCustomer)] = true;
  }
  return recentCustomers;
}

function pickTwoCustomers(availableCustomers, bannedCustomers, executor, monthCode) {
  const allowed = [];

  for (let i = 0; i < availableCustomers.length; i++) {
    const customer = availableCustomers[i];
    if (!bannedCustomers[normalizeCustomerNameKey(customer)]) {
      allowed.push(customer);
    }
  }

  if (allowed.length < 2) {
    throw new Error(
      'Для исполнителя "' + executor + '" в месяце ' + monthCode +
      ' после исключения заказчиков за два предыдущих месяца осталось меньше двух доступных заказчиков.'
    );
  }

  shuffleArray(allowed);
  return [allowed[0], allowed[1]];
}

function mergeBannedCustomers(target, extraCustomers) {
  for (let i = 0; i < extraCustomers.length; i++) {
    target[normalizeCustomerNameKey(extraCustomers[i])] = true;
  }
}

function assignCustomersToTaskItems(taskSet, customerA, customerB) {
  const count = taskSet.length;
  if (count === 0) {
    return [];
  }
  if (count === 1) {
    return [Math.random() < 0.5 ? customerA : customerB];
  }

  const assigned = [customerA, customerB];
  for (let i = 2; i < count; i++) {
    assigned.push(Math.random() < 0.5 ? customerA : customerB);
  }
  shuffleArray(assigned);
  return assigned;
}

function findExecutorInfoOrThrow(executorMap, executor) {
  if (!executorMap[executor]) {
    throw new Error('Не найдены данные (ИНН, Название) для исполнителя: ' + executor);
  }
  return executorMap[executor];
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
  const maxServices = Math.min(8, servicesForExecutor.length);
  const maxAttempts = 360;

  if (servicesForExecutor.length < minServices) {
    return null;
  }

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const serviceCount = randomIntInclusive(minServices, maxServices);
    const selectedServices = pickRandomSubset(servicesForExecutor, serviceCount);
    const attemptResult = tryBuildTaskSetForSelectedServices(selectedServices, targetAmount);
    if (attemptResult) {
      return attemptResult;
    }
  }

  return null;
}

function tryBuildTaskSetForSelectedServices(selectedServices, targetAmount) {
  const items = [];
  let currentSum = 0;

  for (let i = 0; i < selectedServices.length; i++) {
    const service = selectedServices[i];
    const remainingCount = selectedServices.length - i;
    const remainingAmount = targetAmount - currentSum;
    if (remainingAmount <= 0) {
      return null;
    }

    const minTail = sumMinTotals(selectedServices, i + 1);
    const maxTail = sumMaxTotals(selectedServices, i + 1);

    if (remainingCount === 1) {
      const exactLast = buildExactVariantForTotal(service, remainingAmount);
      if (!exactLast) {
        return null;
      }
      items.push(exactLast);
      currentSum += exactLast.total;
      continue;
    }

    if (remainingCount === 2) {
      const pair = buildPairForRemainingAmount(
        service,
        selectedServices[i + 1],
        remainingAmount
      );
      if (!pair) {
        return null;
      }
      items.push(pair[0], pair[1]);
      currentSum += pair[0].total + pair[1].total;
      break;
    }

    const minForCurrent = getMinTotalForService(service);
    const maxForCurrent = getMaxTotalForService(service);
    const minAllowedTotal = Math.max(minForCurrent, remainingAmount - maxTail);
    const maxAllowedTotal = Math.min(maxForCurrent, remainingAmount - minTail);

    if (minAllowedTotal > maxAllowedTotal) {
      return null;
    }

    const variant = pickVariantForService(service, minAllowedTotal, maxAllowedTotal);
    if (!variant) {
      return null;
    }
    items.push(variant);
    currentSum += variant.total;
  }

  return currentSum === targetAmount ? items : null;
}

function buildPairForRemainingAmount(serviceA, serviceB, remainingAmount) {
  const qtyAValues = getQtyCandidates(serviceA);
  const qtyBValues = getQtyCandidates(serviceB);

  for (let i = 0; i < qtyAValues.length; i++) {
    const qtyA = qtyAValues[i];
    const minTotalA = qtyA * serviceA.minCost;
    const maxTotalA = qtyA * serviceA.maxCost;
    const lowA = Math.max(minTotalA, remainingAmount - getMaxTotalForService(serviceB));
    const highA = Math.min(maxTotalA, remainingAmount - getMinTotalForService(serviceB));

    if (lowA > highA) {
      continue;
    }

    const priceA = pickPriceForTotalRange(serviceA, qtyA, lowA, highA, remainingAmount);
    if (priceA === null) {
      continue;
    }

    const totalA = qtyA * priceA;
    const totalB = remainingAmount - totalA;
    const variantB = buildExactVariantForTotal(serviceB, totalB);
    if (!variantB) {
      continue;
    }

    return [
      buildVariant(serviceA, qtyA, priceA),
      variantB,
    ];
  }

  return null;
}

function buildExactVariantForTotal(service, targetTotal) {
  const qtyValues = getQtyCandidates(service);
  shuffleArray(qtyValues);

  for (let i = 0; i < qtyValues.length; i++) {
    const qty = qtyValues[i];
    if (targetTotal % qty !== 0) {
      continue;
    }
    const price = targetTotal / qty;
    if (price < service.minCost || price > service.maxCost) {
      continue;
    }
    if (Math.floor(price) !== price) {
      continue;
    }
    return buildVariant(service, qty, price);
  }

  return null;
}

function pickVariantForService(service, minAllowedTotal, maxAllowedTotal) {
  const qtyValues = getQtyCandidates(service);
  shuffleArray(qtyValues);

  for (let i = 0; i < qtyValues.length; i++) {
    const qty = qtyValues[i];
    const minPriceByTotal = Math.ceil(minAllowedTotal / qty);
    const maxPriceByTotal = Math.floor(maxAllowedTotal / qty);
    const minPrice = Math.max(service.minCost, minPriceByTotal);
    const maxPrice = Math.min(service.maxCost, maxPriceByTotal);

    if (minPrice > maxPrice) {
      continue;
    }

    const price = pickPriceCandidate(minPrice, maxPrice);
    if (price === null) {
      continue;
    }

    return buildVariant(service, qty, price);
  }

  return null;
}

function pickPriceForTotalRange(service, qty, minTotal, maxTotal, preferredTotal) {
  const minPriceByTotal = Math.ceil(minTotal / qty);
  const maxPriceByTotal = Math.floor(maxTotal / qty);
  const minPrice = Math.max(service.minCost, minPriceByTotal);
  const maxPrice = Math.min(service.maxCost, maxPriceByTotal);

  if (minPrice > maxPrice) {
    return null;
  }

  const preferredPrice = Math.floor(preferredTotal / qty);
  const priceCandidates = buildPriceCandidates(minPrice, maxPrice, preferredPrice);
  for (let i = 0; i < priceCandidates.length; i++) {
    const price = priceCandidates[i];
    if (price >= minPrice && price <= maxPrice) {
      return price;
    }
  }

  return null;
}

function pickPriceCandidate(minPrice, maxPrice) {
  const candidates = buildPriceCandidates(minPrice, maxPrice);
  if (candidates.length === 0) {
    return null;
  }
  return candidates[0];
}

function buildPriceCandidates(minPrice, maxPrice, preferredPrice) {
  const unique = {};
  const list = [];

  function push(v) {
    if (v < minPrice || v > maxPrice) {
      return;
    }
    if (!unique[v]) {
      unique[v] = true;
      list.push(v);
    }
  }

  push(minPrice);
  push(maxPrice);
  push(Math.floor((minPrice + maxPrice) / 2));
  if (preferredPrice !== undefined && preferredPrice !== null) {
    push(preferredPrice);
    push(preferredPrice - 1);
    push(preferredPrice + 1);
  }

  const randomCount = Math.min(4, Math.max(0, maxPrice - minPrice - 1));
  for (let i = 0; i < randomCount; i++) {
    push(randomIntInclusive(minPrice, maxPrice));
  }

  shuffleArray(list);
  return list;
}

function getQtyCandidates(service) {
  if (service.unit === 'Штука') {
    return [1, 2, 3, 4, 5];
  }
  return [1];
}

function getMinTotalForService(service) {
  return service.minCost;
}

function getMaxTotalForService(service) {
  const maxQty = service.unit === 'Штука' ? 5 : 1;
  return maxQty * service.maxCost;
}

function sumMinTotals(services, startIndex) {
  let total = 0;
  for (let i = startIndex; i < services.length; i++) {
    total += getMinTotalForService(services[i]);
  }
  return total;
}

function sumMaxTotals(services, startIndex) {
  let total = 0;
  for (let i = startIndex; i < services.length; i++) {
    total += getMaxTotalForService(services[i]);
  }
  return total;
}

function buildVariant(service, qty, price) {
  return {
    serviceId: service.id,
    description: service.description,
    category: service.category,
    unit: service.unit,
    qty: qty,
    price: price,
    total: qty * price,
  };
}

function pickRandomSubset(arr, size) {
  const shuffled = shuffleArray(arr.slice());
  return shuffled.slice(0, size);
}

function randomIntInclusive(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function buildRegistryRow(lastCol, registryHeaders, payload) {
  const row = new Array(lastCol).fill('');
  row[registryHeaders['Месяц']] = payload.month;
  row[registryHeaders['ИНН']] = payload.inn;
  row[registryHeaders['ФИО']] = payload.fullName;
  row[registryHeaders['Название']] = payload.title;
  row[registryHeaders['Период ОТ']] = payload.periodFrom;
  row[registryHeaders['Период ДО']] = payload.periodTo;
  row[registryHeaders['Заказчик']] = payload.customer;
  row[registryHeaders['Описание услуги']] = payload.description;
  row[registryHeaders['Категория услуги']] = payload.category;
  row[registryHeaders['Единица']] = payload.unit;
  row[registryHeaders['Кол-во']] = payload.qty;
  row[registryHeaders['Цена']] = payload.price;
  row[registryHeaders['Стоимость']] = payload.total;
  return row;
}

/**
 * Утверждает задания из листа "Реестр":
 * - переносит данные в "История" с проверкой конфликтов по ключу ИНН+Месяц;
 * - формирует XLSX-файлы по каждому заказчику.
 */
function approveTaskRegistry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const registrySheet = getSheetByNameOrThrow(ss, 'Реестр');
  const historySheet = getHistorySheetOrThrow();

  const registryData = getSheetDataWithHeaders(registrySheet, [
    'Месяц',
    'ИНН',
    'ФИО',
    'Название',
    'Заказчик',
    'Период ОТ',
    'Период ДО',
    'Категория услуги',
    'Описание услуги',
    'Единица',
    'Цена',
    'Кол-во',
    'Стоимость',
  ]);
  const historyData = getSheetDataWithHeaders(historySheet, ['Месяц', 'ИНН', 'ФИО', 'Заказчик']);

  const filledRegistryRows = getFilledRegistryRows(registryData.rows, registryData.headerMap);
  if (filledRegistryRows.length === 0) {
    throw new Error('В листе "Реестр" нет строк для утверждения.');
  }

  const conflictingKeys = findConflictingKeys(historyData.rows, historyData.headerMap, filledRegistryRows, registryData.headerMap);
  if (Object.keys(conflictingKeys).length > 0) {
    const response = ui.alert(
      'Обнаружены конфликты',
      'В "История" уже есть записи по некоторым парам ИНН + Месяц. Заменить существующие записи?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
    replaceHistoryRowsByConflictingKeys(historySheet, historyData.headerMap, conflictingKeys, filledRegistryRows, registryData.headerMap);
  } else {
    appendRegistryRowsToHistory(historySheet, filledRegistryRows, registryData.headerMap);
  }

  createXlsxFilesByCustomer(filledRegistryRows, registryData.headerMap);

  const summaryMessage = buildApprovalSummaryMessage(filledRegistryRows, registryData.headerMap);
  ui.alert('Задание утверждено', summaryMessage, ui.ButtonSet.OK);
}

function getFilledRegistryRows(rows, headerMap) {
  const result = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const month = toTrimmedString(row[headerMap['Месяц']]);
    const inn = toTrimmedString(row[headerMap['ИНН']]);
    const fullName = toTrimmedString(row[headerMap['ФИО']]);
    const customer = toTrimmedString(row[headerMap['Заказчик']]);
    if (!month && !inn && !fullName && !customer) {
      continue;
    }
    if (!month || !inn) {
      throw new Error('В листе "Реестр" обнаружена строка с неполным ключом ИНН + Месяц.');
    }
    result.push(row);
  }
  return result;
}

function buildInnMonthKey(inn, month) {
  return toTrimmedString(inn) + '||' + toTrimmedString(month);
}

function findConflictingKeys(historyRows, historyHeaderMap, registryRows, registryHeaderMap) {
  const historyKeys = {};
  for (let i = 0; i < historyRows.length; i++) {
    const row = historyRows[i];
    const month = toTrimmedString(row[historyHeaderMap['Месяц']]);
    const inn = toTrimmedString(row[historyHeaderMap['ИНН']]);
    if (!month || !inn) {
      continue;
    }
    historyKeys[buildInnMonthKey(inn, month)] = true;
  }

  const conflicts = {};
  for (let i = 0; i < registryRows.length; i++) {
    const row = registryRows[i];
    const key = buildInnMonthKey(row[registryHeaderMap['ИНН']], row[registryHeaderMap['Месяц']]);
    if (historyKeys[key]) {
      conflicts[key] = true;
    }
  }
  return conflicts;
}

function replaceHistoryRowsByConflictingKeys(historySheet, historyHeaderMap, conflictingKeys, registryRows, registryHeaderMap) {
  const lastCol = historySheet.getLastColumn();
  const existingRows = historySheet.getLastRow() > 1
    ? historySheet.getRange(2, 1, historySheet.getLastRow() - 1, lastCol).getValues()
    : [];
  const keptRows = [];

  for (let i = 0; i < existingRows.length; i++) {
    const row = existingRows[i];
    const key = buildInnMonthKey(row[historyHeaderMap['ИНН']], row[historyHeaderMap['Месяц']]);
    if (!conflictingKeys[key]) {
      keptRows.push(row);
    }
  }

  const newHistoryRows = mapRegistryRowsToHistoryRows(historySheet, registryRows, registryHeaderMap);
  const finalRows = keptRows.concat(newHistoryRows);

  if (historySheet.getLastRow() > 1) {
    historySheet.getRange(2, 1, historySheet.getLastRow() - 1, lastCol).clearContent();
  }
  if (finalRows.length > 0) {
    historySheet.getRange(2, 1, finalRows.length, lastCol).setValues(finalRows);
  }
}

function appendRegistryRowsToHistory(historySheet, registryRows, registryHeaderMap) {
  const mappedRows = mapRegistryRowsToHistoryRows(historySheet, registryRows, registryHeaderMap);
  if (mappedRows.length === 0) {
    return;
  }
  const startRow = historySheet.getLastRow() + 1;
  historySheet.getRange(startRow, 1, mappedRows.length, historySheet.getLastColumn()).setValues(mappedRows);
}

function mapRegistryRowsToHistoryRows(historySheet, registryRows, registryHeaderMap) {
  const historyHeaders = getHeaderMapOrThrow(historySheet, ['Месяц', 'ИНН', 'ФИО', 'Заказчик']);
  const result = [];
  for (let i = 0; i < registryRows.length; i++) {
    const registryRow = registryRows[i];
    const row = new Array(historySheet.getLastColumn()).fill('');
    fillIfPresent(row, historyHeaders, 'Месяц', registryRow, registryHeaderMap, 'Месяц');
    fillIfPresent(row, historyHeaders, 'ИНН', registryRow, registryHeaderMap, 'ИНН');
    fillIfPresent(row, historyHeaders, 'ФИО', registryRow, registryHeaderMap, 'ФИО');
    fillIfPresent(row, historyHeaders, 'Название', registryRow, registryHeaderMap, 'Название');
    fillIfPresent(row, historyHeaders, 'Заказчик', registryRow, registryHeaderMap, 'Заказчик');
    fillIfPresent(row, historyHeaders, 'Период ОТ', registryRow, registryHeaderMap, 'Период ОТ');
    fillIfPresent(row, historyHeaders, 'Период ДО', registryRow, registryHeaderMap, 'Период ДО');
    fillIfPresent(row, historyHeaders, 'Категория услуги', registryRow, registryHeaderMap, 'Категория услуги');
    fillIfPresent(row, historyHeaders, 'Описание услуги', registryRow, registryHeaderMap, 'Описание услуги');
    fillIfPresent(row, historyHeaders, 'Единица', registryRow, registryHeaderMap, 'Единица');
    fillIfPresent(row, historyHeaders, 'Цена', registryRow, registryHeaderMap, 'Цена');
    fillIfPresent(row, historyHeaders, 'Кол-во', registryRow, registryHeaderMap, 'Кол-во');
    fillIfPresent(row, historyHeaders, 'Стоимость', registryRow, registryHeaderMap, 'Стоимость');
    result.push(row);
  }
  return result;
}

function fillIfPresent(targetRow, targetHeaderMap, targetHeader, sourceRow, sourceHeaderMap, sourceHeader) {
  if (targetHeaderMap[targetHeader] === undefined || sourceHeaderMap[sourceHeader] === undefined) {
    return;
  }
  targetRow[targetHeaderMap[targetHeader]] = sourceRow[sourceHeaderMap[sourceHeader]];
}

function createXlsxFilesByCustomer(registryRows, registryHeaderMap) {
  const folderId = '1CtIBNkjqSLfYfNruhAKJFGccqujCxK4x';
  const folder = DriveApp.getFolderById(folderId);
  const grouped = groupRegistryRowsByCustomer(registryRows, registryHeaderMap);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const customers = Object.keys(grouped);
  for (let i = 0; i < customers.length; i++) {
    const customer = customers[i];
    const rows = grouped[customer];
    if (rows.length === 0) {
      continue;
    }
    const month = toTrimmedString(rows[0][registryHeaderMap['Месяц']]);
    const customerForFileName = stripLeadingOoo(customer);
    const fileName = sanitizeFileName(customerForFileName + '_' + month + '_Задание_Консоль_' + today + '.xlsx');
    const exportRows = buildExportRows(rows, registryHeaderMap, fileName);
    saveRowsAsXlsx(folder, fileName, exportRows);
  }
}

function buildApprovalSummaryMessage(registryRows, headerMap) {
  const byExecutor = {};

  for (let i = 0; i < registryRows.length; i++) {
    const row = registryRows[i];
    const executor = toTrimmedString(row[headerMap['ФИО']]);
    const customer = toTrimmedString(row[headerMap['Заказчик']]);
    const cost = Number(row[headerMap['Стоимость']]);

    if (!executor) {
      continue;
    }
    if (!byExecutor[executor]) {
      byExecutor[executor] = {
        total: 0,
        customers: {},
      };
    }

    if (Number.isFinite(cost)) {
      byExecutor[executor].total += cost;
    }
    if (customer) {
      byExecutor[executor].customers[customer] = true;
    }
  }

  const executors = Object.keys(byExecutor).sort();
  if (executors.length === 0) {
    return 'Утверждение выполнено.';
  }

  const lines = [];
  for (let i = 0; i < executors.length; i++) {
    const executor = executors[i];
    const item = byExecutor[executor];
    const customers = Object.keys(item.customers).sort();
    const customersText = customers.length > 0 ? customers.join(', ') : '—';
    lines.push(
      executor + ': назначено заданий на сумму ' + item.total + ' в организациях: ' + customersText + '.'
    );
  }

  return lines.join('\n');
}

function groupRegistryRowsByCustomer(registryRows, registryHeaderMap) {
  const grouped = {};
  for (let i = 0; i < registryRows.length; i++) {
    const row = registryRows[i];
    const customer = toTrimmedString(row[registryHeaderMap['Заказчик']]);
    if (!customer) {
      continue;
    }
    if (!grouped[customer]) {
      grouped[customer] = [];
    }
    grouped[customer].push(row);
  }
  return grouped;
}

function buildExportRows(rows, headerMap, fileName) {
  const exportHeaders = [
    'Телефон',
    'ИНН',
    'ФИО',
    'Название',
    'Проект',
    'Локация',
    'Удаленно',
    'Период ОТ',
    'Период ДО',
    'Время ОТ',
    'Время ДО',
    'Комментарий для исп-ля',
    'Внутр. комментарий',
    'Категория услуги',
    'Описание услуги',
    'Единица',
    'Цена',
    'Кол-во',
    'Стоимость',
    'Теги',
    'Перевод задания в Предложено',
    'Разовое задание',
    'Тип оплаты для разового',
    'Сценарий приглашения',
  ];

  const out = [exportHeaders];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    out.push([
      '',
      row[headerMap['ИНН']],
      row[headerMap['ФИО']],
      row[headerMap['Название']],
      '',
      '',
      '',
      row[headerMap['Период ОТ']],
      row[headerMap['Период ДО']],
      '',
      '',
      '',
      fileName,
      row[headerMap['Категория услуги']],
      row[headerMap['Описание услуги']],
      row[headerMap['Единица']],
      row[headerMap['Цена']],
      row[headerMap['Кол-во']],
      row[headerMap['Стоимость']],
      '',
      'да',
      'нет',
      'постоплата',
      '',
    ]);
  }
  return out;
}

function saveRowsAsXlsx(folder, fileName, values) {
  const tempSpreadsheet = SpreadsheetApp.create('temp_export_' + new Date().getTime());
  try {
    const sheet = tempSpreadsheet.getSheets()[0];
    sheet.clear();
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    SpreadsheetApp.flush();

    const url = 'https://docs.google.com/spreadsheets/d/' + tempSpreadsheet.getId() + '/export?format=xlsx';
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: false,
    });
    const blob = response.getBlob().setName(fileName);
    folder.createFile(blob);
  } finally {
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
  }
}

function sanitizeFileName(name) {
  return name
    .replace(/[\\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^_+/, '');
}

function stripLeadingOoo(name) {
  const normalized = toTrimmedString(name).replace(/^ООО\s+/i, '').trim();
  return normalized || toTrimmedString(name);
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

function parseCommaSeparatedValues(value) {
  return toTrimmedString(value)
    .split(/[,\n;\r]+/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length > 0; });
}

function normalizeCustomerNameKey(value) {
  return toTrimmedString(value)
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .toLowerCase();
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
