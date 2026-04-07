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
  const enterprisesSheet = getSheetByNameOrThrow(ss, 'Предприятия');
  const historySheet = getSheetByNameOrThrow(ss, 'История');
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
  const enterprisesData = getSheetDataWithHeaders(enterprisesSheet, ['Заказчик']);
  const historyData = getSheetDataWithHeaders(historySheet, ['Месяц', 'Исполнитель', 'Заказчик']);
  const registryHeaders = getHeaderMapOrThrow(registrySheet, [
    'ИНН',
    'Период ОТ',
    'Период ДО',
    'Заказчик',
    'Описание услуги',
    'Категория услуги',
    'Кол-во',
    'Цена',
    'Стоимость',
  ]);

  clearRegistryData(registrySheet);

  const innMap = buildInnMap(executorsData.rows, executorsData.headerMap);
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
    const inn = findInnByExecutor(innMap, executor);
    const servicesForExecutor = filterServicesForExecutor(services, executor);

    if (servicesForExecutor.length === 0) {
      throw new Error('Для исполнителя ' + executor + ' не найдено ни одной подходящей услуги.');
    }

    const taskSet = buildTaskSetForAmount(servicesForExecutor, amount);
    if (!taskSet) {
      throw new Error('Не удалось подобрать услуги на сумму ' + amount + ' для исполнителя ' + executor + '.');
    }
    const bannedCustomers = buildRecentCustomerSet(
      historyData.rows,
      historyData.headerMap,
      executor,
      monthCodeRaw
    );
    const pickedCustomers = pickTwoCustomers(customerPool, bannedCustomers, executor, monthCodeRaw);
    const assignedCustomers = assignCustomersToTaskItems(taskSet, pickedCustomers[0], pickedCustomers[1]);

    for (let j = 0; j < taskSet.length; j++) {
      const item = taskSet[j];
      outputRows.push(buildRegistryRow(registrySheet.getLastColumn(), registryHeaders, {
        inn: inn,
        periodFrom: period.from,
        periodTo: period.to,
        customer: assignedCustomers[j],
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

function buildRecentCustomerSet(historyRows, historyHeaderMap, executor, monthCode) {
  const previousMonths = getPreviousMonthCodes(monthCode, 2);
  const monthMap = {};
  for (let i = 0; i < previousMonths.length; i++) {
    monthMap[previousMonths[i]] = true;
  }

  const recentCustomers = {};
  for (let i = 0; i < historyRows.length; i++) {
    const row = historyRows[i];
    const rowExecutor = toTrimmedString(row[historyHeaderMap['Исполнитель']]);
    const rowMonth = toTrimmedString(row[historyHeaderMap['Месяц']]);
    const rowCustomer = toTrimmedString(row[historyHeaderMap['Заказчик']]);

    if (!rowExecutor || !rowMonth || !rowCustomer) {
      continue;
    }
    if (rowExecutor !== executor) {
      continue;
    }
    if (!monthMap[rowMonth]) {
      continue;
    }

    recentCustomers[rowCustomer] = true;
  }
  return recentCustomers;
}

function pickTwoCustomers(availableCustomers, bannedCustomers, executor, monthCode) {
  const allowed = [];

  for (let i = 0; i < availableCustomers.length; i++) {
    const customer = availableCustomers[i];
    if (!bannedCustomers[customer]) {
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
  row[registryHeaders['ИНН']] = payload.inn;
  row[registryHeaders['Период ОТ']] = payload.periodFrom;
  row[registryHeaders['Период ДО']] = payload.periodTo;
  row[registryHeaders['Заказчик']] = payload.customer;
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
