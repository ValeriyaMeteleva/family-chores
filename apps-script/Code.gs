const SHEETS = {
  members: { name: "Участники", headers: ["id", "name", "role", "inviteHash"] },
  tasks: { name: "Дела", headers: ["id", "title", "price", "active", "visibleTo"] },
  requests: { name: "Заявки", headers: ["id", "memberId", "taskId", "taskTitle", "price", "comment", "createdAt", "status", "parentComment"] },
  payments: { name: "Выплаты", headers: ["id", "memberId", "amount", "status", "createdAt", "paidAt", "confirmedAt"] },
  funds: { name: "Финансы", headers: ["memberId", "debtTarget", "debtRemaining", "debtPaidTotal", "debtRate", "savingsTarget", "savingsRate", "savingsBalance", "savingsTotal", "savingsWithdrawn"] },
  savingsRequests: { name: "Запросы накоплений", headers: ["id", "memberId", "amount", "status", "createdAt", "paidAt", "confirmedAt"] }
};

const PROPERTY_KEYS = {
  familyHash: "FAMILY_CODE_HASH",
  familySalt: "FAMILY_CODE_SALT",
  pinHash: "PARENT_PIN_HASH",
  pinSalt: "PARENT_PIN_SALT",
  failedPinCount: "FAILED_PIN_COUNT",
  blockedUntil: "PIN_BLOCKED_UNTIL",
  revision: "FAMILY_REVISION"
};
const SESSION_PREFIX = "family-session:";
const SESSION_TTL_SECONDS = 21600;
const MAX_BAD_PIN_ATTEMPTS = 5;
const PIN_LOCKOUT_MS = 15 * 60 * 1000;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Наши дела")
    .addItem("Настроить сервер приложения", "initializeFamily")
    .addToUi();
}

function initializeFamily() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(PROPERTY_KEYS.familyHash)) {
    ui.alert("Сервер уже настроен. Повторный запуск не изменит PIN или семейный код.");
    return;
  }

  const response = ui.prompt(
    "Настройка семейного приложения",
    "Задайте четырёхзначный PIN для родителей. Не используйте PIN от банков или аккаунтов.",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const pin = response.getResponseText().trim();
  if (!/^\d{4}$/.test(pin)) {
    ui.alert("PIN должен состоять ровно из четырёх цифр. Запустите настройку ещё раз.");
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (properties.getProperty(PROPERTY_KEYS.familyHash)) {
      ui.alert("Сервер уже настроен.");
      return;
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(SHEETS).forEach(key => ensureSheet_(spreadsheet, SHEETS[key]));

    const salt = randomToken_();
    const pinSalt = randomToken_();
    const familyCode = randomToken_().replace(/-/g, "").slice(0, 20).toUpperCase();
    properties.setProperties({
      [PROPERTY_KEYS.familySalt]: salt,
      [PROPERTY_KEYS.familyHash]: digest_(familyCode, salt),
      [PROPERTY_KEYS.pinSalt]: pinSalt,
      [PROPERTY_KEYS.pinHash]: digest_(pin, pinSalt),
      [PROPERTY_KEYS.failedPinCount]: "0",
      [PROPERTY_KEYS.blockedUntil]: "0",
      [PROPERTY_KEYS.revision]: "0"
    });

    const tasks = spreadsheet.getSheetByName(SHEETS.tasks.name);
    if (tasks.getLastRow() === 1) {
      [
        ["dishes", "Помыть посуду", 20, true],
        ["room", "Убрать комнату", 40, true],
        ["plants", "Полить цветы", 15, true],
        ["table", "Накрыть на стол", 10, true]
      ].forEach(row => tasks.appendRow(row));
    }

    const html = HtmlService.createHtmlOutput(
      `<div style="font:16px Arial,sans-serif;line-height:1.55;padding:12px">
        <h2>Сервер готов</h2>
        <p>Сохраните семейный код в надёжном месте. Он показывается только один раз и понадобится родителям и при подключении детей.</p>
        <p style="font-size:22px;font-weight:bold;letter-spacing:2px">${familyCode}</p>
        <p>PIN хранится в серверных свойствах в виде хеша. Не отправляйте этот код в публичный чат.</p>
      </div>`
    ).setWidth(440).setHeight(260);
    ui.showModalDialog(html, "Семейный код");
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Bridge")
    .setTitle("Family Chores API bridge")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiDispatch(action, payload) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("Некорректный запрос.");
  }
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(key => ensureSheet_(spreadsheet, SHEETS[key]));
  switch (action) {
    case "parentLogin": return parentLogin_(payload);
    case "childLogin": return childLogin_(payload);
    case "readState": return readState_(payload);
    case "verifyPin": return verifyPin_(payload);
    case "syncState": return syncState_(payload);
    case "addMember": return addMember_(payload);
    default: throw new Error("Неизвестная операция.");
  }
}

function parentLogin_(payload) {
  const properties = PropertiesService.getScriptProperties();
  const blockedUntil = Number(properties.getProperty(PROPERTY_KEYS.blockedUntil) || 0);
  if (Date.now() < blockedUntil) {
    throw new Error("Слишком много неверных PIN. Попробуйте позже.");
  }

  const familyCode = cleanString_(payload.familyCode, 40);
  const pin = cleanString_(payload.pin, 4);
  const name = cleanString_(payload.name, 40);
  if (!familyCode || !/^\d{4}$/.test(pin) || !name) {
    throw new Error("Введите семейный код, имя и четырёхзначный PIN.");
  }

  const expectedFamilyHash = properties.getProperty(PROPERTY_KEYS.familyHash);
  const familySalt = properties.getProperty(PROPERTY_KEYS.familySalt);
  if (!expectedFamilyHash || !constantTimeEquals_(digest_(familyCode.toUpperCase(), familySalt), expectedFamilyHash)) {
    throw new Error("Семейный код не подошёл.");
  }

  const pinSalt = properties.getProperty(PROPERTY_KEYS.pinSalt);
  const expectedPinHash = properties.getProperty(PROPERTY_KEYS.pinHash);
  if (!constantTimeEquals_(digest_(pin, pinSalt), expectedPinHash)) {
    registerBadPin_(properties);
    throw new Error("PIN не подошёл.");
  }
  properties.setProperties({
    [PROPERTY_KEYS.failedPinCount]: "0",
    [PROPERTY_KEYS.blockedUntil]: "0"
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const members = readRows_(SHEETS.members);
    let member = members.find(item => item.role === "parent" && item.name.toLocaleLowerCase() === name.toLocaleLowerCase());
    if (!member) {
      if (members.filter(item => item.role === "parent").length >= 2) {
        throw new Error("В семье уже зарегистрированы два родителя.");
      }
      member = { id: randomToken_(), name: name, role: "parent", inviteHash: "" };
      appendObject_(SHEETS.members, member);
      appendFund_(member.id);
      bumpRevision_();
    }
    return createSession_(member);
  } finally {
    lock.releaseLock();
  }
}

function childLogin_(payload) {
  const properties = PropertiesService.getScriptProperties();
  const familyCode = cleanString_(payload.familyCode, 40);
  const inviteCode = cleanString_(payload.inviteCode, 80);
  if (!familyCode || !inviteCode) throw new Error("Введите семейный код и код профиля ребёнка.");

  const familySalt = properties.getProperty(PROPERTY_KEYS.familySalt);
  const expectedFamilyHash = properties.getProperty(PROPERTY_KEYS.familyHash);
  if (!expectedFamilyHash || !constantTimeEquals_(digest_(familyCode.toUpperCase(), familySalt), expectedFamilyHash)) {
    throw new Error("Семейный код не подошёл.");
  }

  const member = readRows_(SHEETS.members).find(item =>
    item.role === "child" &&
    constantTimeEquals_(digest_(inviteCode.toUpperCase(), item.inviteHashSalt), item.inviteHash)
  );
  if (!member) throw new Error("Код профиля ребёнка не подошёл.");
  return createSession_(member);
}

function readState_(payload) {
  const session = requireSession_(payload.token);
  return {
    memberId: session.memberId,
    role: session.role,
    state: stateForMember_(readStateData_(), session),
    version: currentRevision_()
  };
}

function verifyPin_(payload) {
  requireSession_(payload.token, "parent");
  if (!verifyParentPin_(payload.pin)) throw new Error("PIN не подошёл.");
  return { verified: true };
}

function addMember_(payload) {
  const session = requireSession_(payload.token, "parent");
  if (!verifyParentPin_(payload.pin)) throw new Error("PIN не подошёл.");
  const name = cleanString_(payload.name, 40);
  const role = cleanString_(payload.role, 10);
  if (!name || (role !== "parent" && role !== "child")) throw new Error("Введите имя и выберите роль.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const members = readRows_(SHEETS.members);
    if (members.filter(item => item.role === role).length >= 2) {
      throw new Error(`В семье уже зарегистрированы два участника с ролью «${role === "parent" ? "родитель" : "ребёнок"}».`);
    }
    const inviteCode = role === "child" ? randomToken_().replace(/-/g, "").slice(0, 16).toUpperCase() : "";
    const inviteHashSalt = role === "child" ? randomToken_() : "";
    const member = {
      id: randomToken_(),
      name: name,
      role: role,
      inviteHashSalt: inviteHashSalt,
      inviteHash: role === "child" ? digest_(inviteCode, inviteHashSalt) : ""
    };
    appendObject_(SHEETS.members, member);
    appendFund_(member.id);
    const version = bumpRevision_();
    return {
      member: { id: member.id, name: member.name, role: member.role },
      inviteCode: inviteCode,
      state: readStateData_(),
      version: version
    };
  } finally {
    lock.releaseLock();
  }
}

function syncState_(payload) {
  const session = requireSession_(payload.token);
  if (session.role === "parent" && !verifyParentPin_(payload.pin)) {
    throw new Error("PIN не подошёл. Изменения не сохранены.");
  }
  const candidate = payload.state;
  if (!candidate || typeof candidate !== "object") throw new Error("Нет данных для сохранения.");
  ["members", "tasks", "requests", "payments", "funds", "savingsRequests"].forEach(key => {
    if (!Array.isArray(candidate[key]) || candidate[key].length > 10000) {
      throw new Error("Некорректный формат данных.");
    }
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (Number(payload.version) !== currentRevision_()) {
      throw new Error("Данные уже изменились на другом устройстве. Обновляю семейную базу; повторите действие.");
    }
    const current = readStateData_();
    if (session.role === "parent") {
      validateParentState_(current, candidate);
      applyApprovedChoreAllocations_(current, candidate);
      validateFundTotals_(candidate);
      writeState_(candidate);
    } else {
      const currentForChild = stateForMember_(current, session);
      validateChildState_(currentForChild, candidate, session.memberId);
      const merged = mergeChildState_(current, candidate, session.memberId);
      applySavingsConfirmations_(current, merged);
      validateFundTotals_(merged);
      writeState_(merged);
    }
    const version = bumpRevision_();
    return {
      memberId: session.memberId,
      role: session.role,
      state: stateForMember_(readStateData_(), session),
      version: version
    };
  } finally {
    lock.releaseLock();
  }
}

function validateParentState_(current, next) {
  if (new Set(next.members.map(item => item.id)).size !== next.members.length ||
      new Set(next.tasks.map(item => item.id)).size !== next.tasks.length ||
      new Set(next.requests.map(item => item.id)).size !== next.requests.length ||
      new Set(next.payments.map(item => item.id)).size !== next.payments.length ||
      new Set(next.savingsRequests.map(item => item.id)).size !== next.savingsRequests.length) {
    throw new Error("В данных повторяются идентификаторы.");
  }
  if (next.requests.some(request => typeof request.parentComment !== "string" || request.parentComment.length > 500)) {
    throw new Error("Комментарий родителя слишком длинный.");
  }
  if (!sameJson_(current.members, next.members)) throw new Error("Изменение профилей доступно только через добавление участника.");
  validateRows_(current.requests, next.requests, item => item.status === "pending", ["status", "parentComment"], "заявки");
  validateRows_(current.payments, next.payments, item => item.status === "requested", ["status", "paidAt"], "выплаты");
  validateRows_(current.savingsRequests, next.savingsRequests, item => item.status === "requested",
    ["status", "paidAt"], "запросы накоплений");
  validateParentFunds_(current.funds, next.funds);

  const existingTasks = new Map(current.tasks.map(task => [task.id, task]));
  const nextTaskIds = new Set();
  next.tasks.forEach(task => {
    validateTask_(task);
    if (nextTaskIds.has(task.id)) throw new Error("Повторяется идентификатор дела.");
    nextTaskIds.add(task.id);
    const previous = existingTasks.get(task.id);
    if (previous) {
      if (Number(task.price) < 0 || !task.title) throw new Error("Некорректные данные дела.");
    }
    if (task.visibleTo && !current.members.some(member => member.id === task.visibleTo && member.role === "child")) {
      throw new Error("Для дела выбран неизвестный профиль ребёнка.");
    }
  });
  if (current.tasks.some(task => !nextTaskIds.has(task.id))) {
    throw new Error("Удалять дела с историей нельзя; отключите их.");
  }

  current.requests.forEach(request => {
    const candidate = next.requests.find(item => item.id === request.id);
    if (request.status === "pending" && candidate &&
        candidate.status !== "pending" && candidate.status !== "approved" && candidate.status !== "rejected") {
      throw new Error("Недопустимый статус заявки.");
    }
  });
  current.payments.forEach(payment => {
    const candidate = next.payments.find(item => item.id === payment.id);
    if (payment.status === "requested" && candidate &&
        candidate.status !== "requested" && candidate.status !== "awaiting_child_confirmation") {
      throw new Error("Недопустимый статус выплаты.");
    }
  });
  current.savingsRequests.forEach(request => {
    const candidate = next.savingsRequests.find(item => item.id === request.id);
    if (request.status === "requested" && candidate &&
        candidate.status !== "requested" && candidate.status !== "awaiting_child_confirmation" &&
        candidate.status !== "rejected") {
      throw new Error("Недопустимый статус запроса накоплений.");
    }
  });

  const earnedByMember = new Map();
  next.requests.filter(request => request.status === "approved").forEach(request => {
    earnedByMember.set(
      request.memberId,
      (earnedByMember.get(request.memberId) || 0) + cents_(request.price)
    );
  });
  const paidByMember = new Map();
  next.payments
    .filter(payment => payment.status === "awaiting_child_confirmation" || payment.status === "confirmed")
    .forEach(payment => {
      paidByMember.set(
        payment.memberId,
        (paidByMember.get(payment.memberId) || 0) + cents_(payment.amount)
      );
    });
  const reservedByMember = new Map();
  next.payments.filter(payment => payment.status === "requested").forEach(payment => {
    reservedByMember.set(payment.memberId, (reservedByMember.get(payment.memberId) || 0) + cents_(payment.amount));
  });
  next.funds.forEach(fund => {
    const paid = paidByMember.get(fund.memberId) || 0;
    if (paid + (reservedByMember.get(fund.memberId) || 0) + cents_(fund.debtPaidTotal) +
        cents_(fund.savingsTotal) > (earnedByMember.get(fund.memberId) || 0)) {
      throw new Error("Выплаты и накопления превышают заработанную сумму.");
    }
  });
}

function validateChildState_(current, next, memberId) {
  if (!sameJson_(current.members, next.members) || !sameJson_(current.tasks, next.tasks)) {
    throw new Error("Детский профиль не может редактировать участников и список дел.");
  }
  if (new Set(next.requests.map(item => item.id)).size !== next.requests.length ||
      new Set(next.payments.map(item => item.id)).size !== next.payments.length ||
      new Set(next.savingsRequests.map(item => item.id)).size !== next.savingsRequests.length) {
    throw new Error("В данных повторяются идентификаторы.");
  }
  validateChildFunds_(current.funds, next.funds, memberId);

  const oldRequests = new Map(current.requests.map(request => [request.id, request]));
  const nextRequests = new Map(next.requests.map(request => [request.id, request]));
  current.requests.forEach(request => {
    const candidate = nextRequests.get(request.id);
    if (!candidate) {
      if (request.memberId !== memberId || request.status !== "pending") {
        throw new Error("Удалять можно только свои ожидающие заявки.");
      }
    } else if (!sameJson_(request, candidate)) {
      throw new Error("Нельзя менять существующие заявки.");
    }
  });
  next.requests.forEach(request => {
    if (oldRequests.has(request.id)) return;
    const task = current.tasks.find(item => item.id === request.taskId && item.active);
    if (!task || request.memberId !== memberId || request.status !== "pending" ||
        request.taskTitle !== task.title || Number(request.price) !== Number(task.price) ||
        typeof request.comment !== "string" || request.comment.length > 500 ||
        typeof request.createdAt !== "string" || request.createdAt.length > 80) {
      throw new Error("Некорректная заявка на дело.");
    }
  });

  const oldPayments = new Map(current.payments.map(payment => [payment.id, payment]));
  const nextPayments = new Map(next.payments.map(payment => [payment.id, payment]));
  current.payments.forEach(payment => {
    const candidate = nextPayments.get(payment.id);
    if (!candidate) throw new Error("Удалять записи о выплатах нельзя.");
    if (!sameJson_(payment, candidate)) {
      const confirming = payment.memberId === memberId &&
        payment.status === "awaiting_child_confirmation" &&
        candidate.status === "confirmed" &&
        candidate.memberId === payment.memberId &&
        Number(candidate.amount) === Number(payment.amount) &&
        candidate.paidAt === payment.paidAt &&
        typeof candidate.confirmedAt === "string" && candidate.confirmedAt.length <= 80 &&
        sameExceptFields_(payment, candidate, ["status", "confirmedAt"]);
      if (!confirming) throw new Error("Некорректное изменение выплаты.");
    }
  });
  const oldSavingsRequests = new Map(current.savingsRequests.map(item => [item.id, item]));
  const nextSavingsRequests = new Map(next.savingsRequests.map(item => [item.id, item]));
  current.savingsRequests.forEach(request => {
    const candidate = nextSavingsRequests.get(request.id);
    if (!candidate) throw new Error("Удалять историю запросов накоплений нельзя.");
    if (!sameJson_(request, candidate)) {
      const confirming = request.memberId === memberId &&
        request.status === "awaiting_child_confirmation" &&
        candidate.status === "confirmed" &&
        candidate.memberId === request.memberId &&
        Number(candidate.amount) === Number(request.amount) &&
        candidate.paidAt === request.paidAt &&
        typeof candidate.confirmedAt === "string" && candidate.confirmedAt.length <= 80 &&
        sameExceptFields_(request, candidate, ["status", "confirmedAt"]);
      if (!confirming) throw new Error("Некорректное изменение запроса накоплений.");
    }
  });
  const fund = next.funds.find(item => item.memberId === memberId) || emptyFund_(memberId);
  const reservedSavings = next.savingsRequests
    .filter(item => item.memberId === memberId &&
      (item.status === "requested" || item.status === "awaiting_child_confirmation"))
    .reduce((sum, item) => sum + cents_(item.amount), 0);
  next.savingsRequests.forEach(request => {
    if (oldSavingsRequests.has(request.id)) return;
    const amount = cents_(request.amount);
    if (request.memberId !== memberId || request.status !== "requested" || amount <= 0 ||
        amount > cents_(fund.savingsBalance) - reservedSavings ||
        typeof request.id !== "string" || !request.id || request.id.length > 120 ||
        typeof request.createdAt !== "string" || request.createdAt.length > 80 ||
        request.paidAt !== "" || request.confirmedAt !== "") {
      throw new Error("Запрос превышает сумму доступных накоплений.");
    }
  });

  let earnedCents = current.requests
    .filter(request => request.memberId === memberId && request.status === "approved")
    .reduce((sum, request) => sum + cents_(request.price), 0);
  const currentFund = current.funds.find(item => item.memberId === memberId) || emptyFund_(memberId);
  earnedCents -= cents_(currentFund.debtPaidTotal) + cents_(currentFund.savingsTotal);
  earnedCents -= current.payments
    .filter(payment => payment.memberId === memberId &&
      (payment.status === "awaiting_child_confirmation" || payment.status === "confirmed"))
    .reduce((sum, payment) => sum + cents_(payment.amount), 0);
  let reservedCents = current.payments
    .filter(payment => payment.memberId === memberId && payment.status === "requested")
    .reduce((sum, payment) => sum + cents_(payment.amount), 0);

  next.payments.forEach(payment => {
    if (oldPayments.has(payment.id)) return;
    const amountCents = cents_(payment.amount);
    if (payment.memberId !== memberId || payment.status !== "requested" || amountCents <= 0 ||
        amountCents > earnedCents - reservedCents) {
      throw new Error("Сумма выплаты превышает доступный баланс.");
    }
    reservedCents += amountCents;
  });
}

function sameExceptFields_(current, candidate, allowedFields) {
  const keys = new Set(Object.keys(current).concat(Object.keys(candidate)));
  return Array.from(keys).every(key =>
    allowedFields.includes(key) || current[key] === candidate[key]
  ) && Object.keys(candidate).every(key =>
    Object.prototype.hasOwnProperty.call(current, key) || allowedFields.includes(key)
  );
}

function validateParentFunds_(currentFunds, nextFunds) {
  if (currentFunds.length !== nextFunds.length) throw new Error("Список финансовых счетов изменён.");
  currentFunds.forEach(current => {
    const next = nextFunds.find(item => item.memberId === current.memberId);
    if (!next) throw new Error("Финансовый счёт участника не найден.");
    validateFundShape_(next);
    ["debtPaidTotal", "debtRate", "savingsTarget", "savingsRate", "savingsBalance",
      "savingsTotal", "savingsWithdrawn"].forEach(key => {
      if (cents_(current[key]) !== cents_(next[key])) throw new Error("Изменить можно только остаток долга.");
    });
    const remaining = cents_(next.debtRemaining);
    const target = cents_(next.debtTarget);
    if (remaining < 0 || target < 0 ||
        (remaining > 0 && target !== remaining + cents_(next.debtPaidTotal)) ||
        (remaining === 0 && target !== 0 && target !== cents_(next.debtPaidTotal))) {
      throw new Error("Некорректная сумма долга.");
    }
  });
}

function validateChildFunds_(currentFunds, nextFunds, memberId) {
  if (currentFunds.length !== 1 || nextFunds.length !== 1 ||
      currentFunds[0].memberId !== memberId || nextFunds[0].memberId !== memberId) {
    throw new Error("Изменён финансовый счёт другого участника.");
  }
  const current = currentFunds[0];
  const next = nextFunds[0];
  validateFundShape_(next);
  ["debtPaidTotal", "savingsBalance", "savingsTotal", "savingsWithdrawn"].forEach(key => {
    if (cents_(current[key]) !== cents_(next[key])) {
      throw new Error("Нельзя менять уже накопленные суммы или списывать долг.");
    }
  });
  ["debtRate", "savingsTarget", "savingsRate"].forEach(key => {
    if (!Number.isFinite(Number(next[key]))) throw new Error("Некорректные настройки накоплений.");
  });
  if (Number(next.savingsTarget) < 0 || Number(next.savingsTarget) > 100000 ||
      Number(next.savingsRate) < 0 || Number(next.savingsRate) > 50 ||
      Number(next.debtRate) < 0 || Number(next.debtRate) > 50 ||
      Number(next.savingsTarget) < Number(next.savingsBalance)) {
    throw new Error("Некорректная сумма или процент накоплений.");
  }
  if (Number(next.debtRemaining) > 0 && Number(next.savingsRate) > 20) {
    throw new Error("При активном долге на накопления можно направлять не более 20% остатка.");
  }
  if (Number(current.debtRemaining) > 0) {
    if (cents_(current.debtRemaining) !== cents_(next.debtRemaining) ||
        cents_(current.debtTarget) !== cents_(next.debtTarget) ||
        cents_(current.debtRate) !== cents_(next.debtRate)) {
      throw new Error("Изменить или отменить долг может только родитель.");
    }
  } else {
    const remaining = cents_(next.debtRemaining);
    const target = cents_(next.debtTarget);
    if (remaining < 0 || remaining > 10000000 ||
        (remaining > 0 && target !== remaining + cents_(next.debtPaidTotal)) ||
        (remaining === 0 && target !== cents_(current.debtTarget))) {
      throw new Error("Некорректная сумма долга.");
    }
  }
}

function validateFundShape_(fund) {
  const amounts = ["debtTarget", "debtRemaining", "debtPaidTotal", "savingsTarget",
    "savingsBalance", "savingsTotal", "savingsWithdrawn"];
  const rates = ["debtRate", "savingsRate"];
  if (!fund || typeof fund.memberId !== "string" || !fund.memberId ||
      amounts.some(key => !Number.isFinite(Number(fund[key])) || Number(fund[key]) < 0) ||
      rates.some(key => !Number.isFinite(Number(fund[key])) || Number(fund[key]) < 0 || Number(fund[key]) > 50) ||
      Number(fund.debtRemaining) > 100000 || Number(fund.savingsTarget) > 100000 ||
      (Number(fund.debtRemaining) > 0 && Number(fund.savingsRate) > 20)) {
    throw new Error("Некорректные данные финансового счёта.");
  }
}

function applyApprovedChoreAllocations_(current, next) {
  const previous = new Map(current.requests.map(request => [request.id, request]));
  next.requests.forEach(request => {
    const old = previous.get(request.id);
    if (!old || old.status !== "pending" || request.status !== "approved") return;
    const fund = next.funds.find(item => item.memberId === request.memberId);
    if (!fund) throw new Error("Не найден финансовый счёт ребёнка.");
    const earned = cents_(request.price);
    const debtPart = Math.min(cents_(fund.debtRemaining), Math.round(earned * Number(fund.debtRate) / 100));
    const afterDebt = earned - debtPart;
    const savingsPart = Number(fund.savingsTarget) > Number(fund.savingsBalance)
      ? Math.min(
        cents_(fund.savingsTarget) - cents_(fund.savingsBalance),
        Math.round(afterDebt * Number(fund.savingsRate) / 100)
      )
      : 0;
    fund.debtRemaining = (cents_(fund.debtRemaining) - debtPart) / 100;
    fund.debtPaidTotal = (cents_(fund.debtPaidTotal) + debtPart) / 100;
    fund.savingsBalance = (cents_(fund.savingsBalance) + savingsPart) / 100;
    fund.savingsTotal = (cents_(fund.savingsTotal) + savingsPart) / 100;
    if (fund.debtRemaining === 0) fund.debtTarget = fund.debtPaidTotal;
  });
}

function applySavingsConfirmations_(current, next) {
  const previous = new Map(current.savingsRequests.map(request => [request.id, request]));
  next.savingsRequests.forEach(request => {
    const old = previous.get(request.id);
    if (!old || old.status !== "awaiting_child_confirmation" || request.status !== "confirmed") return;
    const fund = next.funds.find(item => item.memberId === request.memberId);
    if (!fund || cents_(request.amount) > cents_(fund.savingsBalance)) {
      throw new Error("Недостаточно средств в накоплениях для подтверждения выплаты.");
    }
    fund.savingsBalance = (cents_(fund.savingsBalance) - cents_(request.amount)) / 100;
    fund.savingsWithdrawn = (cents_(fund.savingsWithdrawn) + cents_(request.amount)) / 100;
  });
}

function validateFundTotals_(state) {
  const earned = new Map();
  state.requests.filter(request => request.status === "approved").forEach(request => {
    earned.set(request.memberId, (earned.get(request.memberId) || 0) + cents_(request.price));
  });
  state.funds.forEach(fund => {
    validateFundShape_(fund);
    const debtPaid = cents_(fund.debtPaidTotal);
    const savingsTotal = cents_(fund.savingsTotal);
    const savingsBalance = cents_(fund.savingsBalance);
    const savingsWithdrawn = cents_(fund.savingsWithdrawn);
    if (debtPaid < 0 || savingsTotal < 0 || savingsBalance < 0 ||
        savingsWithdrawn < 0 || savingsBalance + savingsWithdrawn !== savingsTotal ||
        debtPaid + savingsTotal > (earned.get(fund.memberId) || 0)) {
      throw new Error("Финансовые суммы превышают заработок ребёнка.");
    }
  });
}

function validateRows_(currentRows, nextRows, canUpdate, updateFields, label) {
  const currentById = new Map(currentRows.map(item => [item.id, item]));
  const nextById = new Map(nextRows.map(item => [item.id, item]));
  currentRows.forEach(current => {
    const candidate = nextById.get(current.id);
    if (!candidate) throw new Error(`Нельзя удалять записи из раздела «${label}».`);
    if (!sameJson_(current, candidate)) {
      const permitted = canUpdate(current) &&
        Object.keys(current).every(key => updateFields.includes(key) || current[key] === candidate[key]) &&
        Object.keys(candidate).every(key => Object.prototype.hasOwnProperty.call(current, key));
      if (!permitted) throw new Error(`Недопустимое изменение записи в разделе «${label}».`);
    }
  });
  if (nextRows.length !== currentRows.length) {
    throw new Error(`Добавление записей в раздел «${label}» через синхронизацию запрещено.`);
  }
}

function validateTask_(task) {
  if (!task || typeof task.id !== "string" || !task.id || typeof task.title !== "string" ||
      task.title.trim().length < 1 || task.title.length > 60 || !Number.isFinite(Number(task.price)) ||
      Number(task.price) <= 0 || Number(task.price) > 100000 || typeof task.active !== "boolean" ||
      (task.visibleTo !== undefined && typeof task.visibleTo !== "string")) {
    throw new Error("Некорректные данные дела.");
  }
}

function writeState_(state) {
  replaceRows_(SHEETS.tasks, state.tasks, ["id", "title", "price", "active", "visibleTo"]);
  replaceRows_(SHEETS.requests, state.requests, SHEETS.requests.headers);
  replaceRows_(SHEETS.payments, state.payments, SHEETS.payments.headers);
  replaceRows_(SHEETS.funds, state.funds, SHEETS.funds.headers);
  replaceRows_(SHEETS.savingsRequests, state.savingsRequests, SHEETS.savingsRequests.headers);
}

function readStateData_() {
  const memberRows = readRows_(SHEETS.members);
  const storedFunds = readRows_(SHEETS.funds);
  return {
    members: memberRows.map(member => ({
      id: member.id, name: member.name, role: member.role
    })),
    tasks: readRows_(SHEETS.tasks).map(task => ({
      id: task.id, title: task.title, price: Number(task.price),
      active: task.active === true || task.active === "TRUE", visibleTo: task.visibleTo || ""
    })),
    requests: readRows_(SHEETS.requests).map(request => ({
      id: request.id, memberId: request.memberId, taskId: request.taskId,
      taskTitle: request.taskTitle, price: Number(request.price), comment: request.comment,
      createdAt: request.createdAt, status: request.status, parentComment: request.parentComment
    })),
    payments: readRows_(SHEETS.payments).map(payment => ({
      id: payment.id, memberId: payment.memberId, amount: Number(payment.amount),
      status: payment.status, createdAt: payment.createdAt, paidAt: payment.paidAt,
      confirmedAt: payment.confirmedAt
    })),
    funds: memberRows.map(member =>
      storedFunds.find(fund => fund.memberId === member.id) || emptyFund_(member.id)
    ),
    savingsRequests: readRows_(SHEETS.savingsRequests).map(request => ({
      id: request.id, memberId: request.memberId, amount: Number(request.amount),
      status: request.status, createdAt: request.createdAt, paidAt: request.paidAt,
      confirmedAt: request.confirmedAt
    }))
  };
}

function stateForMember_(state, session) {
  if (session.role === "parent") return state;
  return {
    members: state.members.filter(member => member.id === session.memberId),
    tasks: state.tasks.filter(task => !task.visibleTo || task.visibleTo === session.memberId),
    requests: state.requests.filter(request => request.memberId === session.memberId),
    payments: state.payments.filter(payment => payment.memberId === session.memberId),
    funds: state.funds.filter(fund => fund.memberId === session.memberId),
    savingsRequests: state.savingsRequests.filter(request => request.memberId === session.memberId)
  };
}

function mergeChildState_(current, childState, memberId) {
  return {
    members: current.members,
    tasks: current.tasks,
    requests: current.requests
      .filter(request => request.memberId !== memberId)
      .concat(childState.requests),
    payments: current.payments
      .filter(payment => payment.memberId !== memberId)
      .concat(childState.payments),
    funds: current.funds
      .filter(fund => fund.memberId !== memberId)
      .concat(childState.funds),
    savingsRequests: current.savingsRequests
      .filter(request => request.memberId !== memberId)
      .concat(childState.savingsRequests)
  };
}

function ensureSheet_(spreadsheet, schema) {
  let sheet = spreadsheet.getSheetByName(schema.name);
  if (!sheet) sheet = spreadsheet.insertSheet(schema.name);
  if (sheet.getLastRow() === 0) sheet.appendRow(schema.headers);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const isPrefix = headers.length <= schema.headers.length &&
    headers.every((header, index) => header === schema.headers[index]);
  if (!isPrefix) throw new Error(`Вкладка «${schema.name}» уже существует с другой структурой.`);
  if (headers.length < schema.headers.length) {
    sheet.getRange(1, headers.length + 1, 1, schema.headers.length - headers.length)
      .setValues([schema.headers.slice(headers.length)]);
  }
  return sheet;
}

function readRows_(schema) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(schema.name);
  if (!sheet) throw new Error(`Не найдена вкладка «${schema.name}». Запустите настройку сервера.`);
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, schema.headers.length).getValues();
  return values.map(row => {
    const result = {};
    schema.headers.forEach((key, index) => { result[key] = row[index]; });
    if (schema === SHEETS.members && result.role === "child") {
      const saltSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Коды детей");
      if (saltSheet && saltSheet.getLastRow() >= 2) {
        const salts = saltSheet.getRange(2, 1, saltSheet.getLastRow() - 1, 3).getValues();
        const stored = salts.find(item => item[0] === result.id);
        result.inviteHashSalt = stored ? stored[1] : "";
      }
    }
    return result;
  });
}

function appendObject_(schema, object) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(schema.name);
  if (!sheet) throw new Error(`Не найдена вкладка «${schema.name}».`);
  sheet.appendRow(schema.headers.map(key => object[key] === undefined ? "" : object[key]));
  if (schema === SHEETS.members && object.role === "child") {
    let salts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Коды детей");
    if (!salts) {
      salts = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Коды детей");
      salts.appendRow(["memberId", "salt", "createdAt"]);
    }
    salts.appendRow([object.id, object.inviteHashSalt, new Date().toISOString()]);
  }
}

function appendFund_(memberId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.funds.name);
  if (!sheet) throw new Error("Не найдена вкладка «Финансы».");
  sheet.appendRow(SHEETS.funds.headers.map(key => {
    const fund = emptyFund_(memberId);
    return fund[key] === undefined ? "" : fund[key];
  }));
}

function emptyFund_(memberId) {
  return {
    memberId: memberId,
    debtTarget: 0,
    debtRemaining: 0,
    debtPaidTotal: 0,
    debtRate: 50,
    savingsTarget: 0,
    savingsRate: 0,
    savingsBalance: 0,
    savingsTotal: 0,
    savingsWithdrawn: 0
  };
}

function replaceRows_(schema, objects, columns) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(schema.name);
  if (!sheet) throw new Error(`Не найдена вкладка «${schema.name}».`);
  const rows = objects.map(object => columns.map(key => object[key] === undefined ? "" : object[key]));
  const existingRows = Math.max(0, sheet.getLastRow() - 1);
  if (existingRows) sheet.getRange(2, 1, existingRows, columns.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, columns.length).setValues(rows);
}

function createSession_(member) {
  const token = randomToken_() + randomToken_();
  CacheService.getScriptCache().put(
    SESSION_PREFIX + digest_(token, "session"),
    JSON.stringify({ memberId: member.id, role: member.role }),
    SESSION_TTL_SECONDS
  );
  return {
    token: token,
    memberId: member.id,
    name: member.name,
    role: member.role,
    state: stateForMember_(readStateData_(), { memberId: member.id, role: member.role }),
    version: currentRevision_()
  };
}

function currentRevision_() {
  return Number(PropertiesService.getScriptProperties().getProperty(PROPERTY_KEYS.revision) || 0);
}

function bumpRevision_() {
  const properties = PropertiesService.getScriptProperties();
  const next = currentRevision_() + 1;
  properties.setProperty(PROPERTY_KEYS.revision, String(next));
  return next;
}

function requireSession_(token, expectedRole) {
  token = cleanString_(token, 120);
  if (!token) throw new Error("Сеанс истёк. Войдите ещё раз.");
  const encoded = CacheService.getScriptCache().get(SESSION_PREFIX + digest_(token, "session"));
  if (!encoded) throw new Error("Сеанс истёк. Войдите ещё раз.");
  const session = JSON.parse(encoded);
  if (expectedRole && session.role !== expectedRole) throw new Error("Недостаточно прав для этого действия.");
  const member = readRows_(SHEETS.members).find(item => item.id === session.memberId);
  if (!member || member.role !== session.role) throw new Error("Профиль больше не найден.");
  return session;
}

function registerBadPin_(properties) {
  const count = Number(properties.getProperty(PROPERTY_KEYS.failedPinCount) || 0) + 1;
  properties.setProperty(PROPERTY_KEYS.failedPinCount, String(count));
  if (count >= MAX_BAD_PIN_ATTEMPTS) {
    properties.setProperties({
      [PROPERTY_KEYS.failedPinCount]: "0",
      [PROPERTY_KEYS.blockedUntil]: String(Date.now() + PIN_LOCKOUT_MS)
    });
  }
}

function verifyParentPin_(pin) {
  const properties = PropertiesService.getScriptProperties();
  const blockedUntil = Number(properties.getProperty(PROPERTY_KEYS.blockedUntil) || 0);
  if (Date.now() < blockedUntil) throw new Error("Слишком много неверных PIN. Попробуйте позже.");
  if (!/^\d{4}$/.test(String(pin || ""))) {
    registerBadPin_(properties);
    return false;
  }
  const valid = constantTimeEquals_(
    digest_(String(pin), properties.getProperty(PROPERTY_KEYS.pinSalt)),
    properties.getProperty(PROPERTY_KEYS.pinHash)
  );
  if (!valid) {
    registerBadPin_(properties);
    return false;
  }
  properties.setProperties({
    [PROPERTY_KEYS.failedPinCount]: "0",
    [PROPERTY_KEYS.blockedUntil]: "0"
  });
  return true;
}

function digest_(value, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + ":" + String(value),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes);
}

function randomToken_() {
  return Utilities.getUuid();
}

function constantTimeEquals_(left, right) {
  left = String(left || "");
  right = String(right || "");
  let difference = left.length ^ right.length;
  const length = Math.max(left.length, right.length);
  for (let index = 0; index < length; index += 1) {
    difference |= (left.charCodeAt(index) || 0) ^ (right.charCodeAt(index) || 0);
  }
  return difference === 0;
}

function cleanString_(value, maxLength) {
  if (typeof value !== "string") return "";
  return value.trim().slice(0, maxLength);
}

function cents_(value) {
  const amount = Number(value);
  return Number.isFinite(amount) ? Math.round(amount * 100) : 0;
}

function sameJson_(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}
