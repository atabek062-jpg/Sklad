// ============================================
// МАГАЗИН — автоматизация заказов
// ============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Магазин")
    .addItem("➕ Новый заказ", "newOrder")
    .addItem("✅ Принять заказ", "markAccepted")
    .addItem("📊 Отчёт по агенту", "agentReport")
    .addItem("🆕 Добавить товар", "addProduct")
    .addToUi();
}

// ---------- Добавить товар в каталог ----------
function addProduct() {
  const ui = SpreadsheetApp.getUi();
  const name = ui.prompt("Новый товар", "Название товара:", ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() != ui.Button.OK) return;
  const unit = ui.prompt("Единица измерения", "кг / шт / полкило:", ui.ButtonSet.OK_CANCEL);
  if (unit.getSelectedButton() != ui.Button.OK) return;
  const price = ui.prompt("Цена", "Цена за единицу:", ui.ButtonSet.OK_CANCEL);
  if (price.getSelectedButton() != ui.Button.OK) return;

  SpreadsheetApp.getActive().getSheetByName("Каталог")
    .appendRow(["", name.getResponseText(), unit.getResponseText(), Number(price.getResponseText()), "да"]);
  ui.alert("Товар добавлен ✅");
}

// ---------- Новый заказ ----------
function newOrder() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const ag = ui.prompt("Новый заказ", "Номер агента:", ui.ButtonSet.OK_CANCEL);
  if (ag.getSelectedButton() != ui.Button.OK) return;

  // показываем список товаров из каталога
  const cat = ss.getSheetByName("Каталог").getDataRange().getValues();
  let list = "";
  for (let i = 1; i < cat.length; i++) {
    list += i + ") " + cat[i][1] + " — " + cat[i][3] + " (" + cat[i][2] + ")\n";
  }
  const prod = ui.prompt("Каталог товаров", "Введите № товара:\n\n" + list, ui.ButtonSet.OK_CANCEL);
  if (prod.getSelectedButton() != ui.Button.OK) return;

  const qty = ui.prompt("Количество", "Например: 5", ui.ButtonSet.OK_CANCEL);
  if (qty.getSelectedButton() != ui.Button.OK) return;

  const row = Number(prod.getResponseText());
  if (!cat[row]) { ui.alert("Нет такого товара ❌"); return; }

  const price = cat[row][3];
  const sum = price * Number(qty.getResponseText());
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");

  ss.getSheetByName("Заказы").appendRow([date, ag.getResponseText(), cat[row][1], qty.getResponseText(), price, sum, "нет"]);
  ui.alert("Заказ записан ✅\nТовар: " + cat[row][1] + "\nСумма: " + sum);
}

// ---------- Принять заказ ----------
function markAccepted() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName("Заказы");
  const data = sheet.getDataRange().getValues();

  let lastRow = -1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][6] == "нет") { lastRow = i + 1; break; }
  }
  if (lastRow == -1) { ui.alert("Нет непринятых заказов 👍"); return; }

  const r = ui.prompt("Принять заказ",
    "Строка " + lastRow + ": " + data[lastRow-1].slice(0, 6).join(" | ") +
    "\n\nНапишите «да» чтобы принять:", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() == ui.Button.OK && r.getResponseText() == "да") {
    sheet.getRange(lastRow, 7).setValue("да");
    ui.alert("Отмечено как принято ✅");
  }
}

// ---------- Отчёт по агенту ----------
function agentReport() {
  const ui = SpreadsheetApp.getUi();
  const ag = ui.prompt("Отчёт по агенту", "Введите номер агента:", ui.ButtonSet.OK_CANCEL);
  if (ag.getSelectedButton() != ui.Button.OK) return;
  const n = ag.getResponseText();

  const data = SpreadsheetApp.getActive().getSheetByName("Заказы").getDataRange().getValues();
  let total = 0, done = 0, wait = 0, lines = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) == n) {
      total += data[i][5];
      if (data[i][6] == "да") done += data[i][5]; else wait += data[i][5];
      lines.push(data[i][0] + " | " + data[i][2] + " | " + data[i][3] + " | " + data[i][5] + " | " + data[i][6]);
    }
  }
  ui.alert("АГЕНТ №" + n + "\n\nВсего: " + total + "\nПринято: " + done +
    "\nОжидает: " + wait + "\n\n" + (lines.join("\n") || "нет заказов"));
}
