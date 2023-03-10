/**
 * @OnlyCurrentDoc
 */

function otrabotka() {
  const SheetByName = 'Квартиранти'
  const SheetByName1 = 'До оплати'
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SheetByName);
  const sh1 = ss.getSheetByName(SheetByName1);
  sh1.getRange(4,2).clearContent();
  sh1.activate();

  //отдельные коммуналки
  const svet = ((raschetTab('Світло3')).toFixed(2)) * 1;
  const vodokanal = ((raschetTab('Водоканал3')).toFixed(2)) * 1;
  const gvenergo = ((raschetTab('ГВ3 Київенерго нов1')).toFixed(2)) * 1;
  const soderjanie = ((raschetTab('Утримання3')).toFixed(2)) * 1;
  const othod = ((raschetTab('Відходи3')).toFixed(2)) * 1;
  const otoplenie = ((raschetTab('Опалення3')).toFixed(2)) * 1;
  const gas = ((raschetTab('Газ3')).toFixed(2)) * 1;
  const domofon = ((raschetTab('Домофон3')).toFixed(2)) * 1;

  console.log(soderjanie);
  
  
  //сумма всех коммуналок
  const summ = svet + vodokanal + gvenergo + soderjanie + othod + otoplenie + gas + domofon;
  
  // колонка внесения суммы
  const col = numColElementa(SheetByName, 'Нараховано комуналка скрипт').numCol;

  // строка внесения суммы
  const row = numRowKvart();

  // номер колонки "Остаток коммуналка скрипт" таблицы "Квартиранты"
  const nCo = numColElementa('Квартиранти', 'Остаток комуналка скрипт').numCol;

  // внесение суммы коммуналки в 'Квартиранты' за минусом предыдущей переплаты/недоплаты
  sh.getRange(row, col).setValue((((summ * 1 + sh.getRange(row - 1, nCo).getValue()) * 1).toFixed(2)) * 1);

  // внесение в 'К оплате'
  sh1.getRange(5,2).setValue(soderjanie);
  sh1.getRange(6,2).setValue(vodokanal);
  sh1.getRange(7,2).setValue(gvenergo);
  sh1.getRange(8,2).setValue(gas);
  sh1.getRange(9,2).setValue(otoplenie);
  sh1.getRange(10,2).setValue(othod);
  sh1.getRange(11,2).setValue(svet);
  sh1.getRange(12,2).setValue(domofon);
  sh1.getRange(13,2).setValue(summ);
  sh1.getRange(14,2).setValue(sh.getRange(row - 1, nCo).getValue());
  sh1.getRange(15,2).setValue((((summ * 1 + sh.getRange(row - 1, nCo).getValue()) * 1).toFixed(2)) * 1);
  sh1.getRange(4,2).setValue(actDate());
}

// строка внесения суммы в таблицу 'Квартиранты'
function numRowKvart() {
  const SheetByName = 'Квартиранти'
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SheetByName);

  // находим номер строки с нужной датой в первой колонке
  const rowCol1 = numColElementa(SheetByName, 'Дата скрипт').numRow;

  // находим номер колонки с нужной датой в первой колонке
  const colCol1 = numColElementa(SheetByName, 'Дата скрипт').numCol;

  // берем весь диапазон данных
  const range = sh.getDataRange();

  // берем все данные с листа
  const values = range.getValues();

  // отбираем только даты с первой колонки
  const date = values.map(r => r[colCol1 - 1]).slice(rowCol1);

  // преобразовываем отобранные в мм.гггг
  const dateTrans = date.map(r => Utilities.formatDate(new Date(r), 'Europe/Kiev', 'MM.yyyy'));

  //находим номер строки нужной даты
  const dateRow = dateTrans.indexOf(actDate()) + rowCol1 + 1;
  return dateRow;
}

// поиск номера столбца и строки
function numColElementa(SheetByName, element) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SheetByName);
  const range = sh.getDataRange();
  const values = range.getValues();
  for (let i = 0; i <= values.length; i++) {
    for (let ii = 0; ii <= values[i].length; ii++) {
      if (values[i][ii] === element) {
        return {
          nameRow: 'рядок', numRow: i + 1,
          nameCol: 'стовпець', numCol: ii + 1
        };
      }
    }
  }
}

//берем цифру оплты в конкретной таблице (по конкретному виду расхода) 
function raschetTab(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  const range = sh.getDataRange();
  const values = range.getValues();
  const dateOplat = 'Дата оплати скрипт'
  const summaOplat = 'Оплата скрипт'
  // находим номер колонки 'Дата оплаты скрипт'
  const col = numColElementa(sheetName, dateOplat).numCol

  // находим номер строки 'Дата оплаты скрипт'
  const row = numColElementa(sheetName, dateOplat).numRow

  // берем одномерный массив дат из 15 колонки
  const val = ss.getSheetByName(sheetName).getRange(row + 1, col, values.length + 1).getValues().flat();

  //заменяем пустые датой 01.01.1000
  val.forEach((element, i) => { if (element == '') { val[i] = new Date('Jan 05, 1000 00:00:00') } });

  //трансформируем дату
  const date = val.map(r => Utilities.formatDate(new Date(r), 'Europe/Kiev', 'MM.yyyy'));

  //находим номер строки нужной даты
  const dateRow = date.indexOf(actDate()) + row + 1;

  //берем нужную цифру оплаты
  const oplata = sh.getRange(dateRow, numColElementa(sheetName, summaOplat).numCol).getValue();

  if (dateRow != row) {
    return oplata;

  } else {
    return 0
  }
}

// дата, на которую производится расчет
function actDate() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('До оплати');
  const range = sh.getRange(1, 1);
  const value = range.getValue();
  return Utilities.formatDate(new Date(value), 'Europe/Kiev', 'MM.yyyy');
}

function onEdit(e) {
  const range = e.range
  const value = e.value
  const ss = e.source
  const sheet = ss.getActiveSheet();
  const sheetname = e.range.getSheet().getName();

  // заполняем "Остаток коммуналка скрипт"
  if (sheetname === 'Квартиранти' && range.getColumn() === 5) {
    range.offset(0, 3).setValue((range.offset(0, 2).getValue() || 0) - (value.replace(',', '.')) * 1);
  }
  if (sheetname === 'Квартиранти' && range.getColumn() === 5 && (value == 0)) {
    range.offset(0, 3).clearContent();
    range.clearContent();
  }
}

function onOpen(){
 
  SpreadsheetApp.getUi()
  .createMenu('Меню')
  .addItem('Розрахунок', 'otrabotka')
  .addToUi();
}
