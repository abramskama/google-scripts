var tab = "    ";
var tab3 = tab + tab + tab;
var tab4 = tab3 + tab;
var tab5 = tab4 + tab;
var tab6 = tab5 + tab;

/**
 * @customFunction
 */
function getCaseStart(caseNumberColumn, caseNumberRow, caseNameColumn, caseNameRowStart, caseNamRowEnd) {
  var sheet = SpreadsheetApp.getActive();

  var caseNumberCell = sheet.getDataRange().getCell(caseNumberRow, caseNumberColumn);

  var caseNameCells = [];
  for(i = caseNameRowStart; i <= caseNamRowEnd; i++){ 
    var columnValue = sheet.getDataRange().getCell(i, caseNameColumn).getValue();

    if (columnValue != '') {
        caseNameCells.push(columnValue)
    }
  }

  var text =  
    tab3 + "'" + caseNumberCell.getValue() + ". " + caseNameCells.join(' - ') + "' => [\n";

  return text;
}

/**
 * @customFunction
 */
function getCaseEnd() {
  return tab3 + "],\n";
}

/**
 * @customFunction
 */
function getLimits(homesColumn, homesRow, arendaColumn, arendaRow) {
  var sheet = SpreadsheetApp.getActive();
  var homesCell = sheet.getDataRange().getCell(homesRow, homesColumn);
  var arendaCell = sheet.getDataRange().getCell(arendaRow, arendaColumn);

  return formatLimits(getActiveLimit(homesCell), getActiveLimit(arendaCell), true) + 
    formatLimits(getExpiredLimit(homesCell), getExpiredLimit(arendaCell), false)
}

/**
 * @customFunction
 */
function getExistingLimits(homesColumn, homesRow, arendaColumn, arendaRow) {
  var sheet = SpreadsheetApp.getActive();
  var homesCell = sheet.getDataRange().getCell(homesRow, homesColumn);
  var arendaCell = sheet.getDataRange().getCell(arendaRow, arendaColumn);

  return formatLimits(getActiveLimit(homesCell), getActiveLimit(arendaCell), true)
}

/**
 * @customFunction
 */
function getExistingOffers(offersColumn, offersRowStart, offersRowEnd) {
  offers = getOffers(offersColumn, offersRowStart, offersRowEnd)

  return formatOffers(true, offers)
}

/**
 * @customFunction
 */
function getExpectedOffers(offersColumn, offersRowStart, offersRowEnd) {
  offers = getOffers(offersColumn, offersRowStart, offersRowEnd)

  return formatOffers(false, offers)
}

/**
 * @customFunction
 */
function getPublished(rubricsColumn, rubricsRow) {
  var sheet = SpreadsheetApp.getActive();

  var offers = {};
  offers['zone1'] = parsePublished(sheet, rubricsColumn, rubricsRow)
  offers['zone2'] = parsePublished(sheet, rubricsColumn + 5, rubricsRow)

  return formatPublished(offers)
}

/**
 * @customFunction
 */
function getExcess(rubricsColumn, rubricsRow) {
  var sheet = SpreadsheetApp.getActive();

  offers = {
    "zone1": {
      "homes": 0,
      "newhomes": 0,
      "arenda": 0,
      "kn": 0,
      "garages": 0,
      "other": 0,
      "cottage": 0,
      "land": 0,
    },
    "zone2": {
      "homes": 0,
      "newhomes": 0,
      "arenda": 0,
      "kn": 0,
      "garages": 0,
      "other": 0,
      "cottage": 0,
      "land": 0,
    }
  }

  for(i = rubricsRow; i <= rubricsRow + 7; i++){ 
    var rawRubric = sheet.getDataRange().getCell(i, rubricsColumn).getValue();
    var rubric = mapRubric(rawRubric)
    var zone1Value = sheet.getDataRange().getCell(i, rubricsColumn + 1).getValue();
    var zone2Value = sheet.getDataRange().getCell(i, rubricsColumn + 2).getValue();

    offers['zone1'][rubric] = parseInt(zone1Value) || 0
    offers['zone2'][rubric] = parseInt(zone2Value) || 0
  }

  var formatedZone1 = formatZoneOffersCount(offers['zone1']);
  var formatedZone2 = formatZoneOffersCount(offers['zone2']);

  var text =  
    tab4 + "'Ожидаемое состояние по превышениям' => [\n" +
    tab4 + tab + "'zone1' => [" + formatedZone1 + "],\n" +
    tab4 + tab + "'zone2' => [" + formatedZone2 + "],\n" +
    tab4 + "],\n";

  return text;
}

function formatPublished(offers) {
  var formatedZone1 = formatZoneOffers(offers['zone1']);
  var formatedZone2 = formatZoneOffers(offers['zone2']);

  var text =  
    tab4 + "'Ожидаемое состояние по опубликованным' => [\n" +
    tab4 + tab + "'zone1' => [" + formatedZone1 + "],\n" +
    tab4 + tab + "'zone2' => [" + formatedZone2 + "],\n" +
    tab4 + "],\n";

  return text;
}

function parsePublished(sheet, rubricsColumn, rubricsRow) {
  offers = {
    "homes": {"sell": 0, "rent_available": 0, "any": 0},
    "newhomes": {"sell": 0, "rent_available": 0, "any": 0},
    "arenda": {"sell": 0, "rent_available": 0, "any": 0},
    "kn": {"sell": 0, "rent_available": 0, "any": 0},
    "garages": {"sell": 0, "rent_available": 0, "any": 0},
    "other": {"sell": 0, "rent_available": 0, "any": 0},
    "cottage": {"sell": 0, "rent_available": 0, "any": 0},
    "land": {"sell": 0, "rent_available": 0, "any": 0},
  };

  for(i = rubricsRow; i <= rubricsRow + 7; i++){ 
    var rawRubric = sheet.getDataRange().getCell(i, rubricsColumn).getValue();
    var rubric = mapRubric(rawRubric)
    var sellValue = sheet.getDataRange().getCell(i, rubricsColumn + 1).getValue();
    var arendaValue = sheet.getDataRange().getCell(i, rubricsColumn + 2).getValue();
    var anyValue = sheet.getDataRange().getCell(i, rubricsColumn + 3).getValue();

    offers[rubric]['sell'] = parseInt(sellValue) || 0
    offers[rubric]['rent_available'] = parseInt(arendaValue) || 0
    offers[rubric]['any'] = parseInt(anyValue) || 0
  }

  return offers;
}

function getOffers(offersColumn, offersRowStart, offersRowEnd) {
  var sheet = SpreadsheetApp.getActive();

  var offers = { 
    "zone1":{
      "homes": {"sell": 0, "rent_available": 0},
      "newhomes": {"sell": 0, "rent_available": 0},
      "arenda": {"sell": 0, "rent_available": 0},
      "kn": {"sell": 0, "rent_available": 0},
      "garages": {"sell": 0, "rent_available": 0},
      "other": {"sell": 0, "rent_available": 0},
      "cottage": {"sell": 0, "rent_available": 0},
      "land": {"sell": 0, "rent_available": 0}
    }, 
    "zone2":{
      "homes": {"sell": 0, "rent_available": 0},
      "newhomes": {"sell": 0, "rent_available": 0},
      "arenda": {"sell": 0, "rent_available": 0},
      "kn": {"sell": 0, "rent_available": 0},
      "garages": {"sell": 0, "rent_available": 0},
      "other": {"sell": 0, "rent_available": 0},
      "cottage": {"sell": 0, "rent_available": 0},
      "land": {"sell": 0, "rent_available": 0}
    }
  };
  var zone = 'zone1';
  var splitted = false;
  for(i = offersRowStart; i <= offersRowEnd; i++){ 
    var columnValue = sheet.getDataRange().getCell(i, offersColumn).getValue();

    if (columnValue == 'Продажа/Аренда') {
      splitted = true
      continue;
    }

    if (columnValue == 'Зона 1') {
      zone = 'zone1'
      continue;
    }

    if (columnValue == 'Зона 2') {
      zone = 'zone2'
      continue;
    }

    if (columnValue == '') {
        continue;
    }

    var columnValues = columnValue.toString().split(' ')

    if (!splitted) {
      if (columnValues.length != 4) {
        throw "error 1: " + columnValues.join();
      }

      if (columnValues[0] != 'Аренда' && columnValues[0] != 'Продажа') {
        throw "error 2: " + columnValues[0];
      }

      var dealType = 'sell'
      if (columnValues[0] == 'Аренда') {
        dealType = 'rent_available';
      }

      var rubric = mapRubric(columnValues[1])

      if (rubric === 'homes' && dealType == 'rent_available') {
        rubric = 'arenda'
      }

      offers[zone][rubric][dealType] = columnValues[3]
      continue;
    }

    if (columnValues.length != 3) {
      throw "error 4";
    }

    var rubric = mapRubric(columnValues[0])

    var dealTypesValues = columnValues[2].split('/')

    if (dealTypesValues.length != 2) {
      throw "error 6";
    }

    if (rubric == 'homes') {
      offers[zone]['homes']['sell'] = dealTypesValues[0]
      offers[zone]['arenda']['rent_available'] = dealTypesValues[1]
      continue;
    }

    offers[zone][rubric]['sell'] = dealTypesValues[0]
    offers[zone][rubric]['rent_available'] = dealTypesValues[1]
  }

  return offers
}

/**
 * @customFunction
 */
function getExpectedOffersState() {
  var text =  
    tab4 + "'Ожидаемое состояние по объявлениям' => [\n" +
    tab4 + tab + "'zone1' => [],\n" +
    tab4 + tab + "'zone2' => [],\n" +
    tab4 + "],\n";

  return text;
}

function formatOffers(exists, offers) {
  var formatedZone1 = formatZoneOffers(offers['zone1']);
  var formatedZone2 = formatZoneOffers(offers['zone2']);

  var text =  
    tab4 + "'" + (exists ? 'Предусловие - существующие объявления' : 'Ожидаемое состояние по объявлениям') + "' => [\n" +
    tab4 + tab + "'zone1' => [" + formatedZone1 + "],\n" +
    tab4 + tab + "'zone2' => [" + formatedZone2 + "],\n" +
    tab4 + "],\n";

  return text;
}

function formatZoneOffers(offers) {
  var strings = [];
  for (const [rubric, dealTypes] of Object.entries(offers)) {
    if (dealTypes['rent_available'] == 0 && dealTypes['sell'] == 0 && (!dealTypes.hasOwnProperty('any') || dealTypes['any'] == 0)) {
      continue;
    }

    var string = '';

    var dealTypeStrings = [];

    if (dealTypes.hasOwnProperty('any') && dealTypes['any'] != 0) {
      dealTypeStrings.push("'any' => " + dealTypes['any']);
    }

    if (dealTypes['sell'] != 0) {
      dealTypeStrings.push("'sell' => " + dealTypes['sell']);
    } 

    if (dealTypes['rent_available'] != 0) {
      dealTypeStrings.push("'rent_available' => " + dealTypes['rent_available']);
    }

    string = tab6 + "'" + rubric + "' => [" + dealTypeStrings.join(", ") + "],";

    strings.push(string)
  }

  return strings.length == 0 ? '' : '\n' + strings.join('\n') + '\n' + tab5;
}

function formatZoneOffersCount(offers) {
  var strings = [];
  for (const [rubric, count] of Object.entries(offers)) {
    if (count === 0) {
      continue;
    }

    strings.push(tab6 + "'" + rubric + "' => " + count + ",")
  }

  return strings.length == 0 ? '' : '\n' + strings.join('\n') + '\n' + tab5;
}

function formatLimits(homesValue, arendaValue, isActive) {
  var text = 
    tab4 + "'Предусловие - " + (isActive ? 'активные' : 'протухшие') + " платные лимиты' => [\n" +
    tab4 + tab + "'zone1' => ['homes' => "+ homesValue +", 'arenda' => "+ arendaValue +"],\n" +
    tab4 + "],\n"
  ;

  return text;
}

function getActiveLimit(cell) {
  isCellRed = cell.getBackground() == '#ff0000';

  return isCellRed ? 0 : cell.getValue();
}

function getExpiredLimit(cell) {
  isCellRed = cell.getBackground() == '#ff0000';

  return isCellRed ? cell.getValue() : 0;
}

function mapRubric(rawRubric) {
  var rubric = '';
  switch (rawRubric) {
    case 'Квартиры':
      rubric = 'homes'
      break;
    case 'Новостройки':
      rubric = 'newhomes'
      break;
    case 'Аренда квартир':
      rubric = 'arenda'
      break;
    case 'Коттеджи':
    case 'Дома':
      rubric = 'cottage'
      break;
    case 'Коммерческая':
      rubric = 'kn'
      break;
    case 'Гаражи':
      rubric = 'garages'
      break;
    case 'Земля':
      rubric = 'land'
      break;
    case 'Дачи':
      rubric = 'other'
      break;
    default:
      rubric = '';
      break;
  }

  if (rubric == '') {
    throw "incorrect rubric " + rawRubric;
  }

  return rubric
}
