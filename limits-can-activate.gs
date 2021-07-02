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
function getExistingOffers(offersColumn, offersRow) {
  offers = getOffers(offersColumn, offersRow)

  return formatOffers(offers)
}

/**
 * @customFunction
 */
function getPublishedOffer(offerColumn, offerRow) {
  offer = getOffer(offerColumn, offerRow)

  return formatOffer(offer)
}

/**
 * @customFunction
 */
function getResult(resultColumn, resultRow) {
  var sheet = SpreadsheetApp.getActive();

  var resultRaw = sheet.getDataRange().getCell(resultRow, resultColumn).getValue().toString();
  
  if (resultRaw != 'нет' && resultRaw != 'да'&& resultRaw != 'Да') {
    throw "error 3: " + resultRaw;
  }

  var result = 'true';
  if (resultRaw == 'нет')  {
    result = 'false'
  }

  var text =  
    tab4 + "'Результат' => " + result + ",\n";

  return text;
}

/**
 * @customFunction
 */
function getLimits(limitsColumn, limitsRow) {
  var sheet = SpreadsheetApp.getActive();

  var limitsCellValue = sheet.getDataRange().getCell(limitsRow, limitsColumn).getValue().toString();
  var limitsRaw = limitsCellValue.toString().split('\n');

  var limits = {
    "arenda": 0,
    "homes": 0,
  };
  for(i = 0; i <= limitsRaw.length - 1; i++) {
    var columnValues = limitsRaw[i].toString().split(' ')

    if (columnValues.length != 4 && columnValues.length != 5) {
      throw "error 1: " + columnValues.join();
    }

    var limitValue = columnValues[columnValues.length - 1];

    if (columnValues.length == 4) {
      limits.arenda = limitValue
    } else {
      limits.homes = limitValue
    }
  }

  var text =  
    tab4 + "'Предусловие - платные лимиты' => [\n" +
    tab4 + tab + "'zone1' => ['homes' => " + limits.homes + ", 'arenda' => " + limits.arenda + "],,\n" +
    tab4 + "],\n";

  return text;
}

function getOffer(offerColumn, offerRow) {
  var sheet = SpreadsheetApp.getActive();

  var offerCellValue = sheet.getDataRange().getCell(offerRow, offerColumn).getValue();
  var offerRaw = offerCellValue.toString().split('\n');

  if (offerRaw.length != 2) {
    throw "error 1: " + offerRaw.join();
  }

  var offer = {
    'zone': '',
    'dealType': '',
    'rubric': '',
  };

  if (offerRaw[0] == 'Зона 1') {
    offer.zone = 'zone1';
  }

  if (offerRaw[0] == 'Зона 2') {
    offer.zone = 'zone2';
  }

  if (offer.zone == '') {
    throw "error 2: " + offerRaw.join();
  }

  var columnValues = offerRaw[1].split(' ');

  if (columnValues.length != 2) {
    throw "error 1: " + columnValues.length;
  }

  if (columnValues[0] != 'Аренда' && columnValues[0] != 'Продажа') {
    throw "error 2: " + columnValues[0];
  }

  offer.dealType = 'sell'
  if (columnValues[0] == 'Аренда') {
    offer.dealType = 'rent_available';
  }

  offer.rubric = mapRubric(columnValues[1])

  if (offer.rubric === 'homes' && offer.dealType == 'rent_available') {
    offer.rubric = 'arenda'
  }

  return offer
}

function formatOffer(offer) {
  var text =  
    tab4 + "'Публикуемое объявление' => [\n" +
    tab4 + tab + "'zone' => '" + offer.zone + "',\n" +
    tab4 + tab + "'rubric' => '" + offer.rubric + "',\n" +
    tab4 + tab + "'deal_type' => '" + offer.dealType + "',\n" +
    tab4 + "],\n";

  return text;
}

function getOffers(offersColumn, offersRow) {
  var sheet = SpreadsheetApp.getActive();

  var offersCellValue = sheet.getDataRange().getCell(offersRow, offersColumn).getValue();
  var offersRaw = offersCellValue.toString().split('\n');

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
  for(i = 0; i <= offersRaw.length - 1; i++){ 
    var columnValue = offersRaw[i]

    if (columnValue == 'нет') {
      break;
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
  }

  return offers
}

function formatOffers(offers) {
  var formatedZone1 = formatZoneOffers(offers['zone1']);
  var formatedZone2 = formatZoneOffers(offers['zone2']);

  var text =  
    tab4 + "'Предусловие - существующие объявления' => [\n" +
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
