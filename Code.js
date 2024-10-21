function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Keuangan')
    .addItem('Entri Pemasukan', 'entriPemasukan')
    .addItem('Entri Pengeluaran', 'entriPengeluaran')
    .addToUi();
}

function entriPemasukan() {
  entriKeuangan('AngIN', 'EntriIN', 'Pemasukan');
}

function entriPengeluaran() {
  entriKeuangan('AngOT', 'EntriOT', 'Pengeluaran');
}

function entriKeuangan(sourceSheetName, targetSheetName, type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeCell = sheet.getActiveCell();
  var activeRow = activeCell.getRow();
  var sourceSheet = sheet.getSheetByName(sourceSheetName);

  // Get kode_program from column C of the active row in source sheet
  var kode_program = sourceSheet.getRange('C' + activeRow).getValue();

  // Create and show dialog
  var result = showDialog(type);

  if (result && result.buttonClicked == 'OK') {
    var targetSheet = sheet.getSheetByName(targetSheetName);
    var lastRow = targetSheet.getLastRow();
    var newRow = lastRow + 1;

    // Get BIDANG NAME and UNIT NAME
    var bidangName = getBidangName(kode_program.substring(0, 2));
    var unitName = getUnitName(kode_program.substring(3, 6));

    // Set values in the new row
    targetSheet.getRange(newRow, 1).setValue(new Date()); // Datetime (GMT+7)
    targetSheet.getRange(newRow, 2).setValue(kode_program); // kode_program
    targetSheet.getRange(newRow, 3).setValue(bidangName); // BIDANG NAME
    targetSheet.getRange(newRow, 4).setValue(unitName); // UNIT NAME
    targetSheet.getRange(newRow, 5).setValue(result.nominal); // Nominal
    targetSheet.getRange(newRow, 6).setValue(result.keterangan); // Keterangan

    // Format datetime to GMT+7
    var dateTimeCell = targetSheet.getRange(newRow, 1);
    dateTimeCell.setNumberFormat('dd/MM/yyyy HH:mm:ss');

    // Format nominal as number
    var nominalCell = targetSheet.getRange(newRow, 5);
    nominalCell.setNumberFormat('#,##0');

    SpreadsheetApp.getUi().alert('Entry ' + type + ' berhasil ditambahkan.');
  }
}

function showDialog(type) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Entri ' + type,
    'Masukkan nominal dan keterangan (dipisahkan dengan koma):',
    ui.ButtonSet.OK_CANCEL
  );

  var button = result.getSelectedButton();
  var text = result.getResponseText();

  if (button == ui.Button.OK) {
    var parts = text.split(',');
    if (parts.length != 2 || isNaN(parts[0].trim())) {
      ui.alert(
        'Input tidak valid. Harap masukkan nominal (angka) dan keterangan, dipisahkan dengan koma.'
      );
      return null;
    }
    return {
      buttonClicked: 'OK',
      nominal: parseFloat(parts[0].trim()),
      keterangan: parts[1].trim(),
    };
  } else {
    return { buttonClicked: 'Cancel' };
  }
}

function getBidangName(abbr) {
  var bidangMap = {
    SK: 'SEKRETARIAT',
    AG: 'KEAGAMAAN',
    SO: 'SOSIAL',
    KM: 'KEMANUSIAAN',
  };
  return bidangMap[abbr] || '';
}

function getUnitName(abbr) {
  var unitMap = {
    SEK: 'SEKRETARIAT',
    MAI: 'TEKNIK MAINTENANCE',
    MUT: 'PENJAMINAN MUTU',
    SDM: 'PERSONALIA SDM',
    KEA: 'KEBERSIHAN KEAMANAN',
    HUM: 'KEHUMASAN',
    DAN: 'USAHA DANA',
    KED: 'KEDAI',
    TKM: 'KETAKMIRAN',
    KID: 'REMASKIDZ',
    TPQ: 'TPQ',
    MUS: 'KEMUSLIMAHAN',
    LAZ: 'LAZMU',
    AMB: 'AMBULANS',
    JNZ: 'SAKITJENAZAH',
    DCR: 'DAYCARE',
    KBT: 'KBTK',
    KOL: 'KOLAM RENANG',
  };
  return unitMap[abbr] || '';
}
