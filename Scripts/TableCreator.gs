function TestTable() {
  GetTable("1v68EaaYj2BJmTiE6o2tkUrW2FiayUo01");
}

function GetTable(ID = "") {
  const t_generator = new TableGenerator();

  // сюда надо ID папки, внутри которой скрипт будет собирать ВСЕ файлы
  if (ID == "") {
    t_generator.FillByFiles("");
  }
  else {
    t_generator.FillByFiles(ID);
  }
  return t_generator.work_table.getId();
}

function GetGenerator(ID = "", fileHolder="") {
  const t_generator = new TableGenerator(fileHolder);

  // сюда надо ID папки, внутри которой скрипт будет собирать ВСЕ файлы
  if (ID == "") {
    t_generator.FillByFiles("");
  }
  else {
    t_generator.FillByFiles(ID);
  }
  return t_generator;
}

function GetTableFrom(generator) {
  return generator.work_table.getId();
}

// - - - - - - - - - - - - - 

var ROOT = "TABLES_files-data"

var FileTypesForSpreadsheets = {
  "DOC": "Д",
  "PRES": "ПР"
}

function GetBaseFileName(folderID) {
  return folderID + "||" + DriveApp.getFolderById(folderID).getName() + "_" + "CHANGER SCRIPT";
}

class TableGenerator {

  constructor(fileHolder=null) {
    if (fileHolder) {
      this.root_fold = DriveApp.getFolderById(fileHolder);
    }
    else {
      this.root_fold = DriveApp.getRootFolder();
      
      const required_fold = this.root_fold.getFoldersByName(ROOT);
      if (required_fold.hasNext()) {
        this.root_fold = required_fold.next();
      }
      else {
        this.root_fold = this.root_fold.createFolder(ROOT);
      }
    }

    BreakPoints.Reload();
  }

  Build() {
    const work_table = SpreadsheetApp.create("default");

    Logger.log("TABLE ID: " + work_table.getId());
    Logger.log("TABLE URL: " + work_table.getUrl());

    DriveApp.getFileById(work_table.getId()).moveTo(this.root_fold);

    this.work_table = work_table;

    return this.InitTable(work_table);
  }

  /**
   * @param {SpreadsheetApp.Spreadsheet} work_table
   */
  InitTable(work_table) {
    work_table.insertSheet();
    this.main_sheet = work_table.getSheets()[0];

    this.main_sheet.appendRow([
      "ID", "URL", "NAME", "PATH", "LAST UPDATED", "TYPE", "NEED REVIEW", "STATUS"
    ]);

    this.data_sheet = work_table.getSheets()[1];
    this.data_sheet.appendRow(
      ["ITEM TYPE", "ITEM LINK", "ITEM NAME", "LINK", "WHERE FOUND", "COMMENT"]
    );

    this.InitRules();

    return this;
  }

  InitRules() {
    this.InitTitleRow();

    let status_r = this.InitStatusRules();

    let all_rules = this.main_sheet.getConditionalFormatRules();
    status_r.forEach(r => {
      all_rules.push(r);
    });

    let newDataRules = this.MakeDataSheetRules();

    let dataRules = this.data_sheet.getConditionalFormatRules();
    newDataRules.forEach(r => {
      dataRules.push(r);
    });

    this.main_sheet.setConditionalFormatRules(all_rules);
    this.data_sheet.setConditionalFormatRules(dataRules);
  }

  MakeDataSheetRules() {
    const typeRange = this.data_sheet.getRange("A:A");

    let rules = new Array();
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(FileTypesForSpreadsheets.DOC)
      .setBackground("#cfe2f3")
      .setFontColor("#296aa5")
      .setRanges([typeRange])
      .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(FileTypesForSpreadsheets.PRES)
      .setBackground("#d2f1cf")
      .setFontColor("#38761d")
      .setRanges([typeRange])
      .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("TRUE")
      .setBackground("#a4c2f4")
      .setRanges([this.data_sheet.getRange(`${GetEngLetterFromNum(TABLE_COLS.CHECKBOX)}:${GetEngLetterFromNum(TABLE_COLS.CHECKBOX)}`)])
      .build()
    );

    return rules;
  }

  InitTitleRow() {
    this.main_sheet.setFrozenRows(1);

    const range = this.main_sheet.getRange("A1:" + GetEngLetterFromNum(TABLE_COLS.STATUS) + "1");
    range.setBackground("#4285f4").setFontColor("#ffffff").setFontSize(12).setFontWeight("bold");

    this.data_sheet.getRange("A1:" + GetEngLetterFromNum(TABLE_COLS.CHECKBOX) + "1").setBackground("#4285f4").setFontColor("#ffffff").setFontSize(12).setFontWeight("bold");
    this.data_sheet.setFrozenRows(1);
  }

  InitStatusRules() {
    const rules_range = this.main_sheet.getRange(GetEngLetterFromNum(TABLE_COLS.STATUS) + ":" + GetEngLetterFromNum(TABLE_COLS.STATUS));
    
    let rules = new Array();
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(FILE_STATUS.NOT_STARTED)
      .setFontColor("#b7b7b7")
      .setRanges([rules_range])
      .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(FILE_STATUS.IN_PROGRESS)
      .setBackground("#fff2cc")
      .setFontColor("#bf9000")
      .setRanges([rules_range])
      .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(FILE_STATUS.DONE)
      .setBackground("#CDFF92")
      .setFontColor("#309A1E")
      .setRanges([rules_range])
      .build()
    );

    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(FILE_STATUS.SKIPPED)
      .setBackground("#efefef")
      .setFontColor("#cccccc")
      .setRanges([rules_range])
      .build()
    );

    return rules;
  }

  FillByFiles(folderID) {
    try {
      const name = GetBaseFileName(folderID);

      const saveFiles = DriveApp.getFilesByName(name);
      var needCountData = false;
      if (saveFiles.hasNext()) {
        const table = saveFiles.next();
        this.work_table = SpreadsheetApp.openById(table.getId());

        needCountData = true;
      } else {
        this.Build();
        DriveApp.getFileById(this.work_table.getId()).setName(GetBaseFileName(folderID));
      }

      const start_folder = DriveApp.getFolderById(folderID);

      this.main_sheet = this.work_table.getSheets()[0];
      this.data_sheet = this.work_table.getSheets()[1];
      if (this.CheckFilled()) return;
      
      if (needCountData) {
        this.counter = this.main_sheet.getDataRange().getNumRows() - 1;
      }
      else {
        this.counter = 0;
      }
      this.t_counter = 0;
      
      this.mainSheetMaxRows = this.main_sheet.getMaxRows();
      this.RecurFill(start_folder, "");
      const lastLetter = GetEngLetterFromNum(TABLE_COLS.STATUS + 1);
      this.main_sheet.getRange(lastLetter + "1").getCell(1, 1).setValue("filled");
      SpreadsheetApp.flush();
    }
    catch (e) {
      throw new Error("(TABLES) Ошибка выполнения: " + e.stack);
    }
  }

  OnEndWorkWithTable() {
    const strokeToDelData = this.data_sheet.getMaxRows() - this.data_sheet.getDataRange().getNumRows();
    if (strokeToDelData > 0){
      this.data_sheet.deleteRows(this.data_sheet.getDataRange().getNumRows() + 1, strokeToDelData);
    }
    
    const strokeToDelMain = this.main_sheet.getMaxRows() - this.main_sheet.getDataRange().getNumRows();
    if (strokeToDelMain > 0) {
      this.main_sheet.deleteRows(this.main_sheet.getDataRange().getNumRows() + 1, strokeToDelMain);
    }

    this.data_sheet.getRange(`${GetEngLetterFromNum(TABLE_COLS.CHECKBOX)}2:${GetEngLetterFromNum(TABLE_COLS.CHECKBOX)}`)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    
    this.data_sheet.getRange(`A2:${GetEngLetterFromNum(TABLE_COLS.CHECKBOX)}`).sort(TABLE_COLS.LINK);
    this.OnEndFormatRules();
    
    const lastLetter = GetEngLetterFromNum(TABLE_COLS.STATUS + 2);
    this.main_sheet.getRange(lastLetter + "1").getCell(1, 1).setValue("done");
    SpreadsheetApp.flush();
  }

  OnEndFormatRules() {
    // писать формулы на англ, иначе некорретно работают (не работают)
    const currRules = this.data_sheet.getConditionalFormatRules();

    const rangeLink = this.data_sheet.getRange(`${GetEngLetterFromNum(TABLE_COLS.LINK)}:${GetEngLetterFromNum(TABLE_COLS.LINK)}`);
    
    currRules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=EQ(D1;D2)")
      .setBackground("#ffe599")
      .setRanges([rangeLink])
      .build()
    );

    currRules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=COUNTIF($D:$D;$D1)>1")
      .setBackground("#f6b26b")
      .setRanges([rangeLink])
      .build()
    );

    this.data_sheet.clearConditionalFormatRules();
    this.data_sheet.setConditionalFormatRules(currRules);
  }

  CheckFilled() {
    const lastLetter = GetEngLetterFromNum(TABLE_COLS.STATUS + 1);
    return this.main_sheet.getRange(lastLetter + "1").getCell(1, 1).getValue() == "filled";
  }
  CheckDone() {
    const lastLetter = GetEngLetterFromNum(TABLE_COLS.STATUS + 2);
    return this.main_sheet.getRange(lastLetter + "1").getCell(1, 1).getValue() == "done";
  }

  /**
   * @param {DriveApp.Folder} folder
   */
  RecurFill(folder, path) {
    const new_path = path + "/" + folder.getName();

    const files = folder.getFiles();
    if (files) {
      while (files.hasNext()) {
        
        // проверка на остановку пользователем
        if (BreakPoints.IsStoppedByUser()) throw new Error("stopped");

        const t_file = files.next();

        if (this.t_counter >= this.counter) {
          this.main_sheet.appendRow([
            t_file.getId(), t_file.getUrl(),
            t_file.getName(), new_path,
            t_file.getLastUpdated(), t_file.getMimeType(), 
            REVIEW_STATUS.UNKNOWN, FILE_STATUS.NOT_STARTED
          ]);

          if (this.mainSheetMaxRows - this.t_counter - 1 < 10) {
            this.main_sheet.insertRowsAfter(this.mainSheetMaxRows, 1000);
            this.mainSheetMaxRows += 1000;
          }
        }
        this.t_counter += 1;
      }
    }

    const folds = folder.getFolders();
    if (folds) {
      while (folds.hasNext()) {
        // проверка на остановку пользователем
        if (BreakPoints.IsStoppedByUser()) throw new Error("stopped");
        
        this.RecurFill(folds.next(), new_path);
      }
    }
  }
}