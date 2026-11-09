
function TestMulti() {
  const data = '1KT1AOC8F_pNkq7XigaJ_rdAUphGxL4Hb';

  MultiChange(data);
}

/** 
 * @param {String} data
 */
function MultiChange(data) {
  if (data == "") throw new Error("Пустые данные");
  
  const fileHolder = "1vgVKOhhvZjy3KVPuLF7ik5ja8BnCw0v2";
  const dataLinks = data.split(";");

  for(var i = dataLinks.length - 1; i > -1; i--) {
    Change(GetIdFromUrl(dataLinks[i]), fileHolder);
    dataLinks.pop();
  }
}

function Change(id = null, fileHolder = null) {
  var changer = new ChangePhrase();

  if (id && fileHolder) {
    changer.SetTableGenerator(GetGenerator(id, fileHolder));
  }
  else {
    changer.SetTableGenerator(GetGenerator("", "1vgVKOhhvZjy3KVPuLF7ik5ja8BnCw0v2"));
  }
  const copyTableID = "";

  // ВАЖНО ПРОВЕРЯТЬ ID!!!
  // 1 - ID таблицы
  // 2 - Массив массивов для замены ССЫЛОК. 1 массив представляет пару: ["что_меняем", "НА_что_меняем"]
  // 3 - Массив массивов для замены ТЕКСТОВ (пока только презентации, документы не совсем). 1 массив представляет пару: ["что_меняем", "НА_что_меняем"]
  changer.StartChangeBySpreadsheet(
    GetTableFrom(changer.generator),
    Features.EMPTY(),
    [],
    copyTableID,
    [new FlagNeedLinkDocks(true), new FlagNeedLinkSlides(true), new FlagNeedTextDocks(false), new FlagNeedTextSlides(false)]
    );
}

/**
 * @param {String} url
 */
function GetIdFromUrl(url) {
  return url.replace("https://drive.google.com/drive/folders/", "").split("?")[0];
}

class ChangePhrase {
  constructor() {
    this.generator = null;
  }

  SetTableGenerator(generator) {
    this.generator = generator;
  }

  OnEnd() {
    this.generator.OnEndWorkWithTable();
  }
  /**
   * @description Изменения происходят по данным из таблицы. 
   * Таблицу можно почистить от лишних строк вручную ДО запуска. Пустых строк между данными НЕ ДОЛЖНО БЫТЬ!
   * @param {Array} flags - 
   */
  StartChangeBySpreadsheet(ID, replacementsURL, replacementsTexts, copyTableID = "", flags = []) {
    
    BreakPoints.Reload();

    try {
      if (copyTableID != "") {
        this.copySheetRangeFinder = SpreadsheetApp.openById(copyTableID).getSheets()[0].getRange("C:C");
        this.copySheetRangeGetter = SpreadsheetApp.openById(copyTableID).getSheets()[0].getRange("F:F");
      }
      else {
        this.copySheetRangeFinder = null;
        this.copySheetRangeGetter = null;
      }

      this.flagManager = new FlagManager(flags);

      const sheets = SpreadsheetApp.openById(ID).getSheets();

      this.main_sheet = sheets[0];

      // если с таблицей закончили запуском ранее
      if (this.generator.CheckDone()) return;

      this.data_sheet = sheets[1];

      this.dataWrited = this.data_sheet.getDataRange().getNumRows() - 1;
      this.dataViewed = 0;
      
      this.replacementsURL = replacementsURL;
      this.replacementsTexts = replacementsTexts;

      this.StartBySpreadsheet();
      this.OnEnd();
    }
    catch (e) {
      throw new Error("(CHANGER) Ошибка выполнения: " + e.stack);
    }
  }

  StartBySpreadsheet() {
    const rows_count = this.main_sheet.getDataRange().getNumRows();
    const startIndex = this.GetAndFixInProgressValue();

    var needFixRows = true;

    if (startIndex < 0) return;

    for (let i = startIndex; i <= rows_count; ++i) {
      // проверка на остановку пользователем
      if (BreakPoints.IsStoppedByUser()) throw new Error("stopped");

      const t_range = this.main_sheet.getRange("A" + i + ":" + GetEngLetterFromNum(TABLE_COLS.STATUS) + i);
      const t_status_cell = t_range.getCell(1, TABLE_COLS.STATUS);

      let skipped = true;

      if (!((t_status_cell.getValue() == FILE_STATUS.DONE) || (t_status_cell.getValue() == FILE_STATUS.SKIPPED))) {
        
        if (needFixRows) {
          const findings = this.main_sheet.getRange(`${GetEngLetterFromNum(TABLE_COLS.STATUS)}:${GetEngLetterFromNum(TABLE_COLS.STATUS)}`).createTextFinder(FILE_STATUS.DONE).findAll();
          if (findings && findings.length >= 7) {
            this.CheckAndFixDataRows(findings.length);
            needFixRows = false;
          }
        }
        
        t_status_cell.setValue(FILE_STATUS.IN_PROGRESS);

        const t_id = t_range.getCell(1, TABLE_COLS.ID).getValue();
        const t_type = t_range.getCell(1, TABLE_COLS.TYPE).getValue();

        if (t_type == Types.DOC()) {
          skipped = false;

          if (this.flagManager.Find(FlagNeedLinkDocks.name)) this.ChangeDocLinks(t_id);
          if (this.flagManager.Find(FlagNeedTextDocks.name)) this.ChangeDocTexts(t_id);

          t_range.getCell(1, TABLE_COLS.NEED_REVIEW).setValue(REVIEW_STATUS.NOT);
        } 
        else if (t_type == Types.SLIDE()) {          
          skipped = false;
          if (!(this.flagManager.Find(FlagNeedLinkSlides.name) || this.flagManager.Find(FlagNeedTextSlides.name))) continue;
          
          const need_review = this.ChangeSlides(t_id);

          if (need_review) {
            t_range.getCell(1, TABLE_COLS.NEED_REVIEW).setValue(REVIEW_STATUS.NEED);
          } else {
            t_range.getCell(1, TABLE_COLS.NEED_REVIEW).setValue(REVIEW_STATUS.NOT);
          }
        }
        
        if (skipped) {
          t_status_cell.setValue(FILE_STATUS.SKIPPED);
        }
        else {
          t_status_cell.setValue(FILE_STATUS.DONE);
        }
      }
    }
    SpreadsheetApp.flush();
  }

  CheckAndFixDataRows(length) {
    const linksCount = GetValueCountFromRange(this.data_sheet.getRange(`${GetEngLetterFromNum(TABLE_COLS.ITEM_TYPE)}:${GetEngLetterFromNum(TABLE_COLS.ITEM_TYPE)}`), [
      FileTypesForSpreadsheets.DOC, FileTypesForSpreadsheets.PRES]);

    const needToCount = GetValueCountFromRange(this.main_sheet.getRange(`${GetEngLetterFromNum(TABLE_COLS.TYPE)}:${GetEngLetterFromNum(TABLE_COLS.TYPE)}`), [
      Types.DOC(), Types.SLIDE()]);

    const rowsNeeded = Math.ceil((linksCount / length)) * needToCount;
    const maxRows = this.data_sheet.getMaxRows();
    
    if (rowsNeeded > maxRows) {
      this.data_sheet.insertRowsAfter(maxRows, rowsNeeded - maxRows);
    }
  }

  GetAndFixInProgressValue() {
    const letter = GetEngLetterFromNum(TABLE_COLS.STATUS);
    const range = this.main_sheet.getRange(`${letter}:${letter}`);
    
    var index = -1;

    const finderProgress = range.createTextFinder(FILE_STATUS.IN_PROGRESS).findNext();
    if (finderProgress) {
      index = finderProgress.getRowIndex();
    }
    else {
      const finderStarted = range.createTextFinder(FILE_STATUS.NOT_STARTED).findNext();
      if (!finderStarted) return -1;
      index = finderStarted.getRowIndex() - 1;
    }

    if (range.getCell(index, 1).getValue() == FILE_STATUS.IN_PROGRESS) {
      const fixFileData = this.main_sheet.getRange(`${GetEngLetterFromNum(TABLE_COLS.URL)}${index}:${GetEngLetterFromNum(TABLE_COLS.NAME)}${index}`);
      const fileID = GetIDBySplit(fixFileData.getCell(1, 1).getValue());

      const fixRange = this.data_sheet.getRange(`${GetEngLetterFromNum(TABLE_COLS.ITEM_NAME)}:${GetEngLetterFromNum(TABLE_COLS.ITEM_NAME)}`).createTextFinder(fixFileData.getCell(1, 2).getValue()).findAll();

      var indicesToDelete = new Array();
      for(var i = 0; i < fixRange.length; ++i) {
        const tID = fixRange[i].getRowIndex();
        if (GetIDBySplit(this.data_sheet.getRange(tID, 2).getCell(1, 1).getValue()) == fileID) {
          indicesToDelete.push(tID);
        }
      }

      indicesToDelete.forEach((item) => {
        this.data_sheet.getRange(`A${item}:L${item}`).clearContent();
      });
      
      range.getCell(index, 1).setValue(FILE_STATUS.NOT_STARTED);
      return index;
    }
    return index + 1;
  }

  // если в тексте были ссылки - скажи им "Прощай, meine Liebe."
  ChangeDocTexts(fileID) {
    var doc = DocumentApp.openById(fileID);
    var docBody = doc.getBody();

    this.replacementsTexts.forEach(item => {
      docBody.replaceText(item[Replacements.OLD_VALUE()], item[Replacements.NEW_VALUE()]);
    });
    
    doc.saveAndClose();
  }

  ChangeDocLinks(fileID) {
    const doc = DocumentApp.openById(fileID);

    this.tFile = doc;

    // работа с содержимым документа
    const docBody = doc.getBody();
    const paragraphs = docBody.getParagraphs();

    paragraphs.forEach(p => {

      const childs_count = p.getNumChildren();
      
      var url_length = 0

      for (var i = 0; i < childs_count; ++i) {
        var child = p.getChild(i);

        if (child != null && child.getType() === DocumentApp.ElementType.TEXT) {
          const textEl = child.asText();
          const indices = textEl.getTextAttributeIndices();

          for (var j = 0; j < indices.length; ++j) {
            const startpos = indices[j];
            var url = textEl.getLinkUrl(startpos);

            if (j + 1 >= indices.length) {
              url_length = textEl.getText().length - 1;
            }
            else {
              url_length = indices[j + 1] - 1
            }

            if (url) {
              const new_url = this.ReplaceUrl(url);

              if (!(new_url == url)) {
                textEl.setLinkUrl(startpos, url_length, new_url);

                this.Info([FileTypesForSpreadsheets.DOC, this.tFile.getUrl(), this.tFile.getName(), url, textEl.getText().substring(startpos, url_length + 1), "Заменено на:", new_url, "TRUE"]);
              } 
              else {
                const new_url_1 = this.TryGetURLByCopyTable(url);
                if (new_url_1) {
                  textEl.setLinkUrl(startpos, url_length, new_url_1);
                  
                  this.Info([FileTypesForSpreadsheets.DOC, this.tFile.getUrl(), this.tFile.getName(), url, textEl.getText().substring(startpos, url_length + 1), "Заменено на:", new_url_1, "TRUE"]);
                }
                else {
                  this.Info([FileTypesForSpreadsheets.DOC, this.tFile.getUrl(), this.tFile.getName(), url, textEl.getText().substring(startpos, url_length + 1)]);
                }
              }

            }
          }
        }
      }
    });

    doc.saveAndClose();
  }
  
  
  /**
   * @param {String} url
   */
  TryGetURLByCopyTable(url) {
    if (!this.copySheetRangeFinder) return null;
    
    if (!url.match(".google.com/")) return null;

    const url_data = url.replace("?", "/").split(["/"])[5];
    if (url_data == null || url_data == "") return null;

    const finding = this.copySheetRangeFinder.createTextFinder(url_data).findNext();
    if (!finding) return null;

    var newID = this.copySheetRangeGetter.getCell(finding.getRowIndex(), 1).getValue();
    if (newID == null || newID == "") return null;

    const res = DriveApp.getFileById(newID);
    if (res) {
      return res.getUrl();
    }
    return DriveApp.getFolderById(newID).getUrl();
  }

  /**
   * @param {SlidesApp.TextRange} run
   */
  SlideTextOperation(run) {
    const old_text = run.asString();
    const new_text = this.ReplaceText(old_text);

    let need_review = false;

    if (!(new_text == old_text)) {
      let link = run.getTextStyle().getLink();
      if (link && link.getLinkType() === SlidesApp.LinkType.URL) {
        link = link.getUrl();
      }
      const text_col = run.getTextStyle().getForegroundColor();

      run.replaceAllText(old_text, new_text);
      
      run.getTextStyle().setLinkUrl(link);
      if (text_col) {
        textStyle.setForegroundColor(text_col);
      } else {
        need_review = true;
      }
    }

    return need_review;
  }
  
  /**
   * @param {SlidesApp.TextRange} run
   */
  SlideLinkOperation(run) {
    const link = run.getLinks()[0];
    
    let need_review = false;

    if (link) {
      const textStyle = link.getTextStyle();
      if (textStyle.hasLink() && textStyle.getLink().getLinkType() === SlidesApp.LinkType.URL) {
        const url = textStyle.getLink().getUrl();

        const new_url = this.ReplaceUrl(url);
        
        if (! (new_url == url)) {
          const text_col = textStyle.getForegroundColor();

          textStyle.setLinkUrl(new_url);
          if (text_col) {
            textStyle.setForegroundColor(text_col);
          } else {
            need_review = true;
          }

          this.Info([FileTypesForSpreadsheets.PRES, this.GetSlideURL(this.tSlideCounter - 1), this.tFile.getName(), url, "сл." + this.tSlideCounter.toString() + ", (ТЕКСТ) " + run.asString(), "Заменено на:", new_url, "TRUE"]);
        }
        else {
          const text_col = textStyle.getForegroundColor();

          const new_url_1 = this.TryGetURLByCopyTable(url);

          if (new_url_1) {
            textStyle.setLinkUrl(new_url_1);
            if (text_col) {
              textStyle.setForegroundColor(text_col);
            } else {
              need_review = true;
            }
            this.Info([FileTypesForSpreadsheets.PRES, this.GetSlideURL(this.tSlideCounter - 1), this.tFile.getName(), url, "сл." + this.tSlideCounter.toString() + ", (ТЕКСТ) " + run.asString(), "Заменено на:", new_url_1, "TRUE"]);
          }
          else {
            this.Info([FileTypesForSpreadsheets.PRES, this.GetSlideURL(this.tSlideCounter - 1), this.tFile.getName(), url, "сл." + this.tSlideCounter.toString() + ", (ТЕКСТ) " + run.asString()]);
          }
          
        }
      }
    }

    return need_review;
  }

  /**
   * @param {Array} data
   */
  Info(data) {
    // this.dataViewed += 1;
    // if (this.dataWrited > this.dataViewed) return;

    this.data_sheet.appendRow(data);
  }

  GetSlideURL(index) {
    const ID = this.tItems[index].getObjectId();
    return `https://docs.google.com/presentation/d/${this.tFile.getId()}/edit?slide=id.${ID}`;
  }

  ChangeSlides(fileID) {
    let need_review = false;

    const pres = SlidesApp.openById(fileID);
    
    this.tFile = pres;
    this.tItems = pres.getSlides();
    this.tSlideCounter = 0;

    const slides = pres.getSlides();

    slides.forEach(slide => {
      this.tSlideCounter += 1;
      slide.getPageElements().forEach(el => {
        const el_type = el.getPageElementType();
        switch (el_type) {
          case SlidesApp.PageElementType.SHAPE:
            const lnk = el.asShape().getLink();

            if (lnk && lnk.getLinkType() === SlidesApp.LinkType.URL) {
              const url_shape = lnk.getUrl();

              const new_url_shape = this.ReplaceUrl(url_shape);
              if (! (new_url_shape == url_shape)) {
                el.asShape().setLinkUrl(url_shape);
                
                this.Info([FileTypesForSpreadsheets.PRES, this.GetSlideURL(this.tSlideCounter - 1), this.tFile.getName(), url_shape, "сл." + this.tSlideCounter.toString() + ", (ФОРМА) " + el.asShape().getText().asString(), "Заменено на:", url_shape_1, "TRUE"]);
              }
              else if (! (url_shape == null || url_shape == "")) {
                const new_url_shape_1 = this.TryGetURLByCopyTable(url_shape);
                if (new_url_shape_1) {
                  el.asShape().setLinkUrl(new_url_shape_1);
                  
                  this.Info([FileTypesForSpreadsheets.PRES, this.GetSlideURL(this.tSlideCounter - 1), this.tFile.getName(), url_shape, "сл." + this.tSlideCounter.toString() + ", (ФОРМА) " + el.asShape().getText().asString(), "Заменено на:", new_url_shape_1, "TRUE"]);
                }
                else {
                  this.Info([FileTypesForSpreadsheets.PRES, this.GetSlideURL(this.tSlideCounter - 1), this.tFile.getName(), url_shape, "сл." + this.tSlideCounter.toString() + ", (ФОРМА) " + el.asShape().getText().asString()]);
                }
              }
            }

            el.asShape().getText().getParagraphs().forEach(p => {
              if (p.getRange() || p.getRange().getRuns()) {
                const runs = p.getRange().getRuns();
                runs.forEach(run => {
                  var nr_1 = false;
                  if (this.flagManager.Find(FlagNeedLinkSlides.name)) nr_2 = this.SlideLinkOperation(run);
                  var nr_2 = false;
                  if (this.flagManager.Find(FlagNeedTextSlides.name)) nr_2 = this.SlideTextOperation(run);
                  
                  need_review = need_review || nr_1 || nr_2;
                });
              }
            });

            break;
          default:
            // nothing
        }
      })
    })
    
    pres.saveAndClose();

    return need_review;
  }

  ReplaceUrl(url) {
    // обработка ссылок
    var new_url = url;

    this.replacementsURL.forEach(pair => {
      new_url = new_url.replace(
        pair[Replacements.OLD_VALUE()],
        pair[Replacements.NEW_VALUE()]);
    });
    
    return new_url;
  }

  ReplaceText(text) {
    // обработка ссылок
    var new_text = text;

    this.replacementsTexts.forEach(pair => {
      new_text = new_text.replace(
        pair[Replacements.OLD_VALUE()],
        pair[Replacements.NEW_VALUE()]);
    });

    return new_text;
  }

  /**
   * @param {String} filetype
   * @param {Array} types
   */
  InTypes(filetype, types) {
    types.forEach(type => {
      if (filetype == type) {
        return true;
      }
    })

    return false;
  }
}



