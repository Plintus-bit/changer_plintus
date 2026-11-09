class Utils {
  static TextCOPY() {
    return "Копия ";
  }
  static CURR_YEAR() {
    return "25/26";
  }
  static GetMimeTypes() {
    var types = "aam.aiff.asf.au.avi.doc.docx.exe.gif.html.htm.jpg.jpeg.jif.mp3.mpe.mpg.mpeg.moov.movie.mov.numbers.pdf.png.ppt.pptx.pps.qt.ra.ram.rm.rtf.swf.conf.py.js.gs.tiff.tif.txt.wav.wma.wmf.wmv.wpd.xls.xlsx.zip.rar.7z.bat.md.dll.sh.dmg.psd.bmp"
    return types.split(".");
  }
}

class NameParams {
  static NAME() {
    return 0;
  }
  static TYPE() {
    return 1;
  }
}

class Types {
  static DOC() {
    return "application/vnd.google-apps.document";
  }
  static SLIDE() {
    return "application/vnd.google-apps.presentation";
  }
  static DOCX() {
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  }
  static PPTX() {
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.presentation";
  }
}

class Replacements {
  static OLD_VALUE() {
    return 0;
  }
  static NEW_VALUE() {
    return 1;
  }
}

/**
 * @param {Integer} num - position of letter in althabet. Starts from 1
 * @param {Integer} delta - if need to shift the value. Add delta to num
 */
function GetEngLetterFromNum(num, delta = 0) {
  const alf = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var fnum = num - 1 + delta;
  if (Math.floor(fnum / alf.length) > 0) {
    return alf[Math.floor(fnum / alf.length) - 1] + alf[fnum % alf.length]
  }
  return alf[fnum % alf.length];
}
/**
 * @param {String} letter - position of letter in althabet. Starts from 1
 * @param {Integer} delta - if need to shift the value. Add delta to num
 */
function GetNumFromEngLetter(letter, delta = 0) {
  const alf = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  return alf.indexOf(letter.toUpperCase()) + 1 - delta;
}

/**
 * @param {Integer} num - position of letter in althabet. Starts from 1
 */
function GetEngLetterFromNum(num) {
  const alf = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  return alf[num - 1];
}

/**
 * @param {String} url - position of letter in althabet. Starts from 1
 */
function GetIDBySplit(url) {
  const data = url.split("/");
  return data[data.length - 2];
}

/**
 * @param {SpreadsheetApp.Range} range
 * @param {Array} data
 */
function GetValueCountFromRange(range, data) {
  var result = 0;
  for (var i = 0; i < data.length; i++) {
    const val = range.createTextFinder(data[i]).findAll();
    if (val) {
      result += val.length;
    }
  }
  return result;
}

var TABLE_COLS = {
  "ID": 1,
  "URL": 2,
  "NAME": 3,
  "PARENT": 4,
  "LAST_UPDATED": 5,
  "TYPE": 6,
  "NEED_REVIEW": 7,
  "STATUS": 8,
  // далее для второго 
  "ITEM_TYPE": 1,
  "ITEM_LINK": 2,
  "ITEM_NAME": 3,
  "LINK": 4,
  "WHERE_FOUND": 5,
  "COMMENT": 6,
  "CHECKBOX": 8,
  "LINK": 4
}

var FILE_STATUS = {
  "NOT_STARTED": "not started",
  "IN_PROGRESS": "in progress",
  "DONE": "done",
  "SKIPPED": "skip"
}

var REVIEW_STATUS = {
  "NEED": "NEED",
  "NOT": "-",
  "UNKNOWN": "?"
}
