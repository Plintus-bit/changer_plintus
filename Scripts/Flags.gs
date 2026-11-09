class FlagManager {
  /**
   * @param {Array} arr
   */
  constructor(arr) {
    this.data = new Array();

    this.FillFrom(arr);
  }

  /**
   * @param {Array} arr
   */
  FillFrom(arr) {
    for (var i = 0; i < arr.length; i++) {
      const item = arr[i];

      if (item.GetValue()) {
        if (this.Find(item.constructor.name)) continue;

        this.data.push(item.constructor.name);
      }
    }
  }

  Find(flag) {
    if (this.data.length < 1) return false;

    if (this.data.indexOf(flag) > -1) {
      return true;
    }
    return false;
  }

}

class Flag {
  /**
   * @param {Boolean} value
   */
  constructor(value) {
    this.value = value;
  }

  /**
   * @return {Boolean}
   */
  GetValue() {
    return this.value;
  }
}

class FlagNeedLinkDocks extends Flag {};
class FlagNeedLinkSlides extends Flag {};
class FlagNeedTextDocks extends Flag {};
class FlagNeedTextSlides extends Flag {};

function TestFlag() {
  
  var flags = new FlagManager(
    [new FlagNeedLinkDocks(false),
    new FlagNeedLinkSlides(false),
    new FlagNeedTextDocks(true),
    new FlagNeedTextDocks(false)]
  );
  
}

