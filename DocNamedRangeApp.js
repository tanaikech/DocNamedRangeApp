/**
 * GitHub  https://github.com/tanaikech/DocNamedRangeApp<br>
 * Library name
 * @type {string}
 * @const {string}
 * @readonly
 */
var appName = "DocNamedRangeApp";

/**
 * @param {Object} document Instance of active Document.
 * @return {DocNamedRangeApp}
 */
function getActiveDocument(document) {
  this.document = document;
  return this;
}

/**
 * Get all named ranges in Document.<br>
 * @return {Object} Object Named range list.
 */
function getAllNamedRanges() {
  return new DocNamedRangeApp(this.document).getAllNamedRanges();
}

/**
 * Get object of name for named range ID.<br>
 * @return {Object} Object Named range list as the key of name. The same name can be created. So, the value of IDs is an array.
 */
function getNamesToIds() {
  return new DocNamedRangeApp(this.document).getNamesToIds();
}

/**
 * Set named range by selected contents in Document.<br>
 * @param {String} name Name of named range.
 * @param {Boolean} forParagraph Default value is false. When this is true, the paragraph is set as the named range even when the part of text is selected.
 * @return {String} String Named range ID.
 */
function setNamedRangeFromSelection(name, forParagraph = false) {
  return new DocNamedRangeApp(this.document).setNamedRangeFromSelection(
    name,
    forParagraph
  );
}

/**
 * Set named range by cursor position in Document.<br>
 * @param {String} name Name of named range.
 * @return {String} String Named range ID.
 */
function setNamedRangeFromCursorPosition(name) {
  return new DocNamedRangeApp(this.document).setNamedRangeFromCursorPosition(
    name
  );
}

/**
 * Select a named range by named range ID in Document.<br>
 * @param {String} id Named range ID.
 * @return {Object} Object Object of Class NamedRange.
 */
function selectNamedRangeById(id) {
  return new DocNamedRangeApp(this.document).selectNamedRangeById(id);
}

/**
 * Move cursor position to the top of named range by ID.<br>
 * @param {String} id Named range ID.
 * @return {Object} Object Object of Class NamedRange.
 */
function moveCursorToNamedRangesById(id) {
  return new DocNamedRangeApp(this.document).moveCursorToNamedRangesById(id);
}

/**
 * Check whether the current cursor position is included in the named range using the named range ID.<br>
 * @param {String} id Named range ID.
 * @return {Boolean} Boolean When the cursor position is included in the named range, true is return. When the cursor position is not included in the named range, false is return.
 */
function checkNamedRangeOfCursorById(id) {
  return new DocNamedRangeApp(this.document).checkNamedRangeOfCursorById(id);
}

/**
 * Delete all named ranges.<br>
 * @return void
 */
function deleteAllNamedRanges() {
  return new DocNamedRangeApp(this.document).deleteAllNamedRanges();
}

/**
 * Delete named ranges by ID.<br>
 * @param {Object} ids Named range IDs as an array.
 * @return {Object} Object Succeeded named range IDs and failed named range IDs are returned.
 */
function deleteNamedRangeById(ids) {
  return new DocNamedRangeApp(this.document).deleteNamedRangeById(ids);
}
var DocNamedRangeApp;

DocNamedRangeApp = function () {
  class DocNamedRangeApp {
    constructor(doc_) {
      this.name = appName;
      if (!doc_) {
        throw new Error(
          "Please give the object of active Document. Ex. DocNamedRangeApp.getActiveDocument(DocumentApp.getActiveDocument())"
        );
      }
      this.doc = doc_;
    }

    getAllNamedRanges() {
      return this.doc.getNamedRanges().reduce((o, e) => {
        var ele, range;
        range = e.getRange();
        ele = range.getRangeElements();
        o[e.getId()] = {
          name: e.getName(),
          namedRange: e,
          range: e.getRange(),
          type: ele.map((f) => {
            return f.getElement().getType().toString();
          }),
          rangeElements: ele,
        };
        return o;
      }, {});
    }

    getNamesToIds() {
      var o;
      o = this.getAllNamedRanges();
      return Object.entries(o).reduce((o, [k, v]) => {
        o[v.name] = o[v.name] ? [...o[v.name], k] : [k];
        return o;
      }, {});
    }

    setNamedRangeFromSelection(name_, forParagraph_) {
      var ele, endE, endtEle, nr, params, range, select, startE, startEle;
      select = this.doc.getSelection();
      if (!select) {
        throw new Error(
          "Please select contents in Document and run script. Selected contents are set as the named range."
        );
      }
      ele = select.getRangeElements();
      startEle = ele[0];
      endtEle = ele.pop();
      startE = startEle.getElement();
      endE = endtEle.getElement();
      params = forParagraph_
        ? [startE, endE]
        : [
            startE,
            startEle.getStartOffset(),
            endE,
            endtEle.getEndOffsetInclusive(),
          ];
      range = this.doc
        .newRange()
        .addElementsBetween(...params)
        .build();
      nr = this.doc.addNamedRange(name_, range);
      return nr.getId();
    }

    setNamedRangeFromCursorPosition(name_) {
      var ele, nr, range;
      ele = this.doc.getCursor().getElement();
      range = this.doc.newRange().addElement(ele).build();
      nr = this.doc.addNamedRange(name_, range);
      return nr.getId();
    }

    selectNamedRangeById(id_) {
      var o;
      o = this.getAllNamedRanges();
      if (!o[id_]) {
        throw new Error(`ID of '${id_}' was not found.`);
      }
      this.doc.setSelection(o[id_].range);
      return o[id_].namedRange;
    }

    moveCursorToNamedRangesById(id_) {
      var o, pos, re;
      o = this.getAllNamedRanges();
      if (!o[id_]) {
        throw new Error(`ID of '${id_}' was not found.`);
      }
      re = o[id_].range.getRangeElements()[0];
      pos = this.doc.newPosition(re.getElement(), re.getStartOffset());
      this.doc.setCursor(pos);
      return o[id_].namedRange;
    }

    checkNamedRangeOfCursorById(id_) {
      var check,
        cursor,
        getIndex,
        indexOfCursor,
        indexOfNamedRange,
        indexOfSelect,
        name,
        namedRange,
        offset,
        select;
      getIndex = (doc, e) => {
        while (
          e.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION
        ) {
          e = e.getParent();
        }
        return this.doc.getBody().getChildIndex(e);
      };
      namedRange = this.doc.getNamedRangeById(id_);
      if (namedRange) {
        indexOfNamedRange = namedRange
          .getRange()
          .getRangeElements()
          .map((e) => {
            return {
              idx: getIndex(this.doc, e.getElement()),
              start: e.getStartOffset(),
              end: e.getEndOffsetInclusive(),
            };
          });
      } else {
        throw new Error("No namedRange.");
      }
      name = namedRange.getName();
      cursor = this.doc.getCursor();
      if (cursor) {
        indexOfCursor = getIndex(this.doc, cursor.getElement());
        offset = cursor.getOffset();
        check = indexOfNamedRange.some(({ idx, start, end }) => {
          return (
            idx === indexOfCursor &&
            ((start === -1 && end === -1) || (offset > start && offset < end))
          );
        });
        if (check) {
          return true;
        }
        return false;
      }
      select = this.doc.getSelection();
      if (select) {
        indexOfSelect = select.getRangeElements().map((e) => {
          return {
            idx: getIndex(this.doc, e.getElement()),
            start: e.getStartOffset(),
            end: e.getEndOffsetInclusive(),
          };
        });
        check = indexOfSelect.some((e) => {
          return indexOfNamedRange.some(({ idx, start, end }) => {
            return (
              idx === e.idx &&
              ((start === -1 && end === -1) ||
                (e.start > start && e.start < end) ||
                (e.end > start && e.end < end))
            );
          });
        });
        if (check) {
          return true;
        }
        return false;
      }
      throw new Error("No cursor and select.");
    }

    deleteAllNamedRanges() {
      this.doc.getNamedRanges().forEach((e) => {
        return e.remove();
      });
      return null;
    }

    deleteNamedRangeById(ids_) {
      var obj;
      if (!Array.isArray(ids_)) {
        throw new Error("Please give the named range IDs as an array.");
      }
      obj = this.getAllNamedRanges();
      return ids_.reduce(
        (o, id) => {
          if (obj[id]) {
            obj[id].namedRange.remove();
            o.success.push(id);
          } else {
            o.failure.push(id);
          }
          return o;
        },
        {
          success: [],
          failure: [],
        }
      );
    }
  }

  DocNamedRangeApp.name = appName;

  return DocNamedRangeApp;
}.call(this);
