import XLSX from "xlsx-js-style";
export default function Sheet(width = 0, x = 0, y = 0) {
  if (!width) {
    throw new Error("Sheet width is not defined.");
  }
  var totalColumns = width;
  var totalRows = 0;

  let rows = [];
  let merges = [];

  var sheetColumnStyles = [];

  /* UTILS */
  //Empty Cell
  const getEmptyCell = () => {
    return {
      v: "",
      t: "s",
    };
  };
  //Empty Row
  const getEmptyRow = () => {
    let emptyRow = [];
    for (let i = 0; i < width; i++) {
      emptyRow.push(getEmptyCell());
    }
    return emptyRow;
  };

  /* ADD MERGES */

  const addMerge = (r, c, w, h) => {
    r = r + y;
    c = c + x;
    w = w < 1 ? 0 : w;
    h = h < 1 ? 0 : h;
    merges.push({
      s: { r: r, c: c },
      e: { r: r + (h - 1), c: c + (w - 1) },
    });
  };

  /* ADD ROW */
  this.addRow = ({ top, row, styles }) => {
    top = top && top > 0 ? top : 0;
    row = row ? row : getEmptyRow();

    var emptyRows = [];
    //Empty Rows
    for (let i = 0; i < top; i++) {
      emptyRows.push(getEmptyRow());
      totalRows++;
    }
    //Current Row
    var currentRow = [];
    var columnPosition = 0;
    row.forEach((currentCell) => {
      let { left, width, height, ...rest } = currentCell;
      //validating
      left = left && left > 0 ? left : 0;
      width = width && width > 1 ? width : 1;
      height = height && height > 1 ? height : 1;

      addMerge(totalRows, columnPosition + left, width, height);

      for (let i = 0; i < left; i++) {
        currentRow.push(getEmptyCell());
        columnPosition++;
      }
      currentRow.push({ ...rest });
      columnPosition++;
      for (let i = 1; i < width; i++) {
        currentRow.push(getEmptyCell());
        columnPosition++;
      }
    });
    /* Filling REST CELLS OF CURRENT ROW */
    for (let i = columnPosition; i < totalColumns; i++) {
      currentRow.push(getEmptyCell());
    }
    /* GATHERING EMPTYROWS & CURRENTROW  AND PUSHING INTO ROWS*/
    const createdRows = [...emptyRows, currentRow];
    totalRows++;
    if (columnPosition > totalColumns) {
      throw new Error(
        `Row no. ${totalRows} have width is greater than sheet-width.`
      );
    }
    createdRows.forEach((val) => {
      rows.push(val);
    });

    if (styles) {
      addStyles(totalRows, styles);
    }
  };

  /* ADD ROWS */
  this.addRows = (rows) => {
    rows.forEach((val) => {
      this.addRow(val);
    });
  };

  /* ADD STYLES */
  function addStyles(rowNumber, styles = []) {
    const rowIndex = rowNumber - 1;
    styles.forEach((val) => {
      if (val.s > totalColumns) {
        throw new Error(
          `value of 's' is greater than sheet-width in row no. ${rowNumber}.`
        );
      }
      if (val.s > val.e) {
        throw new Error(
          `value of 'e' cannot be equal or less than value of 's' on row no. ${rowNumber}.`
        );
      }
      if (val.s < 1) {
        throw new Error(
          `value of 's' cannot be 0 or less on row ${rowNumber}.`
        );
      }
      const startCellIndex = val.s - 1;
      const endCellIndex = val.e - 1;
      const styleToBeAdded = val.style;
      for (let i = startCellIndex; i <= endCellIndex; i++) {
        const currentStyle = rows[rowIndex][i].s || {};

        rows[rowIndex][i].s = mergeObjects(currentStyle, styleToBeAdded);
      }
    });
  }
  /* ADD COLUMNSTYLE */
  this.addColumnStyle = (columnNumber, startRow, endRow, style) => {
    if (columnNumber < 1) {
      throw new Error("'columnNumber' cannot be less than sheet-width.");
    }
    if (columnNumber > totalColumns) {
      throw new Error("'columnNumber' cannot be less than 1.");
    }
    if (startRow > totalRows || endRow > totalRows) {
      throw new Error(
        `'startRow' or endRow cannot greater than totalRows(${totalRows}) of sheet.`
      );
    }
    if (startRow > endRow) {
      throw new Error("'endRow' should be greater than or equal to 'endRow'.");
    }
    if (startRow < 1 || endRow < 1) {
      throw new Error("'startRow' or 'endRow' should be greater than zero.");
    }
    const columnIndex = columnNumber - 1;
    const startRowIndex = startRow - 1;
    const endRowIndex = endRow - 1;
    for (let i = startRowIndex; i <= endRowIndex; i++) {
      const currentStyle = rows[i][columnIndex].s || {};
      rows[i][columnIndex].s = mergeObjects(currentStyle, style);
    }
  };
  /* MERGE STYLE */
  function mergeObjects(obj /*, …*/) {
    for (var i = 1; i < arguments.length; i++) {
      for (var prop in arguments[i]) {
        var val = arguments[i][prop];
        if (!obj[prop]) {
          obj[prop] = val;
        } else {
          if (typeof val == "object") {
            mergeObjects(obj[prop], val);
          } else {
            obj[prop] = val;
          }
        }
      }
    }
    return JSON.parse(JSON.stringify(obj));
  }

  /* GET ROWS AND COLUMNSTYLES */
  Object.defineProperty(this, "columnStyles", {
    set: (value) => {
      sheetColumnStyles = value;
    },
  });

  /* ADD SHEET COLUMNSTYLE */
  const addColumnStyles = () => {
    sheetColumnStyles.forEach(({ cn, styles }) => {
      cn.forEach((cNum) => {
        styles.forEach(({ s, e, style }) => {
          if (!s) throw new Error("Value of 's' is not defined.");
          if (!e) throw new Error("Value of 'e' is not defined.");
          s = s < 1 ? 1 : s;
          e = e > totalRows ? totalRows : e;
          if (s > e) throw new Error("'s' cannot be greater than 'e'.");
          for (let i = s - 1; i < e; i++) {
            this.addColumnStyle(cNum, s, e, style);
          }
        });
      });
    });
  };
  /* SAVE */
  this.save = function (name = "sheetstyle-demo.xlsx") {
    addColumnStyles();
    let yOffsetRows = [];
    let xOffsetRows = [];
    for (let i = 0; i < y; i++) {
      yOffsetRows.push([getEmptyCell()]);
    }
    let xOffsetCells = [];
    for (let i = 0; i < x; i++) {
      xOffsetCells.push(getEmptyCell());
    }
    for (let i = 0; i < totalRows; i++) {
      xOffsetRows.push([...xOffsetCells, ...rows[i]]);
    }

    let arr = [];
    arr = [...yOffsetRows, ...xOffsetRows];
    const ws = XLSX.utils.aoa_to_sheet(arr);
    ws["!merges"] = merges;
    /* WORKBOOK */
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "readme demo");

    /* SAVE WORKBOOK */
    XLSX.writeFile(wb, name);
  };
}
