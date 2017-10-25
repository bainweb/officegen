module.exports = {

  // assume passed in an array of row objects
  getTable: function(rows, tblOpts) {
    tblOpts = tblOpts || {};

    var self = this;

    return self._getBase(
      rows.map(function(row) {
        return self._getRow(
          row.map(function(cell) {
            cell = cell || {};
            if (typeof cell === 'string' || typeof cell === 'number') {
              var val = cell;
              cell = {
                val: val
              };
            }

            return self._getCell(cell.val, cell.opts, tblOpts);
          }),
          tblOpts
        );
      }),
      self._getColSpecs(rows, tblOpts),
      tblOpts
    );
  },

  _getBase: function(rowSpecs, colSpecs, opts) {
    var self = this;

    var baseTable = {
      "w:tbl": {
        "w:tblPr": {
          "w:tblStyle": {
            "@w:val": "a3"
          },
          "w:tblW": {
            "@w:w": "0",
            "@w:type": "auto"
          },
          "w:tblLook": {
            "@w:val": "04A0",
            "@w:firstRow": "1",
            "@w:lastRow": "0",
            "@w:firstColumn": "1",
            "@w:lastColumn": "0",
            "@w:noHBand": "0",
            "@w:noVBand": "1"
          }
        },
        "w:tblGrid": {
          "#list": colSpecs
        },
        "#list": [rowSpecs]
      }
    };
    if (opts.borders) {
      baseTable["w:tbl"]["w:tblPr"]["w:tblBorders"] = {
        "w:top": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        },
        "w:bottom": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        },
        "w:left": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        },
        "w:right": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        },
        "w:insideH": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        },
        "w:insideV": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        }
      };
    }
    if (opts.rtl) {
      baseTable["w:tbl"]["w:tblPr"]["w:bidiVisual"] = {};
    }
    if (opts.align) {
      baseTable["w:tbl"]["w:tblPr"]["w:jc"] = {
        "@w:val": opts.align || "left"
      };
    }
    if (opts.indent) {
      baseTable["w:tbl"]["w:tblPr"]["w:tblInd"] = {
        "@w:w": opts.indent || 0,
        "@w:type": "dxa"
      };
    }
    return baseTable;
  },

  _getColSpecs: function(cols, opts) {
    var self = this;
    return cols[0].map(function(val, idx) {
      return self._tblGrid(opts);
    });
  },

  // TODO
  _tblGrid: function(opts) {
    return {
      "w:gridCol": {
        "@w:w": opts.tableColWidth || "1"
      }
    };
  },


  _getRow: function(cells, opts) {
    return {
      "w:tr": {
        "@w:rsidR": "00995B51",
        "@w:rsidTr": "007F1D13",
        "#list": [cells] // populate this with an array of table cell objects
      }
    };
  },

  _getCell: function(val, opts, tblOpts) {
    opts = opts || {};
    // var b = {};

    // if (opts.b) {
    //   b = {
    //     "w:tc": {
    //       "w:p": {
    //         "w:r": {
    //           "w:rPr": {
    //             "w:b": {}
    //           }
    //         }
    //       }
    //     }
    //   }
    // }
    
    opts.margin = Object.assign(
      { left: 240, right: 240, top: 100, bottom: 100 },
      tblOpts.cellMargin,
      opts.margin
    ); 
    opts.spacing = Object.assign(
      { before: 0, after: 0, line: 240, lineRule: "auto" },
      tblOpts.cellSpacing,
      opts.spacing
    ); 

    var cellObj = {
      "w:tc": {
        "w:tcPr": {
          "w:tcMar": {
            "w:left": {
              "@w:w": opts.margin.left,
              "@w:type": "dxa"
            },
            "w:right": {
              "@w:w": opts.margin.right,
              "@w:type": "dxa"
            },
            "w:top": {
              "@w:w": opts.margin.top,
              "@w:type": "dxa"
            },
            "w:bottom": {
              "@w:w": opts.margin.bottom,
              "@w:type": "dxa"
            }
          },
          "w:tcW": {
            "@w:w": opts.cellColWidth || tblOpts.tableColWidth || "0",
            "@w:type": "dxa"
          },
          "w:gridSpan": {
            "@w:val" : opts.gridSpan || "1"
          },
          "w:vAlign": {
            "@w:val": opts.vAlign || "top"
          },
          "w:shd": {
            "@w:val": "clear",
            "@w:color": "auto",
            "@w:fill": opts.shd && opts.shd.fill || "",
            "@w:themeFill": opts.shd && opts.shd.themeFill || "",
            "@w:themeFillTint": opts.shd && opts.shd.themeFillTint || ""
          }
        },
        "w:p": {
          "@w:rsidR": "00995B51",
          "@w:rsidRPr": "00722E63",
          "@w:rsidRDefault": "00995B51",
          "w:pPr": {
            "w:jc": {
              "@w:val": opts.align || tblOpts.tableAlign || "center"
            },
            "w:spacing": {
              "@w:before": opts.spacing.before,
              "@w:after": opts.spacing.after,
              "@w:line": opts.spacing.line,
              "@w:lineRule": opts.spacing.lineRule
            }
            // "w:rPr": {
            //   "w:rFonts": {
            //     "@w:asciiTheme": "majorEastAsia",
            //     "@w:eastAsiaTheme": "majorEastAsia",
            //     "@w:hAnsiTheme": "majorEastAsia"
            //   },
            //   // "w:b": {},
            //   "w:sz": {
            //     "@w:val": "24"
            //   },
            //   "w:szCs": {
            //     "@w:val": "24"
            //   }
            // }
          },
          "w:r": {
            "@w:rsidRPr": "00722E63",
            "w:rPr": {
              "w:rFonts": {
                "@w:ascii": opts.fontFamily || tblOpts.tableFontFamily || "宋体",
                "@w:hAnsi": opts.fontFamily || tblOpts.tableFontFamily || "宋体"
              },
              "w:color": {
                "@w:val": opts.color || tblOpts.tableColor || "000"
              },
              "w:b": {},
              "w:sz": {
                "@w:val": opts.sz || tblOpts.sz || "24"
              },
              "w:szCs": {
                "@w:val": opts.sz || tblOpts.sz || "24"
              }
            },
            "w:t": val
          }
        }
      }
    };

    if (!opts.b) {
      delete cellObj["w:tc"]["w:p"]["w:r"]["w:rPr"]["w:b"];
    }
    if (tblOpts.rtl) {
      cellObj["w:tc"]["w:p"]["w:pPr"]["w:bidi"] = {"@w:val": "1"}
    }
    if (opts.rtl) {
      cellObj["w:tc"]["w:p"]["w:r"]["w:rPr"]["w:rtl"] = {"@w:val": "1"}
    }

    return cellObj;
  }
};
