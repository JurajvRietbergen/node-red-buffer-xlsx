const xlsx = require('better-xlsx');

module.exports = function (RED) {
    function BufferXlsx(config) {
        RED.nodes.createNode(this, config);
        this.styleMerging = config.stylemerge;
        var node = this;
        node.on('input', function (msg) {
            SimpleToXlsx(node, msg);
        });
    }

    function SimpleToXlsx(node, msg) {
        const file = new xlsx.File();
        msg.payload.forEach(sheet => {
            readSheet(sheet, file, node.styleMerging);
        });
        let type = "base64"
        // Convert to buffer before continuing node.
        file
            .saveAs(type).then(b64 => {
                msg.payload = Buffer.from(b64, 'base64');
                node.send(msg);
            }).catch(err => {
                console.log(err);
                node.send(msg);
            })
    }

    function readSheet(sheet, file, styleMerging) {
        let sheetStyling = null;
        let headerStyling = null;
        let columnsStyling = null;
        let add_sheet;
        Object.entries(sheet).forEach(([keyS, valueS]) => {
            switch (keyS) {
                case 'sheet_name':
                    add_sheet = file.addSheet(valueS);
                    break;
                case 'sheet_styling':
                    sheetStyling = valueS;
                    break;
                case 'header_styling':
                    headerStyling = valueS;
                    break;
                case 'columns_styling':
                    columnsStyling = valueS;
                    break;
                case 'rows':
                    valueS.forEach((row, index_row) => {
                        readRow(add_sheet, row, index_row, sheetStyling, headerStyling, columnsStyling, styleMerging)
                    });
                    break;
                default:
                    break;
            }
        })
    }

    // Row JSON Reading
    function readRow(add_sheet, row, index_row, sheetStyling, headerStyling, columnsStyling, styleMerging) {
        const add_row = add_sheet.addRow();
        let rowStyling = null;

        Object.entries(row).forEach(([keyC, valueC]) => {
            switch (keyC) {
                case 'row_styling':
                    rowStyling = valueC;
                    break;
                case 'cells':
                    valueC.forEach((cell, index_cell) => {
                        readCell(add_row, cell, index_row, index_cell, sheetStyling, headerStyling, columnsStyling, rowStyling, styleMerging)
                    });
                    break;
                default:
                    break;
            }
        })
    }

    // Cell JSON Reading
    function readCell(add_row, cell, index_row, index_cell, sheetStyling, headerStyling, columnsStyling, rowStyling, styleMerging) {
        const add_cell = add_row.addCell();
        let cellStyling = null;

        Object.entries(cell).forEach(([keyL, valueL]) => {
            switch (keyL) {
                case 'cell_value':
                    add_cell.value = valueL;
                    break;
                case 'cell_styling':
                    cellStyling = valueL;
                    break;
                default:
                    break;
            }
        })
        styleCell(add_cell, index_row, index_cell, sheetStyling, headerStyling, columnsStyling, rowStyling, cellStyling, styleMerging);
    }

    // Cell Styling
    function styleCell(add_cell, index_row, index_cell, sheetStyling, headerStyling, columnsStyling, rowStyling, cellStyling, styleMerging) {

        // Based on following priority styling is chosen:
        // 1. Cell
        // 2. Header
        // 3. Column
        // 3. Row
        // 4. Sheet
        let stylePriority = null;
        let columnStyling = null;
        if (columnsStyling) {
            columnStyling = columnsStyling.find(i => i.index === index_cell);
        }

        const style = new xlsx.Style();

        if (cellStyling) {
            stylePriority = cellStyling;
        } else if (index_row === 0) {
            stylePriority = headerStyling;
        } else if (columnStyling) {
            stylePriority = columnStyling.column_styling;
        } else if (rowStyling) {
            stylePriority = rowStyling;
        } else if (sheetStyling) {
            stylePriority = sheetStyling;
        }

        console.log(styleMerging);
        if (styleMerging) {
            let styles = [cellStyling, headerStyling, columnStyling, rowStyling, sheetStyling];
            styles.forEach(style => {
                if (style) {
                    Object.entries(style).forEach(([keyT, valueT]) => {
                        let found = false;
                        Object.entries(stylePriority).forEach(([keyStyle, valueStyle]) => {
                            if (keyStyle === keyT) {
                                console.log("Found: " + keyStyle + " = " + keyT);
                                found = true;
                            }
                        })
                        if (!found) {
                            stylePriority[keyT] = valueT;
                        }
                        console.log(stylePriority);
                    })
                }
            });
        }

        if (stylePriority) {
            Object.entries(stylePriority).forEach(([keySt, valueSt]) => {

                // Styling parameters
                switch (keySt) {
                    case 'pattern_type':
                        // Fill Parameters
                        style.fill.patternType = valueSt;
                        break;
                    case 'fgColor':
                        style.fill.fgColor = valueSt;
                        break;
                    case 'bgColor':
                        style.fill.bgColor = valueSt;
                        break;
                    case 'hAlign':
                        // Align Parameters
                        style.align.h = valueSt;
                        break;
                    case 'vAlign':
                        style.align.v = valueSt;
                        break;
                    case 'indent':
                        style.align.indent = valueSt;
                        break;
                    case 'shrinkToFit':
                        style.align.shrinkToFit = valueSt;
                        break;
                    case 'textRotation':
                        style.align.textRotation = valueSt;
                        break;
                    case 'wrapText':
                        style.align.wrapText = valueSt;
                        break;
                    case 'fSize':
                        // Font parameters
                        style.font.size = valueSt;
                        break;
                    case 'fName':
                        style.font.name = valueSt;
                        break;
                    case 'fFamily':
                        style.font.family = valueSt;
                        break;
                    case 'fCharset':
                        style.font.charset = valueSt;
                        break;
                    case 'fColor':
                        style.font.color = valueSt;
                        break;
                    case 'fBold':
                        style.font.bold = valueSt;
                        break;
                    case 'fItalic':
                        style.font.italic = valueSt;
                        break;
                    case 'fUnderline':
                        style.font.underline = valueSt;
                        break;
                    case 'cell_format':
                        add_cell.numFmt = valueSt;
                        break;
                    case 'cell_formula':
                        add_cell.setFormula = valueSt;
                        break;
                    case 'borders':
                        Object.entries(valueSt).forEach(([keyBr, valueBr]) => {
                            switch (keyBr) {
                                case 'all':
                                    style.border.top = valueBr.style;
                                    style.border.topColor = valueBr.bColor;
                                    style.border.right = valueBr.style;
                                    style.border.rightColor = valueBr.bColor;
                                    style.border.bottom = valueBr.style;
                                    style.border.bottomColor = valueBr.bColor;
                                    style.border.left = valueBr.style;
                                    style.border.leftColor = valueBr.bColor;
                                case 'top':
                                    style.border.top = valueBr.style;
                                    style.border.topColor = valueBr.bColor;
                                    break;
                                case 'right':
                                    style.border.right = valueBr.style;
                                    style.border.rightColor = valueBr.bColor;
                                    break;
                                case 'bottom':
                                    style.border.bottom = valueBr.style;
                                    style.border.bottomColor = valueBr.bColor;
                                    break;
                                case 'left':
                                    style.border.left = valueBr.style;
                                    style.border.leftColor = valueBr.bColor;
                                    break;
                                default:
                                    break;
                            }
                        })
                        break;
                    default:
                        break;
                }

                // Border parameters
                // TODO: Add border styling
            })
        }
        add_cell.style = style;
    }

    RED.nodes.registerType("buffer-xlsx", BufferXlsx);
}