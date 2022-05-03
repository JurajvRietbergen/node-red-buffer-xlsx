const xlsx = require('better-xlsx');

module.exports = function (RED) {
    function BufferXlsx(config) {
        RED.nodes.createNode(this, config);
        this.complex = config.complex;
        var node = this;
        node.on('input', function (msg) {
            if (node.complex) {
                ComplexToXlsx(node, msg);
            } else {
                SimpleToXlsx(node, msg);
            }
        });
    }

    function ComplexToXlsx(node, msg) {
        msg.payload = "TBD";
        node.send(msg);
    }

    function SimpleToXlsx(node, msg) {
        const file = new xlsx.File();
        console.log(node.complex);
        msg.payload.forEach(sheet => {
            let sheetStyling = null;
            let headerStyling = null;
            let add_sheet;
            Object.entries(sheet).forEach(([keyS, valueS]) => {
                if (keyS === 'sheet_name') {
                    add_sheet = file.addSheet(valueS);
                } else if (keyS === 'sheet_styling') {
                    sheetStyling = valueS;
                } else if (keyS === 'header_styling') {
                    headerStyling = valueS;
                } else if (keyS === 'rows') {
                    valueS.forEach((row, index) => {
                        const add_row = add_sheet.addRow();
                        let rowStyling = null;
                        Object.entries(row).forEach(([keyC, valueC]) => {
                            if (keyC === 'row_styling') {
                                rowStyling = valueC;
                            } else if (keyC === 'cells') {
                                valueC.forEach((cell) => {
                                    const add_cell = add_row.addCell();
                                    let cellStyling = null;
                                    Object.entries(cell).forEach(([keyL, valueL]) => {
                                        if (keyL === 'cell_value') {
                                            add_cell.value = valueL;
                                        } else if (keyL === 'cell_format') {
                                            add_cell.numFmt = valueL;
                                        } else if (keyL === 'cell_formula') {
                                            add_cell.setFormula = valueL;
                                        }
                                        else if (keyL === 'cell_styling') {
                                            cellStyling = valueL;
                                        }
                                    })

                                    // Based on following priority styling is chosen:
                                    // 1. Header
                                    // 2. Cell
                                    // 3. Row
                                    // 4. Sheet
                                    let stylePriority = null;
                                    const style = new xlsx.Style();
                                    if (index === 0) {
                                        stylePriority = headerStyling;
                                    } else if (cellStyling) {
                                        stylePriority = cellStyling;
                                    } else if (rowStyling) {
                                        stylePriority = rowStyling;
                                    } else if (sheetStyling) {
                                        stylePriority = sheetStyling;
                                    }

                                    Object.entries(stylePriority).forEach(([keySt, valueSt]) => {
                                        // Styling parameters
                                        // Fill parameters
                                        if (keySt === 'pattern_type') {
                                            style.fill.patternType = valueSt;
                                        } else if (keySt === 'fgColor') {
                                            style.fill.fgColor = valueSt;
                                        } else if (keySt === 'bgColor') {
                                            style.fill.bgColor = valueSt;
                                        }

                                        // Align parameters
                                        else if (keySt === 'hAlign') {
                                            style.align.h = valueSt;
                                        } else if (keySt === 'vAlign') {
                                            style.align.v = valueSt;
                                        } else if (keySt === 'indent') {
                                            style.align.indent = valueSt;
                                        } else if (keySt === 'shrinkToFit') {
                                            style.align.shrinkToFit = valueSt;
                                        } else if (keySt === 'textRotation') {
                                            style.align.textRotation = valueSt;
                                        } else if (keySt === 'wrapText') {
                                            style.align.wrapText = valueSt;
                                        }

                                        // Font parameters
                                        else if (keySt === 'fSize') {
                                            style.font.sz = valueSt;
                                        } else if (keySt === 'fName') {
                                            style.font.name = valueSt;
                                        } else if (keySt === 'fFamily') {
                                            style.font.family = valueSt;
                                        } else if (keySt === 'fCharset') {
                                            style.font.charset = valueSt;
                                        } else if (keySt === 'fColor') {
                                            style.font.color = valueSt;
                                        } else if (keySt === 'fBold') {
                                            style.font.bold = valueSt;
                                        } else if (keySt === 'fItalic') {
                                            style.font.italic = valueSt;
                                        } else if (keySt === 'fUnderline') {
                                            style.font.underline = valueSt;
                                        }

                                        // Border parameters
                                        // TODO: Add border styling
                                    })
                                    add_cell.style = style;
                                });
                            }
                        })
                    });
                }
            })
        });
        let type = "base64"
        // Convert to buffer before continueing node.
        file
            .saveAs(type).then(b64 => {
                msg.payload = Buffer.from(b64, 'base64');
                msg.send(msg);
            }).catch(err => {
                console.log(err);
                node.send(msg);
            })


    }
    RED.nodes.registerType("buffer-xlsx", BufferXlsx);
}