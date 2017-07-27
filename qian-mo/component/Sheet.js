/**
 * 表格对象
 */


var configModule = require('config')
var CellModule = require('Cell')
var Cell = CellModule.Cell
var sheetConfig = configModule.SheetConfig
// var UtilModule = require('Util')
// var Util = UtilModule.Util
// var util = new Util()
var CellRenderModule = require('CellRender')
var CellRender = CellRenderModule.CellRender
//var cellRender = new CellRender
var Sheet = function (UndoStack) {
    this.cellRender = new CellRender(this)
    this.height = sheetConfig.height
    this.width = sheetConfig.width

    this.headWidth = sheetConfig.headWidth
    this.headHeight = sheetConfig.headHeight


    this.rowNum = sheetConfig.rowNum
    this.colNum = sheetConfig.colNum

    this.maxCol = 65
    this.maxRow = 1

    this.cells = {}
    this.firstCopiedCell = null
    this.copiedCells = {}
    // todo cellAttrs
    this.cellAttrs = [['backgroundColor', ''], ['foregroundColor', ''], ['bold', false], ['italic', false],
        ['content', ''], ['fontSize', ''], ['alignment', 'left'], ['font', ''], ['show', true],
        ['colSpan', 1], ['rowSpan', 1]]
    //['bottomFrame', false], ['leftFrame', false], ['topFrame', false], ['rightFrame', false]
    this.UndoStack = UndoStack
    this.undostackCache = []

    this.isMouseDown = false
    this.isKeyDown = false
    this.isEditing = false
    this.isMultiLineEditing = false
    this.isDraging = false

    //this.isFont = false
    this.isPreview = false

    this.firstCellid = 'A_1'
    this.lastCellid = 'A_1'
    this.editCells = this.getColAndRow()
}

Sheet.prototype.addCell = function (id) {
    if (!this.cells[id]) {
        this.cells[id] = new Cell(id)
        if (id.split('_')[0].charCodeAt(0) > this.maxCol) this.maxCol = id.split('_')[0].charCodeAt(0)
        if (parseInt(id.split('_')[1]) > this.maxRow) this.maxRow = parseInt(id.split('_')[1])
    }
}

Sheet.prototype.setAttr = function (attr, value, cells) {
    if (!cells) var cells = this.editCells
    if (cells != null) {
        switch (attr) {
            //缩进
            case 'indentation':
                if (value > 0) {
                    var html = '';
                    for (var i = 0; i < value; i++) {
                        html += '&nbsp;'
                    }
                    var i = cells.firstCellCol
                    var j = cells.firstCellRow
                    var id = String.fromCharCode(i) + '_' + j
                    if (this.cells[id]) {
                        if (this.cells[id]['content']) {
                            if (this.cells[id].show) {
                                html = html + this.cells[id]['content']
                                this.undostackCache.push([id, 'content', html, this.cells[id]['content']])
                                this.cells[id]['content'] = html
                            }
                        }
                    }
                }
                break
            case 'cellname':
            case 'area':
                break
            case 'merge':
                if (value) {
                    for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                        for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                            var id = String.fromCharCode(i) + '_' + j
                            this.addCell(id)
                            if (this.cells[id].show) {
                                if (i === cells.firstCellCol && j === cells.firstCellRow) {
                                    var colSpan = cells.lastCellCol - cells.firstCellCol + 1
                                    var rowSpan = cells.lastCellRow - cells.firstCellRow + 1
                                    if (colSpan !== this.cells[id].colSpan || rowSpan !== this.cells[id].rowSpan) {
                                        this.undostackCache.push([id, 'colSpan', colSpan, this.cells[id].colSpan])
                                        this.undostackCache.push([id, 'rowSpan', rowSpan, this.cells[id].rowSpan])
                                        //this.undostackCache.push([id, 'area', [colSpan, rowSpan], [this.cells[id].colSpan, this.cells[id].rowSpan]])
                                        this.cells[id].colSpan = colSpan
                                        this.cells[id].rowSpan = rowSpan
                                    }
                                }
                                else {
                                    this.cells[id].show = false
                                    this.undostackCache.push([id, 'show', false, true])
                                    this.undostackCache.push([id, 'colSpan', cells.firstCellCol - i + 1, this.cells[id].colSpan])
                                    this.undostackCache.push([id, 'rowSpan', cells.firstCellRow - j + 1, this.cells[id].rowSpan])
                                    this.cells[id].colSpan = cells.firstCellCol - i + 1
                                    this.cells[id].rowSpan = cells.firstCellRow - j + 1
                                }
                            }
                        }
                    }
                    this.lastCellid = this.firstCellid
                }
                else {
                    this.firstCellid = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
                    this.lastCellid = String.fromCharCode(cells.lastCellCol) + '_' + cells.lastCellRow
                    for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                        for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                            var id = String.fromCharCode(i) + '_' + j
                            if (!this.cells[id]) {
                                continue
                            }
                            if (this.cells[id].colSpan != 1 || this.cells[id].rowSpan != 1) {
                                this.undostackCache.push([id, 'colSpan', 1, this.cells[id].colSpan])
                                this.undostackCache.push([id, 'rowSpan', 1, this.cells[id].rowSpan])
                                this.cells[id].colSpan = 1
                                this.cells[id].rowSpan = 1
                            }
                            if (!this.cells[id].show) {
                                this.undostackCache.push([id, 'show', true, false])
                                this.cells[id].show = true
                            }
                        }
                    }
                }
                break
            case 'leftFrame':
            case 'topFrame':
            case 'rightFrame':
            case 'bottomFrame':
                switch (attr) {
                    case 'leftFrame':
                        var mini = cells.firstCellCol
                        var minj = cells.firstCellRow
                        var maxi = cells.firstCellCol
                        var maxj = cells.lastCellRow
                        break
                    case 'topFrame':
                        var mini = cells.firstCellCol
                        var minj = cells.firstCellRow
                        var maxi = cells.lastCellCol
                        var maxj = cells.firstCellRow
                        break
                    case 'rightFrame':
                        var mini = cells.lastCellCol
                        var minj = cells.firstCellRow
                        var maxi = cells.lastCellCol
                        var maxj = cells.lastCellRow
                        break
                    case 'bottomFrame':
                        var mini = cells.firstCellCol
                        var minj = cells.lastCellRow
                        var maxi = cells.lastCellCol
                        var maxj = cells.lastCellRow
                        break

                }
                for (var i = mini; i <= maxi; i++) {
                    for (var j = minj; j <= maxj; j++) {
                        var id = String.fromCharCode(i) + '_' + j
                        this.addCell(id)
                        if (this.cells[id][attr] != value) {
                            this.cells[id][attr] = value
                            this.undostackCache.push([id, attr, value, !value])
                        }
                    }
                }
                break
            case 'allFrameEx':
                for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                    for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                        var id = String.fromCharCode(i) + '_' + j
                        this.addCell(id)
                        // var env = this
                        //
                        // if (env.cells[id].rightFrame != value && i !== cells.firstCellCol - 1) {
                        //     env.undostackCache.push([id, 'topFrame', value, !value])
                        //     env.cells[id].rightFrame = value
                        // }
                        // if (env.cells[id].bottomFrame != value && j !== cells.firstCellRow - 1) {
                        //     env.undostackCache.push([id, 'rightFrame', value, !value])
                        //     env.cells[id].bottomFrame = value
                        // }
                        var env = this

                        var a = ['leftFrame', 'rightFrame', 'topFrame', 'bottomFrame']

                        a.forEach(function (frame) {
                            if (env.cells[id][frame] != value) {
                                env.undostackCache.push([id, frame, value, !value])
                                env.cells[id][frame] = value
                            }
                        })

                    }
                }
                if (value == false) {
                    if (cells.firstCellCol > 65) {
                        var firstCellid = String.fromCharCode(cells.firstCellCol - 1)
                            + '_' + cells.firstCellRow
                        var lastCellid = String.fromCharCode(cells.firstCellCol - 1)
                            + '_' + cells.lastCellRow
                        var newCell = this.getColAndRow(firstCellid, lastCellid)
                        this.setAttr('rightFrame', false, newCell)
                        delete  newCell
                    }
                    if (cells.firstCellRow > 1) {

                        var firstCellid = String.fromCharCode(cells.firstCellCol)
                            + '_' + (cells.firstCellRow - 1)
                        var lastCellid = String.fromCharCode(cells.lastCellCol)
                            + '_' + (cells.firstCellRow - 1)
                        var newCell = this.getColAndRow(firstCellid, lastCellid)
                        this.setAttr('bottomFrame', false, newCell)
                        delete  newCell
                    }
                    if (document.getElementById(String.fromCharCode(cells.lastCellCol)
                            + '_' + (cells.lastCellRow + 1))) {
                        var firstCellid = String.fromCharCode(cells.firstCellCol)
                            + '_' + (cells.lastCellRow + 1)
                        var lastCellid = String.fromCharCode(cells.lastCellCol)
                            + '_' + (cells.lastCellRow + 1)
                        var newCell = this.getColAndRow(firstCellid, lastCellid)
                        this.setAttr('topFrame', false, newCell)
                        delete  newCell
                    }
                    if (document.getElementById(String.fromCharCode(cells.lastCellCol + 1)
                            + '_' + (cells.lastCellRow))) {
                        var firstCellid = String.fromCharCode(cells.lastCellCol + 1)
                            + '_' + (cells.firstCellRow)
                        var lastCellid = String.fromCharCode(cells.lastCellCol + 1)
                            + '_' + (cells.lastCellRow )
                        var newCell = this.getColAndRow(firstCellid, lastCellid)
                        this.setAttr('leftFrame', false, newCell)
                        delete  newCell
                    }
                }
                break
            case 'allFrame':
                for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                    for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                        var id = String.fromCharCode(i) + '_' + j
                        this.addCell(id)
                        var env = this

                        var a = ['leftFrame', 'rightFrame', 'topFrame', 'bottomFrame']

                        a.forEach(function (frame) {
                            if (env.cells[id][frame] != value) {
                                env.undostackCache.push([id, frame, value, !value])
                                env.cells[id][frame] = value
                            }
                        })

                    }
                }
                break
            default:
                for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                    for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                        var id = String.fromCharCode(i) + '_' + j
                        this.addCell(id)

                            if (this.cells[id][attr] !== value) {
                                this.undostackCache.push([id, attr, value, this.cells[id][attr]])
                                this.cells[id][attr] = value
                            }

                    }
                }
                break
        }
    }
}
//将value赋值给attr属性，赋值范围由sheet.firstcell和sheet.lastcell指定

Sheet.prototype.sheet2json = function () {
//todo
}

Sheet.prototype.json2sheet = function () {
//todo
}

Sheet.prototype.render = function () {
    var sheet = this
    this.undostackCache.forEach(function (cache) {
        sheet.cellRender.renderCell(cache[0], cache[1], cache[2])
    })
    this.UndoStack.pushStep(this.undostackCache)
    this.undostackCache = []
    //遍历sheet.undostackCache,将sheet.undostackCache压入sheet.undostack
}
Sheet.prototype.renderCell = function (id, cmd, value) {
    this.cellRender.renderCell(id, cmd, value)
}

Sheet.prototype.undo = function () {
    var undoSteps = this.UndoStack.unDo()
    if (undoSteps) {
        var sheet = this
        undoSteps.forEach(function (step) {
            sheet.cells[step[0]][step[1]] = step[3]
            sheet.cellRender.renderCell(step[0], step[1], step[3])
        })
    }
}

Sheet.prototype.redo = function () {
    var undoSteps = this.UndoStack.reDo()
    if (undoSteps) {
        var sheet = this
        undoSteps.forEach(function (step) {
            sheet.cells[step[0]][step[1]] = step[2]
            sheet.cellRender.renderCell(step[0], step[1], step[2])
        })
    }

}

Sheet.prototype.setIdAttr = function (id,cmd,value) {
    this.cellRender.renderCell(id,cmd,value)
}

Sheet.prototype.getColAndRow = function (firstCellid, lastCellid) {
    if (firstCellid == null) {
        firstCellid = this.firstCellid
        lastCellid = this.lastCellid
    }
    //if (firstCell && firstCell.id && lastCell && lastCell.id) {
    var result = {}
    result.firstCellCol = Math.min(firstCellid.split('_')[0].charCodeAt(0), lastCellid.split('_')[0].charCodeAt(0))
    result.firstCellRow = Math.min(parseInt(firstCellid.split('_')[1]), parseInt(lastCellid.split('_')[1]))
    result.lastCellCol = Math.max(lastCellid.split('_')[0].charCodeAt(0), firstCellid.split('_')[0].charCodeAt(0))
    result.lastCellRow = Math.max(parseInt(lastCellid.split('_')[1]), parseInt(firstCellid.split('_')[1]))
    //}
    var sheet = this
    return getRectangle(result, sheet)
}

getRectangle = function (cells, sheet) {
    //console.log(cells)
    var flag = false
    var minCol = cells.firstCellCol
    var minRow = cells.firstCellRow
    var maxCol = cells.lastCellCol
    var maxRow = cells.lastCellRow

    for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
        for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
            var id = String.fromCharCode(i) + '_' + j
            if (!sheet.cells[id]) {
                continue

            }
            if (!sheet.cells[id].show) {
                if ((i + sheet.cells[id].colSpan - 1) < minCol) {
                    flag = true
                    minCol = (i + sheet.cells[id].colSpan - 1)

                }
                if ((j + sheet.cells[id].rowSpan - 1) < minRow) {
                    flag = true
                    minRow = (j + sheet.cells[id].rowSpan - 1)
                }
            }
            else {
                if (sheet.cells[id].colSpan - 1 + i > maxCol) {
                    flag = true
                    maxCol = sheet.cells[id].colSpan - 1 + i
                }
                if (sheet.cells[id].rowSpan - 1 + j > maxRow) {
                    flag = true
                    maxRow = sheet.cells[id].rowSpan - 1 + j
                }
            }

        }
    }
    cells.firstCellCol = minCol
    cells.firstCellRow = minRow
    cells.lastCellCol = maxCol
    cells.lastCellRow = maxRow
    if (flag) return getRectangle(cells, sheet)
    else return cells
}


module.exports.Sheet = Sheet