/**
 工具栏绑定事件
 */

var ToolEventHandlerModule = require('ToolEventHandler')
var ToolEventHandler = ToolEventHandlerModule.ToolEventHandler
var toolEventHandler = null

var SheetEventHandlerModule = require('SheetEventHandler')
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null


var ToolEventBinder = function (sheet) {
    this.sheet = sheet
    sheetEventHandler = new SheetEventHandler(sheet)
    this.sheet.colorOrBackgroudcolor = null
    toolEventHandler = new ToolEventHandler(sheet)
}
ToolEventBinder.prototype.buttonClick = function (action) {
    toolEventHandler.buttonClick(action)
}
ToolEventBinder.prototype.initFontFamily = function (fontFamilySelect) {
    var toolEventBinder = this
    fontFamilySelect.onchange = function () {
        toolEventBinder.sheet.setAttr('font', fontFamilySelect.value)
        toolEventBinder.sheet.render()
        fontFamilySelect.value = 'fontFamily'
    }
}
ToolEventBinder.prototype.initFontSize = function (fontSizeSelect) {
    var toolEventBinder = this
    fontSizeSelect.onchange = function () {
        //todo
        //switch
        toolEventBinder.sheet.setAttr('fontSize', fontSizeSelect.value)
        toolEventBinder.sheet.render()
        fontSizeSelect.value = 'fontSize'
    }
}
ToolEventBinder.prototype.initFontWeight = function (fontWeightDiv) {
    var toolEventBinder = this
    fontWeightDiv.onclick = function () {
        var lastCellId = toolEventBinder.sheet.lastCellid
        if (!toolEventBinder.sheet.cells[lastCellId]) {
            toolEventBinder.sheet.setAttr('bold', true)
        }
        toolEventBinder.sheet.setAttr('bold', !toolEventBinder.sheet.cells[lastCellId].bold)
        toolEventBinder.sheet.render()
    }
}
ToolEventBinder.prototype.initFontStyle = function (fontStyleDiv) {
    var toolEventBinder = this
    fontStyleDiv.onclick = function () {
        var lastCellId = toolEventBinder.sheet.lastCellid
        if (!toolEventBinder.sheet.cells[lastCellId]) {
            toolEventBinder.sheet.setAttr('italic', true)
        }
        toolEventBinder.sheet.setAttr('italic', !toolEventBinder.sheet.cells[lastCellId].italic)
        toolEventBinder.sheet.render()
    }
}
ToolEventBinder.prototype.initColor = function (colorDiv) {
    var toolEventBinder = this
    colorDiv.onclick = function () {
        var colorSelectDiv = document.getElementById('colorSelect')
        if (colorSelectDiv.style.display === 'none' || toolEventBinder.fontType === 'backgroundColor') {
            toolEventBinder.colorOrBackgroudcolor = 'foregroundColor'
            colorSelectDiv.style.display = 'block'
            var left = this.offsetLeft
            var top = this.offsetTop
            colorSelectDiv.style.top = top + 18 + 'px'
            colorSelectDiv.style.left = left + 'px'
        } else {
            colorSelectDiv.style.display = 'none'
        }

    }
}
ToolEventBinder.prototype.initBackgroundColor = function (backgroundColorDiv) {
    var toolEventBinder = this
    backgroundColorDiv.onclick = function () {
        var colorSelectDiv = document.getElementById('colorSelect')
        if (colorSelectDiv.style.display === 'none' || toolEventBinder.fontType === 'foregroundColor') {
            toolEventBinder.colorOrBackgroudcolor = 'backgroundColor'
            colorSelectDiv.style.display = 'block'
            var left = this.offsetLeft
            var top = this.offsetTop
            colorSelectDiv.style.top = top + 18 + 'px'
            colorSelectDiv.style.left = left + 'px'
        } else {
            colorSelectDiv.style.display = 'none'
        }
    }
}
ToolEventBinder.prototype.initColorSelect = function (td) {
    var toolEventBinder = this
    td.onclick = function () {
        document.getElementById('colorSelect').style.display = 'none'
        toolEventBinder.sheet.setAttr(toolEventBinder.colorOrBackgroudcolor, this.style.backgroundColor)
        toolEventBinder.sheet.render()

    }
}
ToolEventBinder.prototype.initMerge = function (mergeDiv, value) {
    var toolEventBinder = this
    mergeDiv.onclick = function () {
        sheetEventHandler.setCellBackgroundColor('#fff')
        toolEventBinder.sheet.setAttr('merge', value)
        toolEventBinder.sheet.render()
        sheetEventHandler.mouseDown(toolEventBinder.sheet.firstCellid)
        sheetEventHandler.mouseUp(document.getElementById(toolEventBinder.sheet.lastCellid))
    }
}

ToolEventBinder.prototype.initFontBorder = function (value) {
    var toolEventBinder = this
    switch (value) {
        case 'left':
            toolEventBinder.sheet.setAttr('leftFrame', true)

            break
        case 'top':
            toolEventBinder.sheet.setAttr('topFrame', true)
            break
        case 'right':
            toolEventBinder.sheet.setAttr('rightFrame', true)
            break
        case 'bottom':
            toolEventBinder.sheet.setAttr('bottomFrame', true)
            break
        case 'clear':
            toolEventBinder.sheet.setAttr('allFrameEx', false)
            break
        case 'outer':
            toolEventBinder.sheet.setAttr('leftFrame', true)
            toolEventBinder.sheet.setAttr('topFrame', true)
            toolEventBinder.sheet.setAttr('rightFrame', true)
            toolEventBinder.sheet.setAttr('bottomFrame', true)
            break
        case 'all':
            toolEventBinder.sheet.setAttr('allFrameEx', true)
            break
    }
    toolEventBinder.sheet.render()
}

ToolEventBinder.prototype.initTextAlign = function (value) {
    var toolEventBinder = this
    toolEventBinder.sheet.setAttr('alignment', value)
    toolEventBinder.sheet.render()
}
ToolEventBinder.prototype.initFileInput = function (fileInput) {
    fileInput.onchange = function () {
        var ajax
        if (window.XMLHttpRequest) {
            ajax = new XMLHttpRequest()
        } else if (window.ActiveXObject) {
            ajax = new window.ActiveXObject()
        } else {
            alert("请升级至最新版本的浏览器")
        }
        if (ajax !== null) {
            ajax.open("GET", "json.json", true)
            ajax.send(null)
            ajax.onreadystatechange = function () {
                if (ajax.readyState === 4 && ajax.status === 200) {
                    var CellList = JSON.parse(ajax.responseText)
                    CellList.forEach(function (e) {
                        for (var key in e) {
                            cellRender.renderCell(e['cellName'], key, e[key])
                        }
                    })
                }
            }
        }
    }
}
module.exports.ToolEventBinder = ToolEventBinder