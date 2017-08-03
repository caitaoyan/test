//事件绑定对象
var SheetEventHandlerModule = require('SheetEventHandler')
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null

var SheetEventBinder = function (sheet) {
    this.sheet = sheet
    sheetEventHandler = new SheetEventHandler(sheet)
    window.onmouseup = function () {
        sheet.isMouseDown = false
    }
    document.onkeydown = function (event) {
        if(!sheet.isPreview)  sheetEventHandler.keyDown(event)

        if (!sheet.isMultiLineEditing && !sheet.isEditing) return false
    }
    document.onkeyup = function (event) {
        if(event.which === 17) sheet.isKeyDown = false
    }
}

SheetEventBinder.prototype.initRowTD = function (rowTD) {
    var sheet = this.sheet
    rowTD.onmousedown = function () {
        sheetEventHandler.mouseDown(rowTD.id)
    }
    rowTD.onmouseup = function () {
        sheetEventHandler.mouseUp(rowTD)
    }
    rowTD.onmousemove = function () {
        if (sheet.isMouseDown || sheet.isDraging) {
            sheetEventHandler.mouseMove(rowTD.id)
        }
    }
    rowTD.ondblclick = function () {
        sheetEventHandler.dblclick(this)
    }
}
SheetEventBinder.prototype.initInput = function (input) {
    input.onblur = function () {
        sheetEventHandler.inputBlur()
    }
    input.onfocus = function () {
        sheetEventHandler.inputFocus()
    }
}
SheetEventBinder.prototype.initMultiLine = function (multiLine) {
    multiLine.onblur = function () {
        sheetEventHandler.multiLineBlur(multiLine.value)
        multiLine.value = ''
        multiLine.style.display = 'none'
    }
}
SheetEventBinder.prototype.initDragBar = function (dragBar) {
    dragBar.onmousedown = function () {
        sheetEventHandler.dragBarMouseDown()
    }
    dragBar.onmouseup = function () {
        sheetEventHandler.dragBarMouseUp()
    }
}
SheetEventBinder.prototype.initFormulaButton = function (button) {
    button.onclick = function () {
        sheetEventHandler.formulaButonClick()
    }
}
module.exports.SheetEventBinder = SheetEventBinder