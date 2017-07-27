//鼠标和键盘事件

// var UtilModule=require('Util')
// var Util=UtilModule.Util
// var util=new Util()

var config = require('config')

var SheetRenderModule = require('SheetRender')
var SheetRender = SheetRenderModule.SheetRender
var SliderBarHandler = require('SliderBarHandler').SliderBarHandler

var SheetEventHandler = function (sheet) {

    this.sheet = sheet
    this.sheetRender = new SheetRender(sheet)
	this.sliderBarHandler = new SliderBarHandler(sheet)
}
/**
 * 鼠标按下事件
 * @param element
 */
SheetEventHandler.prototype.mouseDown = function (elementid) {
    //var cellPropDiv=document.getElementById("cellProp")
    //cellPropDiv.style.display='none'
    //cellPropDiv.innerHTML=''
    if (!this.sheet.isEditing) {
        //if (!this.sheet.firstCellid) {
        this.sheet.firstCellid = elementid
        this.sheet.lastCellid = this.sheet.firstCellid
    }

    this.setCellBackgroundColor('#fff')
    if (elementid) {
        //setCellProp(element,this.sheet.cells[element.id])
        this.sheet.firstCellid = elementid
    }
    document.getElementById('dragBar').style.display = 'none'
    this.sheet.isMouseDown = true
}
SheetEventHandler.prototype.dragBarMouseDown = function () {
    this.setCellBackgroundColor('#fff')
    this.sheet.firstCellid = this.sheet.lastCellid
    this.sheet.isDraging = true
}
/**
 * 鼠标移动事件
 * @param element
 */
SheetEventHandler.prototype.mouseMove = function (elementid) {
    this.setCellBackgroundColor('#fff')
    if (elementid) {
        //var cellPropDiv=document.getElementById("cellProp")
        //cellPropDiv.style.display='none'
        //cellPropDiv.innerHTML=''
        this.sheet.lastCellid = elementid
        this.sheet.editCells = this.sheet.getColAndRow()
        if (this.sheet.isDraging) this.setCellBackgroundColor('#f69')
        else this.setCellBackgroundColor('#69f')
    }
}
/**
 * 鼠标抬起事件
 * @param element
 */
SheetEventHandler.prototype.mouseUp = function (element) {
    if (this.sheet.isDraging) {
        if (!this.sheet.cells[this.sheet.firstCellid]) {
            this.sheet.setAttr('content', '')
        } else {
            this.sheet.setAttr('content', this.sheet.cells[this.sheet.firstCellid].content)
        }

        this.sheet.render()
    }
    if (this.sheet.isMouseDown || this.sheet.isDraging) {
        if (element.id) {
            var dragBar = document.getElementById('dragBar')
            dragBar.style.left = element.offsetLeft +element.offsetWidth-document.getElementById(config.SheetTableDivConfig.id).scrollLeft - 8 + 'px'
            dragBar.style.top = element.offsetTop +element.offsetHeight -document.getElementById(config.SheetTableDivConfig.id).scrollTop  - 8 + 'px'
            dragBar.style.display = 'block'
            this.sheet.lastCellid = element.id
            this.sheet.editCells = this.sheet.getColAndRow()
            this.setCellBackgroundColor('#69f')

			this.sliderBarHandler.autoOpen()
			this.sliderBarHandler.addCellAttr()

        }
    }
    this.sheet.isMouseDown = false
    this.sheet.isDraging = false
}
SheetEventHandler.prototype.dragBarMouseUp = function () {
    this.setCellBackgroundColor('#69f')
    this.sheet.isMouseDown = false
    this.sheet.isDraging = false
}
/**
 * 鼠标双击事件
 * @param element
 */
SheetEventHandler.prototype.dblclick = function (element) {
    if (element.id) {
        var input = document.getElementById('input')
        input.style.display = 'block'
        input.style.backgroundColor = '#efe'
        input.style.left = element.offsetLeft -document.getElementById(config.SheetTableDivConfig.id).scrollLeft + 'px'
        input.style.top = element.offsetTop -document.getElementById(config.SheetTableDivConfig.id).scrollTop+ 'px'
        input.style.height = element.offsetHeight + 'px'
        input.style.width = element.offsetWidth + 'px'
        if (this.sheet.cells[element.id]) {
            input.value = this.sheet.cells[element.id].content
        }
        else {
            input.value = ''
        }
        input.focus()
        this.sheet.isEditing = true
        this.sheet.isMouseDown = false
    }
}
SheetEventHandler.prototype.inputBlur = function () {
    this.sheet.isEditing = false
    var input = document.getElementById('input')
    input.blur()
    input.style.display = 'none'
    //var ele = this.sheet.lastCell
    if (input.value) {
        this.sheet.setAttr('content', input.value)
        this.sheet.render()
        // var sp = document.getElementById('sp')
        // sp.value = input.value
        // sp.style.font = ele.firstChild.style.font
        // sp.style.fontSize = ele.firstChild.style.fontSize
        // this.sheet.cells[ele.id].viewWidth = sp.offsetWidth
        // if(this.sheet.cells[ele.id].autoLF){
        //     ele.width = this.sheet.cells[ele.id].width
        // }else{
        //     ele.width = this.sheet.cells[ele.id].viewWidth
        // }
    }
    this.sheet.isEditing = false
}
SheetEventHandler.prototype.inputFocus = function () {
    this.sheet.isEditing = true
    var input = document.getElementById('input')
    input.focus()
}
// //鼠标失去焦点事件
// SheetEventHandler.prototype.mouseBlur=function(element){
//     element.style.color='transparent'
//     element.style.background='transparent'
//     if(config.WSConfig.isKeyDown){
//         var value=element.value
//         if(element.nextSibling.innerHTML==''&&value!=''){
//             var command = 'set '+ element.id + ' ' + value
//             var type = 'set'
//             this.sheet.UndoStack.pushCommand(command,type)
//         }
//         element.nextSibling.innerHTML=value
//     }
//     config.WSConfig.isKeyDown=true
//     element.parentNode.style.backgroundColor = 'transparent'
// }
/**
 * 设置单元格背景颜色
 * @param backgroundColor
 */
SheetEventHandler.prototype.setCellBackgroundColor = function (backgroundColor) {
    var cells = this.sheet.editCells
    if (cells != null) {
        for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
            for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                var ele = document.getElementById(String.fromCharCode(i) + '_' + j)
                if (ele) {
                    if (backgroundColor == '#fff'
                        && this.sheet.cells[ele.id] != undefined
                        && this.sheet.cells[ele.id].backgroundColor != undefined) {
                        ele.style.backgroundColor = this.sheet.cells[ele.id].backgroundColor
                    } else {
                        ele.style.backgroundColor = backgroundColor
                    }
                }
            }
        }
    }
}

/**
 * 键盘按下事件
 * @param element
 * @param event
 */
SheetEventHandler.prototype.keyDown = function (event) {
    if (!this.sheet.isMultiLineEditing) {
        var col = this.sheet.editCells.lastCellCol//element.id.split('_')[0].charCodeAt(0)
        var row = this.sheet.editCells.lastCellRow//parseInt(element.id.split('_')[1])
        //console.log(event.which)
        switch (event.which) {
            case 37://左键
                do {
                    col = col - 1
                } while (this.sheet.cells[String.fromCharCode(col) + '_' + row] && !this.sheet.cells[String.fromCharCode(col) + '_' + row].show)
                if (!this.sheet.isEditing && col > 64) {
                    var cellid = String.fromCharCode(col) + '_' + row

                    this.mouseDown(cellid)
                    this.mouseUp(document.getElementById(cellid))

                }
                return false
            case 38://上键
                do {
                    row = row - 1
                } while (this.sheet.cells[String.fromCharCode(col) + '_' + row] && !this.sheet.cells[String.fromCharCode(col) + '_' + row].show)
                var cellid = String.fromCharCode(col) + '_' + row

                if (this.sheet.isEditing) {
                    this.sheet.isEditing = false
                    document.getElementById('input').blur()
                }
                if (row > 0) {
                    this.mouseDown(cellid)
                    this.mouseUp(document.getElementById(cellid))
                }
                return false
            case 39://右键
                do {
                    col = col + 1
                } while (this.sheet.cells[String.fromCharCode(col) + '_' + row] && !this.sheet.cells[String.fromCharCode(col) + '_' + row].show)
                if (!this.sheet.isEditing) {
                    var cellid = String.fromCharCode(col) + '_' + row
                    if (document.getElementById(cellid)) {
                        this.mouseDown(cellid)
                        this.mouseUp(document.getElementById(cellid))
                    }
                }
                return false
            case 40://下键
                do {
                    row = row + 1
                } while (this.sheet.cells[String.fromCharCode(col) + '_' + row] && !this.sheet.cells[String.fromCharCode(col) + '_' + row].show)
                var cellid = String.fromCharCode(col) + '_' + row
                if (document.getElementById(cellid)) {
                    if (this.sheet.isEditing) {
                        this.sheet.isEditing = false
                        document.getElementById('input').blur()
                    }
                    this.mouseDown(cellid)
                    this.mouseUp(document.getElementById(cellid))
                }
                return false
            case 17://ctrl
                break
            case 18://alt
                break
            case 67://ctrl+c
                var ta = document.getElementById('ta')
                var text = ''
                //var firstCell=this.sheet.firstCell
                //var lastCell=this.sheet.lastCell
                var cells = this.sheet.getColAndRow()
                if (cells != null) {
                    var rowCount = 0
                    for (var i = cells.firstCellRow; i <= cells.lastCellRow; i++) {
                        var colCount = 0
                        for (var j = cells.firstCellCol; j <= cells.lastCellCol; j++) {
                            var eleOld = document.getElementById(
                                String.fromCharCode(j)
                                + '_' + i)
                            text += eleOld.firstChild.innerHTML + '\t'
                            colCount++
                        }
                        text = text.substr(0, text.length - 1)
                        text += '\n'
                        rowCount++
                    }
                    text = text.substr(0, text.length - 1)
                }
                ta.value = text
                ta.style.display = 'block'
                ta.focus()
                ta.select()

                window.setTimeout(function () {
                    ta.style.display = 'none'
                    var cell = document.getElementById(col + '_' + row)
                    //cell.focus()
                }, 200)

                break
            case 86://ctrl+v
                var ta = document.getElementById('ta')
                ta.style.display = 'block'
                ta.focus()
                ta.select()
                var sheet = this.sheet
                window.setTimeout(function () {
                    ta.blur()
                    var v = ta.value
                    ta.style.display = 'none'
                    v = v.split('\n')
                    for (var i = 0; i < v.length; i++) {
                        var vv = v[i].split('\t')
                        for (var j = 0; j < vv.length; j++) {
                            var cvCol = col + j
                            var cvRow = row + i
                            var cells = {}
                            cells.firstCellCol = cvCol
                            cells.firstCellRow = cvRow
                            cells.lastCellCol = cvCol
                            cells.lastCellRow = cvRow
                            sheet.setAttr('content', vv[j], cells)
                        }
                    }
                    var cell = document.getElementById(String.fromCharCode(col) + '_' + row)
                    cell.focus()
                    sheet.render()
                }, 200)

                this.sheet.isKeyDown = false
                break
            case 88://ctrl+x
                var ta = document.getElementById('ta')
                var text = ''
                //var firstCell=this.sheet.firstCell
                //var lastCell=this.sheet.lastCell
                var cells = this.sheet.getColAndRow()
                if (cells != null) {
                    var rowCount = 0
                    for (var i = cells.firstCellRow; i <= cells.lastCellRow; i++) {
                        var colCount = 0
                        for (var j = cells.firstCellCol; j <= cells.lastCellCol; j++) {
                            var eleOld = document.getElementById(
                                String.fromCharCode(j)
                                + '_' + i)
                            text += eleOld.firstChild.innerHTML + '\t'
                            colCount++
                        }
                        text = text.substr(0, text.length - 1)
                        text += '\n'
                        rowCount++
                    }
                    text = text.substr(0, text.length - 1)
                }
                ta.value = text
                ta.style.display = 'block'
                ta.focus()
                ta.select()

                window.setTimeout(function () {
                    ta.style.display = 'none'
                    var cell = document.getElementById(String.fromCharCode(col) + '_' + row)
                    cell.focus()
                }, 200)

                this.sheet.setAttr('content', '')
                this.sheet.render()
                break
            case 8://backspace
                this.sheet.setAttr('content', '')
                this.sheet.render()
                break
            case 46://delete
                var sheet = this.sheet
                sheet.cellAttrs.forEach(function (attr) {
                    sheet.setAttr(attr[0], attr[1])
                })
                sheet.render()
                break
            default:
                if (!this.sheet.isEditing) {
                    if (this.sheet.lastCellid) this.dblclick(document.getElementById(this.sheet.lastCellid))
                }
        }
    }
}

SheetEventHandler.prototype.multiLineBlur = function (text) {
    this.sheet.setAttr('content', text)
    this.sheet.setAttr('wrapText', true)
    this.sheet.render()
    this.sheet.isMultiLineEditing = false
}

//获取元素的纵坐标

function getTop(e) {

    var offset = e.offsetTop;

    if (e.offsetParent != null) offset += getTop(e.offsetParent);

    return offset;

}


//获取元素的横坐标

function getLeft(e) {

    var offset = e.offsetLeft

    if (e.offsetParent != null) offset += getLeft(e.offsetParent)

    return offset

}

module.exports.SheetEventHandler = SheetEventHandler