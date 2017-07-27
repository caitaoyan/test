/**
 * 渲染表格
 */

var config = require('config')
var style = require('style')
var SliderBarRender = require('SliderBarRender').SliderBarRender
// config.CellConfig
var SheetRender = function (sheet) {
	this.sheet=sheet
}

SheetRender.prototype.init = function (sheetDiv) {

	sheetDiv.style = style.SheetDivStyle

	var sheetTableDiv = document.createElement('div')
	sheetTableDiv.style = style.SheetTableDivStyle
	sheetTableDiv.id = config.SheetTableDivConfig.id

	var sheetTable = document.createElement('table')
    sheetTable.style = style.SheetTableStyle
	sheetTable.cellSpacing = '10px'

	var sliderBarDiv = document.createElement('div')
	sliderBarDiv.id = config.SlideBarConfig.id
	var sliderBarRender = new SliderBarRender(this.sheet, sheetDiv)
	sliderBarRender.init(sliderBarDiv, sheetTableDiv)

	//sheetTable.cellSpacing = '10px'//
	sheetTable.id = 'sheetTable'
    sheetDiv.appendChild(sheetTable)

    sheetTableDiv.appendChild(sheetTable)

	sheetDiv.appendChild(sheetTableDiv)
	sheetDiv.appendChild(sliderBarDiv)

    // var cellPropDiv = document.createElement("div")
    // cellPropDiv.id = config.CellPropConfig.id
	//

    // cellPropDiv.style = style.CellPropDivStyle

    // sheetDiv.appendChild(cellPropDiv)

	//表格的具体渲染
	this.renderSheet(this.sheet, sheetTable)
}

/**
 * 构建表格的数据模型
 * 一个存放表头的二维数组
 * @param rowNum
 * @param colNum
 */
function createHeader(rowNum, colNum) {
	var gridHeader =  []
	for(var a=0;a<rowNum+1;a++){
        gridHeader[a] = []

		for(var b=0;b<colNum+1;b++){
            gridHeader[a][b] = null
		}
	}
	var startString = "A"

	for(var i=1; i<colNum+1; i++){
        gridHeader[0][i] = String.fromCharCode(startString.charCodeAt(0) + i - 1)
	}

	for(var j=1; j<rowNum+1; j++){
        gridHeader[j][0] = j
	}

	return gridHeader
}
/**
 * 表格渲染
 * @param sheet
 * @param sheetTable
 */
SheetRender.prototype.renderSheet = function(sheet, sheetTable) {
    for (var child = sheetTable.firstChild; child != null; child = sheetTable.firstChild) {
        sheetTable.removeChild(child);
    }
    var gridHeader = createHeader(sheet.rowNum, sheet.colNum)

    var rowNum = gridHeader.length
    var colNum = gridHeader[0].length

    var rowHead = document.createElement('tr')
    sheetTable.appendChild(rowHead)
    //渲染第一行表头
    for (var i = 0; i < colNum; i++) {
        var rowHeadTH = document.createElement('th')
        // rowHeadTH.style.height = '20px'
        rowHeadTH.id = gridHeader[0][i] + '_0'
        rowHead.appendChild(rowHeadTH)

        var rowHeadDiv = document.createElement('div')
        rowHeadTH.appendChild(rowHeadDiv)
        // 第一行第一列特殊处理
        if (i === 0) {
            //rowHeadTH.style.width = '29px'
        } else {
            rowHeadDiv.innerHTML = gridHeader[0][i]
            rowHeadDiv.style.width = config.CellConfig.width + 'px'
        }
    }

    var SheetEventBinderModule = require('SheetEventBinder')
    var SheetEventBinder = SheetEventBinderModule.SheetEventBinder
    var sheetEventBinder = new SheetEventBinder(sheet)

    //渲染除第一行外单元格
    for (var j = 1; j < rowNum; j++) {
        var rowTR = document.createElement('tr')
        sheetTable.appendChild(rowTR)

        for (var k = 0; k < colNum; k++) {
            //表格第一列特殊处理
            if (k === 0) {
                var rowTH = document.createElement('th')
                rowTH.id = '@_' + gridHeader[j][0]
                rowTR.appendChild(rowTH)
                var rowDiv = document.createElement('div')
                rowDiv.style.height = 25 + 'px'
                rowTH.appendChild(rowDiv)
                rowDiv.innerHTML = gridHeader[j][k]
            } else {
                var rowTD = document.createElement('td')
                rowTD.id = gridHeader[0][k] + "_" + j
                rowTR.appendChild(rowTD)
                if (!sheet.isPreview) {
                    sheetEventBinder.initRowTD(rowTD)
                }
                var rowDiv = document.createElement('div')
                rowTD.className = "noWrap"
                rowTD.appendChild(rowDiv)
            }
        }
    }
    if (!sheet.isPreview) {
        sheetTable.style.borderCollapse = 'collapse'
        sheetTable.style.border = '1px solid #ccc'
        var input = document.createElement('input')
        input.style = style.InputStyle
        input.id = config.InputConfig.id

        sheetEventBinder.initInput(input)


        sheetTable.appendChild(input)

        var clipBoard = document.createElement('textarea') // used for ctrl-c/ctrl-v where an invisible text area is needed
        clipBoard.style = style.ClipBoardStyle
        clipBoard.value = config.ClipBoardConfig.value;
        clipBoard.id = config.ClipBoardConfig.id;
        sheetTable.appendChild(clipBoard)

        // var span = document.createElement('span') // 用于获得字符串的显示长度
        // span.style = 'visibility: hidden;white-space: nowrap;filter:alpha(opacity=0);'
        // span.value = '';
        // span.id = 'sp'
        // sheetTable.appendChild(span)

        var multiLine = document.createElement('textarea')
        multiLine.style = 'position: absolute;left: 30px;top: 20px;display:none;'
        multiLine.style.left = screen.width / 2 - 150 + 'px'
        multiLine.style.top = screen.height / 2 - 300 + 'px'
        multiLine.style.width = 400 + 'px'
        multiLine.style.height = 400 + 'px'
        multiLine.value = '';
        multiLine.id = 'multiLine1'
        if (!config.WSConfig.isPreview) {
            sheetEventBinder.initMultiLine(multiLine)
        }
        sheetTable.appendChild(multiLine)
        var dragBar = document.createElement('div')
        dragBar.style = 'position: absolute;'
        dragBar.style.backgroundColor = 'yellow'
        dragBar.style.width = 8 + 'px'
        dragBar.style.height = 8 + 'px'
        dragBar.style.display = 'none'
        dragBar.id = 'dragBar'
        dragBar.webkitUserDrag = 'true'
        if (!config.WSConfig.isPreview) {
            sheetEventBinder.initDragBar(dragBar)
        }
        sheetTable.appendChild(dragBar)
    }

}

module.exports.SheetRender = SheetRender