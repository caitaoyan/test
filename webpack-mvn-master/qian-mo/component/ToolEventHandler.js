/**
 * 工具栏事件
 */
var config = require('config')


var SheetRenderModule = require('SheetRender')
var SheetRender = SheetRenderModule.SheetRender
var sheetRender = null
// var CellRenderModule = require('CellRender')
// var CellRender = CellRenderModule.CellRender
// var cellRender = null

var SheetEventHandlerModule = require('SheetEventHandler')
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null

var ToolEventHandler = function (sheet) {
    this.sheet = sheet
    sheetRender = new SheetRender(sheet)
    //cellRender = new CellRender(sheet)
    sheetEventHandler = new SheetEventHandler(sheet)
}
ToolEventHandler.prototype.buttonClick = function (action, value) {
    var sheet = this.sheet
    switch (action) {
        case "wrapText":
            var lastCellId = this.sheet.lastCellid
            if (!sheet.cells[lastCellId]) {
                sheet.setAttr('wrapText', true)
            }
            else {
                sheet.setAttr('wrapText', !this.sheet.cells[lastCellId].wrapText)
            }
            sheet.render()
            break
        case "preview":
            var e = this.sheet.cells
            if (JSON.stringify(e) !== "{}") {
                var needEditCells = []
                for (var key in e) {
                    if (isNaN(e[key]["content"]) && e[key]["content"] != '' && e[key]["content"].indexOf('=') == 0) {
                        var formula = {
                            'coord': key,
                            'value': e[key]["content"]
                        }
                        needEditCells.push(formula)
                        this.sheet.cells[key]["formula"] = e[key]["content"]
                    }
                    // if(e[key]["formula"]!==undefined && e[key]["formula"]!=''){
                    //     var formula = {
                    //         'coord': key,
                    //         'value': e[key]["formula"]
                    //     }
                    //     needEditCells.push(formula)
                    // }
                }
                if (needEditCells.length > 0) {
                    var ajax
                    if (window.XMLHttpRequest) {
                        ajax = new XMLHttpRequest()
                    } else if (window.ActiveXObject) {
                        ajax = new window.ActiveXObject()
                    } else {
                        alert("请升级至最新版本的浏览器")
                    }
                    if (ajax !== null) {
                        ajax.open("POST", "/qianmo-service/changeContent", true)
                        // ajax.open("POST", "http://localhost:8088/qianmo-service/changeContent", true)
                        needEditCells = JSON.stringify(needEditCells)
                        ajax.send(needEditCells)
                        ajax.onreadystatechange = function () {
                            if (ajax.readyState == 4 && ajax.status == 200) {
                                var CellList = JSON.parse(ajax.responseText)
                                for (var key in CellList) {
                                    var value = CellList[key].value.substring(1)
                                    if (value.indexOf("+") != -1) {
                                        var result = 0
                                        var values = value.split("+")
                                        for (var v in values) {
                                            result += parseInt(values[v])
                                        }
                                        value = result
                                    }
                                    e[CellList[key].coord]["content"] = value
                                }
                                sheet.isPreview = true
                                document.getElementById('previewDiv').style.display = 'block'
                                document.getElementById('editDiv').style.display = 'none'
                                document.getElementById('menuDiv').isDisabled = true
                                sheetRender.renderSheet(sheet,document.getElementById('sheetTable'))

                                for (cell in sheet.cells){
                                    for (attr in sheet.cells[cell]){
                                        if(attr != 'coord')sheet.setIdAttr(sheet.cells[cell].coord,attr,sheet.cells[cell][attr])
                                        //console.log(attr)
                                    }
                                }
                            }
                        }
                    }
                }else{
                    sheet.isPreview = true
                    document.getElementById('previewDiv').style.display = 'block'
                    document.getElementById('editDiv').style.display = 'none'
                    document.getElementById('menuDiv').isDisabled = true
                    sheetRender.renderSheet(sheet,document.getElementById('sheetTable'))

                    for (cell in sheet.cells){
                        for (attr in sheet.cells[cell]){
                            if(attr != 'coord')sheet.setIdAttr(sheet.cells[cell].coord,attr,sheet.cells[cell][attr])
                        }
                    }
                }
            } else {
                alert("无内容，不允许预览！")
            }

            break
        case "Edit":
            sheet.isPreview = false
            document.getElementById('previewDiv').style.display = 'none'
            document.getElementById('editDiv').style.display = 'block'
            //document.getElementById('menuDiv').style.display = 'block'
            sheetRender.renderSheet(sheet,document.getElementById('sheetTable'))
            // console.log(sheet.cells)
            for (cell in sheet.cells){
                for (attr in sheet.cells[cell]){
                    if(attr != 'coord'){
                        // console.log(sheet.cells[cell]['formula'])
                        if(attr==='content' && sheet.cells[cell]['formula']!==undefined&& sheet.cells[cell]['formula']!==''){
                            sheet.setIdAttr(sheet.cells[cell].coord,attr,sheet.cells[cell]['formula'])
                        }else{
                            sheet.setIdAttr(sheet.cells[cell].coord,attr,sheet.cells[cell][attr])
                        }
                    }
                }
            }
            break
        case "Down":
            var myMask=document.getElementById('mask')
            myMask.style.display="block"
            var e = this.sheet.cells
            var a = {}
            var CellList = []
            for (var key in e) {
                if(e[key].show){
                    var col=key.split('_')[0]
                    var row=key.split('_')[1]
                    col=String.fromCharCode(col.charCodeAt(0)+e[key].colSpan-1)
                    row=parseInt(row)+e[key].rowSpan-1
                    var area=key+'_'+col+'_'+row
                    var content=e[key]["content"] === undefined ? "" : e[key]["content"]=== null ?"":e[key]["content"]=== ""?"":e[key]["content"]+''
                    var addCell = {
                        "cellName": key,
                        "area": area,
                        "content":content.replace(/&nbsp;/g,''),
                        "format": e[key]["format"] === undefined ? "" : e[key]["format"],
                        "font": e[key]["font"] === undefined ? "" : e[key]["font"],
                        "fontSize": e[key]["fontSize"] === undefined ? "" : e[key]["fontSize"],
                        "foregroundColor": e[key]["foregroundColor"] === undefined ? "" : e[key]["foregroundColor"],
                        "backgroundColor": e[key]["backgroundColor"] === undefined ? "" : e[key]["backgroundColor"],
                        "formula": e[key]["formula"] === undefined ? "" : e[key]["formula"],
                        "leftFrame": e[key]["leftFrame"] === undefined ? "" : e[key]["leftFrame"],
                        "topFrame": e[key]["topFrame"] === undefined ? "" : e[key]["topFrame"],
                        "rightFrame": e[key]["rightFrame"] === undefined ? "" : e[key]["rightFrame"],
                        "bottomFrame": e[key]["bottomFrame"] === undefined ? "" : e[key]["bottomFrame"],
                        "indentation": e[key]["indentation"] === undefined ? 0 : e[key]["indentation"],
                        "alignment": e[key]["alignment"] === undefined ? "" : e[key]["alignment"],
                        "bold": e[key]["bold"] === undefined ? "" : e[key]["bold"],
                        "italic": e[key]["italic"] === undefined ? "" : e[key]["italic"],
                        "vertical":e[key]["vertical"] === undefined ? "" : e[key]["vertical"],
                        "wrapText":e[key]["wrapText"] === undefined ? "" : e[key]["wrapText"],
                        "width":e[key]["width"] === undefined ? "" : e[key]["width"],
                        "height":e[key]["height"] === undefined ? "" : e[key]["height"]
                    }
                    // console.log(addCell.content)
                    CellList.push(addCell)
                }
            }
            var json= JSON.stringify(CellList)
            var ajax
            if (window.XMLHttpRequest) {
                ajax = new XMLHttpRequest()
            } else if (window.ActiveXObject) {
                ajax = new window.ActiveXObject()
            } else {
                alert("请升级至最新版本的浏览器")
            }
            if(ajax !=null){
                ajax.open("POST","/qianmo-service/excelDownload",true)
                // ajax.open("POST","http://localhost:8088/qianmo-service/excelDownload",true)
                ajax.onload=function(){
                    if(ajax.status==200){
                        var filename = "";
                        var disposition = ajax.getResponseHeader('Content-Disposition');
                        if (disposition && disposition.indexOf('attachment') !== -1) {
                            var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                            var matches = filenameRegex.exec(disposition);
                            if (matches != null && matches[1]) filename = matches[1].replace(/['"]/g, '');
                        }
                        var type = ajax.getResponseHeader('Content-Type');
                        var blob = new Blob([this.response], { type: type });
                        if (typeof window.navigator.msSaveBlob !== 'undefined') {
                            // IE workaround for "HTML7007: One or more blob URLs were revoked by closing the blob for which they were created. These URLs will no longer resolve as the data backing the URL has been freed."
                            window.navigator.msSaveBlob(blob, filename);
                        } else {
                            var URL = window.URL || window.webkitURL;
                            var downloadUrl = URL.createObjectURL(blob);
                            if (filename) {
                                // use HTML5 a[download] attribute to specify filename
                                var a = document.createElement("a");
                                // safari doesn't support this yet
                                if (typeof a.download === 'undefined') {
                                    window.location = downloadUrl;
                                } else {
                                    a.href = downloadUrl;
                                    a.download = filename;
                                    document.body.appendChild(a);
                                    a.click();
                                }
                            } else {
                                window.location = downloadUrl;
                            }

                            setTimeout(function () { URL.revokeObjectURL(downloadUrl); }, 100); // cleanup
                        }
                    }
                    myMask.style.display="none"
                }
                ajax.responseType = 'blob'
                ajax.setRequestHeader('Content-type', 'application/json;charset=utf-8')
                ajax.send(json)
            }
            break
        case  "Init":
            var parentNode = document.getElementById("QianMoApp")
            var param = parentNode.getAttribute("url")
            var url = ''
            var method="GET"
            if(!param&& param!=''){
                url='json.json'
                param=null
            }else{
                method="POST"
                url='/qianmo-service/getContentJson'
                // url='http://localhost:8088/qianmo-service/getContentJson'
                param=param.substring(param.lastIndexOf("\\")+1)
            }
            var ajax
            if (window.XMLHttpRequest) {
                ajax = new XMLHttpRequest()
            } else if (window.ActiveXObject) {
                ajax = new window.ActiveXObject()
            } else {
                alert("请升级至最新版本的浏览器")
            }
            if (ajax !== null) {
                ajax.open(method, url, true)
                ajax.send(param)
                ajax.onreadystatechange = function () {
                    if (ajax.readyState === 4 && ajax.status === 200) {
                        var CellList = JSON.parse(ajax.responseText)
                        CellList.forEach(function (e) {

                            var area = e['area']
                            var firstCell = e['cellName'].split('_')

                            var cells = {}
                            cells.firstCellCol = firstCell[0].charCodeAt(0)
                            cells.firstCellRow = parseInt(firstCell[1])
                            var areaSubstring = area.split('_')
                            cells.lastCellCol = areaSubstring[2].charCodeAt(0)
                            cells.lastCellRow = parseInt(areaSubstring[3])
                            if (cells.firstCellCol !== cells.lastCellCol || cells.firstCellRow !== cells.lastCellRow) {
                                sheet.setAttr('merge', true, cells)
                            }


                            cells.lastCellCol = firstCell[0].charCodeAt(0)
                            cells.lastCellRow = parseInt(firstCell[1])


                            for (var key in e) {
                                switch (e[key]) {
                                    case 'false':
                                        var value = false
                                        break
                                    case 'true':
                                        var value = true
                                        break
                                    default:
                                        var value = e[key]
                                }
                                if ((key === 'bottomFrame' || key === 'topFrame' || key === 'leftFrame' ||
                                        key === 'rightFrame') && value === false) continue
                                // if(key==='content' && e['formula']!==undefined && e['formula']!==''){
                                //     value=e['formula']
                                // }
                                sheet.setAttr(key, value, cells)
                            }
                        })

                        sheet.render()
                    }
                }
            }
            break
        case 'copy':
            cells = sheet.editCells
            sheet.firstCopiedCell = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
            sheet.copiedCells = []
            for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var id = String.fromCharCode(i) + '_' + j
                    var newCell = {}
                    newCell.coord = id
                    sheet.cellAttrs.forEach(function (attr) {
                        if (!sheet.cells[id]) {
                            newCell[attr[0]] = attr[1]
                        }
                        else {
                            newCell[attr[0]] = sheet.cells[id][attr[0]]
                        }
                    })
                    if (!sheet.cells[id]) {
                        newCell.bottomFrame = false
                        newCell.topFrame = false
                        newCell.leftFrame = false
                        newCell.rightFrame = false
                    }
                    else {
                        newCell.bottomFrame = sheet.cells[id].bottomFrame
                        newCell.topFrame = sheet.cells[id].topFrame
                        newCell.leftFrame = sheet.cells[id].leftFrame
                        newCell.rightFrame = sheet.cells[id].rightFrame
                    }
                    sheet.copiedCells.push(newCell)
                }
            }
            console.log('copy done!')
            break
        case 'cut':
            cells = sheet.editCells
            sheet.firstCopiedCell = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
            sheet.copiedCells = []
            for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var id = String.fromCharCode(i) + '_' + j
                    var newCell = {}
                    newCell.coord = id
                    sheet.cellAttrs.forEach(function (attr) {
                        if (!sheet.cells[id]) {
                            newCell[attr[0]] = attr[1]
                        }
                        else {
                            newCell[attr[0]] = sheet.cells[id][attr[0]]
                        }
                    })
                    if (!sheet.cells[id]) {
                        newCell.bottomFrame = false
                        newCell.topFrame = false
                        newCell.leftFrame = false
                        newCell.rightFrame = false
                    }
                    else {
                        newCell.bottomFrame = sheet.cells[id].bottomFrame
                        newCell.topFrame = sheet.cells[id].topFrame
                        newCell.leftFrame = sheet.cells[id].leftFrame
                        newCell.rightFrame = sheet.cells[id].rightFrame
                    }
                    sheet.copiedCells.push(newCell)
                }
            }
            sheet.cellAttrs.forEach(function (attr) {
                sheet.setAttr(attr[0], attr[1], cells)
            })
            sheet.setAttr('allFrame', false, cells)
            sheet.render()
            sheet.lastCellid = sheet.firstCellid

            console.log('cut done!')
            break
        case 'paste':
            var borders = [['bottomFrame', false], ['leftFrame', false], ['topFrame', false], ['rightFrame', false]]
            if (sheet.firstCopiedCell) {
                cells = sheet.editCells
                sheet.copiedCells.forEach(function (cell) {
                    var newCell = {}
                    var firstCopiedCellCol = sheet.firstCopiedCell.split('_')[0].charCodeAt(0)
                    var firstCopiedCellRow = parseInt(sheet.firstCopiedCell.split('_')[1])
                    cellCol = cell.coord.split('_')[0].charCodeAt(0)
                    cellRow = parseInt(cell.coord.split('_')[1])
                    newCell.firstCellCol = cells.firstCellCol + cellCol - firstCopiedCellCol
                    newCell.firstCellRow = cells.firstCellRow + cellRow - firstCopiedCellRow
                    newCell.lastCellCol = newCell.firstCellCol
                    newCell.lastCellRow = newCell.firstCellRow
                    sheet.cellAttrs.forEach(function (attr) {
                        sheet.setAttr(attr[0], cell[attr[0]], newCell)
                    })
                    borders.forEach(function (frame) {
                        sheet.setAttr(frame[0], cell[frame[0]], newCell)
                    })
                    delete newCell
                })
                sheet.render()
                console.log('paste done!')
                // console.log(sheet.cells)
            } else {
                console.log('nothing to paste')
            }
            break
        case 'font':
            sheet.setAttr('font', value)
            sheet.render()
            break
        case 'fontSize':
            sheet.setAttr('fontSize', value)
            sheet.render()
            break
        case 'bold':
            var lastCellId = sheet.lastCell.id
            if (!sheet.cells[lastCellId]) {
                sheet.setAttr('bold', true)
            }
            sheet.setAttr('bold', !sheet.cells[lastCellId].bold)
            sheet.render()
            break
        case 'italic':
            var lastCellId = sheet.lastCell.id
            if (!sheet.cells[lastCellId]) {
                sheet.setAttr('italic', true)
            }
            sheet.setAttr('italic', !this.sheet.cells[lastCellId].italic)
            sheet.render()
            break
        case 'merge':
            this.setCellBackgroundColor('#fff')
            this.sheet.setAttr('merge', value)
            this.sheet.render()
            this.mouseDown(sheet.firstCellid)
            // console.log(sheet.firstCellid)
            this.mouseUp(document.getElementById(sheet.lastCellid))
            break
        // case 'border':
        // switch(value){
        // 	case 'left':
        // 		sheet.setAttr('leftFrame',false)
        //             console.log('a')
        // 		break
        // 	case 'top':
        // 		sheet.setAttr('topFrame',false)
        // 		break
        // 	case 'right':
        // 		sheet.setAttr('rightFrame',false)
        // 		break
        // 	case 'bottom':
        // 		sheet.setAttr('bottomFrame',false)
        // 		break
        // 	case 'clear':
        // 		sheet.setAttr('allFrame',false)
        // 		break
        // 	case 'outer':
        // 		sheet.setAttr('leftFrame',true)
        // 		sheet.setAttr('topFrame',true)
        // 		sheet.setAttr('rightFrame',true)
        // 		sheet.setAttr('bottomFrame',true)
        // 		break
        // 	case 'all':
        // 		sheet.setAttr('allFrame',true)
        // 		break
        //
        // 	//fontBorderSelect.value='border'
        // }
        // sheet.render()
        //     break
        case 'alignment':
            sheet.setAttr('alignment', value)
            sheet.render()
            break
        case 'undo':
            sheet.undo()
            break
        case 'redo':
            sheet.redo()
            break
        // case 'sort':
        //     break
        case 'multiLine':
            document.getElementById('dragBar').style.display = 'none'
            sheet.isMultiLineEditing = true
            var ml = document.getElementById('multiLine1')
            ml.style.display = 'block'
            if (sheet.cells[sheet.lastCellid]) ml.value = sheet.cells[sheet.lastCellid].content
            else ml.value = ''
            ml.focus()
            break
        case 'cleanText':
            sheet.setAttr('content', '')
            sheet.render()
            break
        case 'cleanStyle':
            sheet.cellAttrs.forEach(function (attr) {
                if (attr[0] !== 'content') sheet.setAttr(attr[0], attr[1])
            })
            sheet.render()
            break
        case 'cleanAll':
            sheet.cellAttrs.forEach(function (attr) {
                sheet.setAttr(attr[0], attr[1])
            })
            sheet.setAttr('allFrameEx', false)
            sheet.render()
            break
        case 'addCol':
        case 'deleteCol':
            cells = sheet.getColAndRow(sheet.firstCellid.split('_')[0]+'_1',String.fromCharCode(sheet.maxCol)+'_'+sheet.maxRow)
            //sheet.firstCopiedCell = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
            var sheetcopiedCells = []
            for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var id = String.fromCharCode(i) + '_' + j
                    var newCell = {}
                    newCell.coord = id
                    sheet.cellAttrs.forEach(function (attr) {
                        if (!sheet.cells[id]) {
                            newCell[attr[0]] = attr[1]
                        }
                        else {
                            newCell[attr[0]] = sheet.cells[id][attr[0]]
                        }
                    })
                    if (!sheet.cells[id]) {
                        newCell.bottomFrame = false
                        newCell.topFrame = false
                        newCell.leftFrame = false
                        newCell.rightFrame = false
                    }
                    else {
                        newCell.bottomFrame = sheet.cells[id].bottomFrame
                        newCell.topFrame = sheet.cells[id].topFrame
                        newCell.leftFrame = sheet.cells[id].leftFrame
                        newCell.rightFrame = sheet.cells[id].rightFrame
                    }
                    sheetcopiedCells.push(newCell)
                }
            }
            sheet.cellAttrs.forEach(function (attr) {
                sheet.setAttr(attr[0], attr[1], cells)
            })
            sheet.setAttr('allFrame', false, cells)

            var borders = [['bottomFrame', false], ['leftFrame', false], ['topFrame', false], ['rightFrame', false]]
            if (sheetcopiedCells) {
                if(action === 'deleteCol'){
                    var firstCopiedCellCol = cells.firstCellCol+1
                    var firstCopiedCellRow = cells.firstCellRow
                }else{
                    var firstCopiedCellCol = cells.firstCellCol-1
                    var firstCopiedCellRow = cells.firstCellRow
                }
                sheetcopiedCells.forEach(function (cell) {
                    var newCell = {}
                    cellCol = cell.coord.split('_')[0].charCodeAt(0)
                    cellRow = parseInt(cell.coord.split('_')[1])
                    newCell.firstCellCol = cells.firstCellCol + cellCol - firstCopiedCellCol
                    newCell.firstCellRow = cells.firstCellRow + cellRow - firstCopiedCellRow
                    newCell.lastCellCol = newCell.firstCellCol
                    newCell.lastCellRow = newCell.firstCellRow
                    sheet.cellAttrs.forEach(function (attr) {
                        sheet.setAttr(attr[0], cell[attr[0]], newCell)
                    })
                    borders.forEach(function (frame) {
                        sheet.setAttr(frame[0], cell[frame[0]], newCell)
                    })
                    delete newCell
                })
                sheet.render()
                console.log('addCol')
            }
            break
        case 'addRow':
        case 'deleteRow':
            cells = sheet.getColAndRow('A_'+sheet.firstCellid.split('_')[1],String.fromCharCode(sheet.maxCol)+'_'+sheet.maxRow)
            //sheet.firstCopiedCell = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
            var sheetcopiedCells = []
            for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var id = String.fromCharCode(i) + '_' + j
                    var newCell = {}
                    newCell.coord = id
                    sheet.cellAttrs.forEach(function (attr) {
                        if (!sheet.cells[id]) {
                            newCell[attr[0]] = attr[1]
                        }
                        else {
                            newCell[attr[0]] = sheet.cells[id][attr[0]]
                        }
                    })
                    if (!sheet.cells[id]) {
                        newCell.bottomFrame = false
                        newCell.topFrame = false
                        newCell.leftFrame = false
                        newCell.rightFrame = false
                    }
                    else {
                        newCell.bottomFrame = sheet.cells[id].bottomFrame
                        newCell.topFrame = sheet.cells[id].topFrame
                        newCell.leftFrame = sheet.cells[id].leftFrame
                        newCell.rightFrame = sheet.cells[id].rightFrame
                    }
                    sheetcopiedCells.push(newCell)
                }
            }
            sheet.cellAttrs.forEach(function (attr) {
                sheet.setAttr(attr[0], attr[1], cells)
            })
            sheet.setAttr('allFrame', false, cells)

            var borders = [['bottomFrame', false], ['leftFrame', false], ['topFrame', false], ['rightFrame', false]]
            if (sheetcopiedCells) {
                if(action === 'deleteRow'){
                    var firstCopiedCellCol = cells.firstCellCol
                    var firstCopiedCellRow = cells.firstCellRow+1
                }else{
                    var firstCopiedCellCol = cells.firstCellCol
                    var firstCopiedCellRow = cells.firstCellRow-1
                }
                sheetcopiedCells.forEach(function (cell) {
                    var newCell = {}
                    cellCol = cell.coord.split('_')[0].charCodeAt(0)
                    cellRow = parseInt(cell.coord.split('_')[1])
                    newCell.firstCellCol = cells.firstCellCol + cellCol - firstCopiedCellCol
                    newCell.firstCellRow = cells.firstCellRow + cellRow - firstCopiedCellRow
                    newCell.lastCellCol = newCell.firstCellCol
                    newCell.lastCellRow = newCell.firstCellRow
                    sheet.cellAttrs.forEach(function (attr) {
                        sheet.setAttr(attr[0], cell[attr[0]], newCell)
                    })
                    borders.forEach(function (frame) {
                        sheet.setAttr(frame[0], cell[frame[0]], newCell)
                    })
                    delete newCell
                })
                sheet.render()
            }
            break
        default:
            alert('该功能未实现，请期待')
            break
    }
    if(action !== ('multiLine' || 'copy' ) && !sheet.isPreview){
        sheetEventHandler.mouseDown(sheet.firstCellid)
        sheetEventHandler.mouseUp(document.getElementById(sheet.lastCellid))
    }
}


module.exports.ToolEventHandler = ToolEventHandler