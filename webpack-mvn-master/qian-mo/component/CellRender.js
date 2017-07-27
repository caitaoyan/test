var config = require('config')

// var UtilModule=require('Util')
// var Util=UtilModule.Util
// var util=new Util()

var CellModule = require('Cell')
var Cell = CellModule.Cell

var CellRender = function (sheet) {
    this.sheet = sheet
}
/**
 * 渲染单元格
 * @param id
 * @param cmd
 * @param value
 */
CellRender.prototype.renderCell = function (id, cmd, value) {
    // var sheet=this.sheet
    // if(id && id.indexOf('_')==-1){
    //     var index=id.search(/\d/)
    //     id=id.substring(0,index)+'_'+id.substring(index,id.length)
    // }
    var ele = document.getElementById(id)

    // if(id && !sheet.cells[id]){
    //     sheet.cells[id] = new Cell(id)
    // }
    if (ele) {
        switch (cmd) {
            //合并范围
            case "colSpan":
                if(value>0) ele.colSpan = value
                break
            case 'rowSpan':
                if(value>0)ele.rowSpan = value
                break
            case "show":
                if (value) ele.style.display = 'table-cell'
                else {
                    ele.style.display = 'none'
                }
                break
            //字体颜色
            case "foregroundColor":
                ele.style.color = value


                break
            //背景色
            case "backgroundColor":

                ele.style.backgroundColor = value

                break
            //文本内容
            case "content":
                if (value === null) {
                    value = ''
                }
                ele.firstChild.innerHTML = value
                break
            //字体
            case "font":
                if (value == 'Default') {
                    value = ''
                }
                ele.style.fontFamily = value
                break
            //字体大小
            case "fontSize":
                if (value == 'Default' || value == '-----') {
                    value = ''
                }
                ele.style.fontSize = value
                //this.sheet.cells[id].fontSize=value
                break
            //字体类型
            //case "format":
            //    this.sheet.cells[id].format=value
            //    break
            //加粗
            case "bold":
                if (value) ele.style.fontWeight = 'bold'
                else ele.style.fontWeight = ''
                break

            //斜体
            case "italic":
                if (value) ele.style.fontStyle = 'italic'
                else ele.style.fontStyle = ''
                break
            //
            //对齐方向
            case "alignment":

                ele.style.textAlign = value
                ele.firstChild.style.textAlign = value

                break
            //换行
            case "wrapText":
                if (value) ele.firstChild.className = 'wrap'
                else ele.firstChild.className = 'noWrap'
                break
            //边框
            case "bottomFrame":
                var newId = id.split('_')[0] + '_' + (parseInt(id.split('_')[1])+1)
                if (value ) {
                    ele.style.borderBottom = '1px solid #000'
                } else if(!this.sheet.cells[newId] || !this.sheet.cells[newId].topFrame){
                    ele.style.borderBottom = ''
                    // console.log(newId)
                }
                // if (value) ele.style.borderBottom = '2px solid #000'
                // else ele.style.borderBottom = ''
                break
            case "leftFrame":
                var newId = String.fromCharCode(id.split('_')[0].charCodeAt(0) - 1) + '_' + id.split('_')[1]
                if (value) {
                    document.getElementById(newId).style.borderRight = '1px solid #000'
                } else if(!this.sheet.cells[newId] || !this.sheet.cells[newId].rightFrame)document.getElementById(newId).style.borderRight = ''
                // if (value) ele.style.borderLeft = '2px solid #000'
                // else ele.style.borderLeft = ''
                break
            case "rightFrame":
                var newId = String.fromCharCode(id.split('_')[0].charCodeAt(0) + 1) + '_' + id.split('_')[1]
                if (value) {
                    ele.style.borderRight = '1px solid #000'
                } else if(!this.sheet.cells[newId] || !this.sheet.cells[newId].leftFrame) ele.style.borderRight = ''
                // if (value) ele.style.borderRight = '2px solid #000'
                // else ele.style.borderRight = ''
                break
            case "topFrame":
                var newId = id.split('_')[0] + '_' + (parseInt(id.split('_')[1])-1)
                if (value) {
                    document.getElementById(newId).style.borderBottom = '1px solid #000'
                } else if(!this.sheet.cells[newId] || !this.sheet.cells[newId].bottomFrame) document.getElementById(newId).style.borderBottom = ''
                // if (value) ele.style.borderTop = '2px solid #000'
                // else ele.style.borderTop = ''
                break
            //宽高
            case 'width':
                if (value) {
                    var newId = id.split('_')[0]+'_0'
                    document.getElementById(newId).firstChild.style.width = parseInt(value) + 'px'
                }
                break
            case 'height':
                if (value){
                    var newId = '@_'+id.split('_')[1]
                    document.getElementById(newId).firstChild.style.height = parseInt(value) + 'px'
                }
                break
            case 'area':
                console.log('!!!')
                break
            // case "formula":
            //     this.sheet.cells[id].formula=value
            //     break
        }
    }
}

module.exports.CellRender = CellRender