/**
 * 单元格对象
 */

var configModule = require('config')

var cellConfig = configModule.CellConfig

var Cell = function (coord) {

	//成员属性
	//设置默认值
	//this.autoLF = false

	//this.viewWidth = 0

	this.height = cellConfig.height
	this.width = cellConfig.width

	this.foregroundColor = ''
	this.backgroundColor = ''

	this.font = ''
	this.fontSize = ''
	// // this.fontWeight = false
	// // this.fontStyle = false

	this.topFrame = false
	this.bottomFrame = false
	this.leftFrame = false
	this.rightFrame = false

	this.alignment = 'left'
	this.wrapText = false

	this.coord = coord

	this.content = ''
	this.formula = ''

	this.show = true

	this.bold = false
	this.italic = false

	this.colSpan = 1
	this.rowSpan = 1
}

module.exports.Cell = Cell