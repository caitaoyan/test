/**
 * Created by Ian on 17/7/22.
 */


var config = require('config')
var style = require('style')
var renderPane = require('SliderBarRender').renderPane
var renderBar = require('SliderBarRender').renderBar
var renderToggle = require('SliderBarRender').renderToggle

var SliderBarHandler = function (sheet) {
	this.sheet = sheet
}

SliderBarHandler.prototype.autoOpen = function () {
	var sheet = this.sheet
	var cell = sheet.cells[sheet.firstCellid]

	if ((sheet.firstCellid !== sheet.lastCellid) || !cell) {
		return
	}

	var sliderBarDiv = document.getElementById(config.SlideBarConfig.id)
	var sheetTableDiv = document.getElementById(config.SheetTableDivConfig.id)
	var toggleDiv = document.getElementById(config.SlideBarConfig.toggleId)

	toggleDiv.isOpen = true

	renderBar(sliderBarDiv, sheetTableDiv, toggleDiv)
	renderToggle(toggleDiv)
}

SliderBarHandler.prototype.addCellAttr = function () {
	var sheet = this.sheet
	var cell = sheet.cells[sheet.firstCellid]

	if ((sheet.firstCellid !== sheet.lastCellid) || !cell) {
		return
	}

	// var paneDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.cellAttr.paneId)
	var titleDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.cellAttr.titleId)
	var arrowDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.cellAttr.arrowId)
	var contentDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.cellAttr.contentId)

	titleDiv.isOpen = true

	renderPane(arrowDiv, contentDiv, titleDiv)

	removeAllChild(contentDiv)

	var cellAttrTable = document.createElement('table')
	cellAttrTable.style = style.SliderTableStyle
	contentDiv.appendChild(cellAttrTable)

	var i = 0

	for (var attr in cell) {

		if(cell[attr] == false){continue}

		var attrTr = document.createElement('tr')
		cellAttrTable.appendChild(attrTr)

		i++
		if(i%2 === 1){
			attrTr.style = style.SliderEvenTrStyle
		}else{
			attrTr.style = style.SliderOddTrStyle
		}

		var attrTd = document.createElement('td')
		attrTd.innerHTML = attr
		attrTd.style = style.SliderTdStyle
		var valueTd = document.createElement('td')
		valueTd.innerHTML = cell[attr]
		valueTd.style = style.SliderTdStyle
		attrTr.appendChild(attrTd)
		attrTr.appendChild(valueTd)
	}
}

SliderBarHandler.prototype.addSheetAttr = function () {
	var sheet = this.sheet
	var contentDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.sheetAttr.contentId)


	console.log(sheet)
}

function removeAllChild(node) {

	while (node.hasChildNodes()) {
		node.removeChild(node.firstChild)
	}
}

module.exports.SliderBarHandler = SliderBarHandler