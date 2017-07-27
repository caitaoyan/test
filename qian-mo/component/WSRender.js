/**
 * 工作区域渲染
 * @param wsManager
 * @param parentNode
 * @constructor
 */


var WSRender = function (wsManager, parentNode) {
	this.manager = wsManager
	this.parNode = parentNode
}

WSRender.prototype.init = function () {
	var ws = this.manager.workspace

	//workspace
	var WSDiv = document.createElement('div')
	// WSDiv.style.width = ws.width + 'px'
	// WSDiv.style.height = ws.height + 'px'
	WSDiv.style = require('style').WSStyle
	this.parNode.appendChild(WSDiv)

	//tool
	var ToolRenderModule = require('ToolRender')
	var ToolRender = ToolRenderModule.ToolRender
	var toolRender = new ToolRender(ws.Sheet)
	var tool = ws.Tool
	var toolDiv = document.createElement('div')
	toolRender.init(toolDiv, tool.width, tool.height)
	WSDiv.appendChild(toolDiv)

	//sheet
	var SheetRenderModule = require('SheetRender')
	var SheetRender = SheetRenderModule.SheetRender
	var sheetRender = new SheetRender(ws.Sheet)
	var sheet = ws.Sheet
	var sheetDiv = document.createElement('div')
	sheetRender.init(sheetDiv)
	WSDiv.appendChild(sheetDiv)

	//slideBar
	// var SliderBarRenderModule = require('SliderBarRender')
	// var SliderBarRender = SliderBarRenderModule.SliderBarRender
	// var sliderBarRender = new SliderBarRender(ws.Sheet)
	// var sliderBar = ws.SliderBar
	// var sliderBarDiv = document.createElement('div')
	// sliderBarRender.init(sliderBarDiv)
	// WSDiv.appendChild(sliderBarDiv)

}
module.exports.WSRender = WSRender