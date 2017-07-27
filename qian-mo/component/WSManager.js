/**
 * 整体工作管理对象
 * @constructor
 */
var config = require('config')
var WSManager = function () {
}

WSManager.prototype.init = function (parentNode) {

	//实例化初始化Tool对象
	var ToolModule = require('Tool')
	var Tool = ToolModule.Tool
	var tool = new Tool()

	//实例化初始化UndoStack对象
	var UndoStackModule = require('UndoStack')
	var UndoStack = UndoStackModule.UndoStack
	var undoStack = new UndoStack()

	//实例化初始化Sheet对象
	var SheetModule = require('Sheet')
	var Sheet = SheetModule.Sheet
	var sheet = new Sheet(undoStack)

	// //实例化初始化SliderBar对象
	// var SliderBarModule = require('component/SliderBarHandler')

	//实例化初始化Workspace对象
	var WorkspaceModule = require('Workspace')
	var Workspace = WorkspaceModule.Workspace
	this.workspace = new Workspace(tool, sheet)

	//实例化初始化WSRender对象
	var WSRenderModule = require('WSRender')
	var WSRender = WSRenderModule.WSRender
	var wsRender = new WSRender(this, parentNode)

	wsRender.init()

	if (config.WSConfig.isInit && parentNode.getAttribute("url")) {
		document.getElementById("Init").onclick()
		config.WSConfig.isInit = false
	}
}

module.exports.WSManager = WSManager