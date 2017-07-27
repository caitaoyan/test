/**
 * 工作区对象，包括工具栏，表格等
 * @type {*}
 */
var config= require('config')

var Workspace = function (Tool, Sheet, SliderBar) {

    this.width = config.WSConfig.width
    this.height = config.WSConfig.height

    this.Tool =Tool
    this.Sheet= Sheet
    this.SliderBar = SliderBar
}

module.exports.Workspace = Workspace