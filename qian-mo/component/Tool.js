/**
 * 工具栏对象
 */
var configModule = require('config')

var ToolConfig = configModule.ToolConfig

var Tool = function () {
    this.width = ToolConfig.width
    this.height = ToolConfig.height
}

module.exports.Tool = Tool