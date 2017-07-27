/**
 * Created by Ian on 17/6/8.
 */
/**
 * 样式常量类
 *
 * 存储各种默认样式
 * 以及后续要添加的主题样式
 *
 * 注意，大多数宽与高，是成员属性，写在config里面
 */

var config = require('config')
var WSConfig = config.WSConfig
var SheetConfig = config.SheetConfig
var ToolConfig = config.ToolConfig

var WSStyle =
	'position: absolute;' +
	'left: 0;' +
	'top: 0;' +
	'width:100%;' +
	'height:100%;' +
	'overflow: hidden'

var SheetDivStyle =
	'position: relative;' +
	'top: 0;' +
	'left: 0;' +
	'width: 100%;' +
	'height: 87%'

var SheetTableStyle = 'user-select: none;'

var SheetTableDivStyle =
	'display: inline-block;' +
	'height:100%;' +
	'width: 100%;' +
	'overflow: auto;' +
	'border: none;' +
	'transition: all 0.5s ease'

var SliderBarStyle =
	'display: inline-block;' +
	'position: absolute;' +
	'right :0;' +
	'top: 0;' +
	'width: 0;' +
	'height: 100%;' +
	'background-color: gray;' +
	'transition: all 0.5s ease;' +
	'overflow: auto;'

var ToggleDivCloseLeft = 'left: 95%;'
var ToggleDivOpenLeft = 'left: 76%;'

var ToggleDivStyle =
	'position: fixed;' +
	'bottom: 20px;' +
	'background-color: gray;' +
	'width: 40px;' +
	'height: 40px;' +
	'border-radius: 20px;' +
	'color: white;' +
	'line-height: 40px;' +
	'text-align: center;' +
	'z-index: 100;' +
	'user-select: none;' +
	'cursor: pointer;' +
	'transition: all 0.5s ease;'

var ToggleDivHoverStyle =
	'position: fixed;' +
	'bottom: 20px;' +
	'background-color: white;' +
	'width: 40px;' +
	'height: 40px;' +
	'border-radius: 20px;' +
	'color: gray;' +
	'line-height: 40px;' +
	'text-align: center;' +
	'z-index: 100;' +
	'user-select: none;' +
	'cursor: pointer;' +
	'transition: all 0.5s ease;'

var PaneStyle =
	'padding: 10px;' +
	'top: 0;' +
	'left: 0;'

var PaneTitleStyle =
	'border-bottom: solid 1px rgba(255,255,255,0.3);' +
	'color: rgba(255,255,255,0.9);' +
	'margin-bottom: 10px;' +
	'font-size: 0.9em;' +
	'user-select: none;' +
	'cursor: pointer;'

var PaneContentCloseStyle =
	'height: 0;' +
	'width: 100%;' +
	'min-height: 0px;' +
	'background-color: rgba(255,255,255,0.1);' +
	'border-radius: 8px;' +
	'transition: all 0.5s ease;'

var PaneContentOpenStyle =
	'width: 100%;' +
	'height: auto;' +
	'min-height: 200px;' +
	'background-color: rgba(255,255,255,0.1);' +
	'border-radius: 8px;' +
	'transition: all 0.5s ease;'

var ArrowDownStyle =
	'display: inline-block;' +
	'right: 0;' +
	'transition: all 0.5s ease;'

var ArrowUpStyle =
	'display: inline-block;' +
	'right: 0;' +
	'transform: rotate(180deg);' +
	'transition: all 0.5s ease'

var SliderTableStyle =
	'width: 100%;' +
	'border: none;'

var SliderOddTrStyle =
	'border: none;' +
	'background-color: rgba(0,0,0,0.1);'

var SliderEvenTrStyle=
	'border: none;' +
	'background-color: rgba(0,0,0,0);'

var  SliderTdStyle =
	'width: 50%;' +
	'border: none;' +
	'padding-left: 5px;' +
	'color: whiteSmoke;' +
	'font-size: 0.9em;' +
	'line-height: 1.8em'

var InputStyle =
	'display: none;' +
	'position: absolute'

var ClipBoardStyle =
	'display:none;' +
	'position:absolute;' +
	'height:1px;' +
	'width:1px;' +
	'opacity:0;' +
	'filter:alpha(opacity=0);'

var SpanStyle =
	''

var ToolStyle = ''

var MenuDivStyle =
	'height:37%; ' +
	'width: 100%;' +
	'line-height: 30px;' +
	'min-height: 28px'

var MenuBoxStyle =
	'display: inline-block; ' +
	'height: 100%; ' +
	'letter-spacing: 7px; ' +
	'padding-left: 7px; ' +
	'cursor: pointer; ' +
	'transition: all 0.1s'+
	'background-color: transparent; ' +
	'border-bottom: none;'

var MenuBoxSelectedStyle =
	'display: inline-block; ' +
	'height: 100%; ' +
	'letter-spacing: 10px; ' +
	'padding-left: 10px; ' +
	'cursor: pointer; ' +
	'transition: all 0.1s; ' +
	'background-color: #ddd; ' +
	'border-bottom: solid 2px grey;'

var ButtonBoxStyle =
	'height: 63%;' +
	'width: 100%;' +
	'background-color: #ddd;' +
	'line-height: 50px' +
	'min-height: 50px;'

var ButtonDivStyle =
	'display: inline-block;' +
	'height: 50px;' +
	'cursor: pointer;' +
	'transition: all 0.1s;' +
	'color: #333;' +
	'line-height: 50px;' +
	'margin-left: 20px;' +
	'font-size: 20px'

var ButtonDivSelectedStyle =
	'display: inline-block;' +
	'height: 50px;' +
	'cursor: pointer;' +
	'transition: all 0.1s;' +
	'color: #fff;' +
	'line-height: 50px;' +
	'margin-left: 20px;' +
	'font-size: 20px'

var CellStyle = ''

var ColorDivStyle =
	'display: inline; ' +
	'margin-left: 10px; ' +
	'padding-left: 15px; ' +
	'cursor: pointer; ' +
	'border: none; ' +
	'width: 15px; ' +
	'height: 15px; ' +
	'background-color: #fff;'

var backgroundColorDivStyle =
	'display: inline; ' +
	'margin-left: 10px; ' +
	'padding-left: 15px; ' +
	'cursor: pointer; ' +
	'border: none; ' +
	'width: 15px; ' +
	'height: 15px; ' +
	'background-color: #fff;'

var ColorSelectDivStyle =
	'display: none;' +
	'position: absolute;' +
	'z-index: 100;' +
	'background-color: #fff;' +
	'border: 1px solid black;' +
	'width: 106px;'

module.exports.WSStyle = WSStyle

module.exports.SheetDivStyle = SheetDivStyle

module.exports.SheetTableStyle = SheetTableStyle

module.exports.SheetTableDivStyle = SheetTableDivStyle

module.exports.ToggleDivOpenLeft = ToggleDivOpenLeft
module.exports.ToggleDivCloseLeft = ToggleDivCloseLeft
module.exports.ToggleDivStyle =  ToggleDivStyle
module.exports.ToggleDivHoverStyle = ToggleDivHoverStyle

module.exports.PaneStyle = PaneStyle
module.exports.PaneTitleStyle = PaneTitleStyle

module.exports.PaneContentCloseStyle = PaneContentCloseStyle
module.exports.PaneContentOpenStyle = PaneContentOpenStyle

module.exports.ArrowDownStyle = ArrowDownStyle
module.exports.ArrowUpStyle = ArrowUpStyle

module.exports.SliderTableStyle = SliderTableStyle
module.exports.SliderOddTrStyle = SliderOddTrStyle
module.exports.SliderEvenTrStyle = SliderEvenTrStyle
module.exports.SliderTdStyle = SliderTdStyle

module.exports.InputStyle = InputStyle

module.exports.ToolStyle = ToolStyle

module.exports.MeunDivStyle = MenuDivStyle

module.exports.MeunBoxStyle = MenuBoxStyle

module.exports.MeunBoxSelectedStyle = MenuBoxSelectedStyle

module.exports.CellStyle = CellStyle

module.exports.ClipBoardStyle = ClipBoardStyle

module.exports.ButtonBoxStyle = ButtonBoxStyle

module.exports.ButtonDivStyle = ButtonDivStyle

module.exports.ButtonDivSelectedStyle = ButtonDivSelectedStyle

module.exports.SpanStyle = SpanStyle

module.exports.ColorDivStyle = ColorDivStyle

module.exports.backgroundColorDivStyle = backgroundColorDivStyle

module.exports.ColorSelectDivStyle = ColorSelectDivStyle

module.exports.SliderBarStyle = SliderBarStyle