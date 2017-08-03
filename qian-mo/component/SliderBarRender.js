/**
 * Created by Ian on 17/7/22.
 */

var style  = require('style')
var config = require('config')

var SliderBarRender = function (sheet, sheetDiv) {
	this.sheet = sheet
	this.sheetDiv = sheetDiv

	// alert(this.sheetDiv)
}

SliderBarRender.prototype.init = function (sliderBarDiv, sheetTableDiv) {

	var toggleDiv = document.createElement('div')
	toggleDiv.id = config.SlideBarConfig.toggleId
	toggleDiv.isOpen = false
	this.sheetDiv.appendChild(toggleDiv)

	renderBar(sliderBarDiv, sheetTableDiv, toggleDiv)
	renderToggle(toggleDiv)

	var sliderPaneConfig = config.SlideBarConfig.sliderPaneConfig
	var sliderPanes = {}

	for(var paneConfig in sliderPaneConfig){
		var paneDiv = document.createElement('div')
		paneDiv.style = style.PaneStyle

		paneDiv.id = sliderPaneConfig[paneConfig].paneId
		// paneDiv.config = paneConfig
		sliderPanes[paneConfig] = paneDiv

		sliderBarDiv.appendChild(paneDiv)
	}
	renderPanes(sliderPanes)

	// var isOpen = this.isOpen
	// var sheetDiv = this.sheetDiv

	toggleDiv.onclick = function () {
		toggleDiv.isOpen = !toggleDiv.isOpen

		renderBar(sliderBarDiv, sheetTableDiv, toggleDiv)
		renderToggle(toggleDiv)
	}
}

function renderToggle(toggleDiv) {

	var iconHtmlMap = config.SlideBarConfig.iconHtmlMap
	var toggleOpenIcon = config.SlideBarConfig.toggleOpenIcon
	var toggleCloseIcon = config.SlideBarConfig.toggleCloseIcon

	var isOpen = toggleDiv.isOpen
	// var toggleDiv = document.createElement('div')

	if(isOpen){

		toggleDiv.innerHTML = iconHtmlMap[toggleCloseIcon]
		toggleDiv.style = style.ToggleDivOpenLeft + style.ToggleDivStyle

		toggleDiv.onmouseover = function () {
			toggleDiv.style = style.ToggleDivOpenLeft + style.ToggleDivHoverStyle
		}

		toggleDiv.onmouseout = function () {
			toggleDiv.style = style.ToggleDivOpenLeft + style.ToggleDivStyle
			// toggleDiv.style.left = '76%'
		}
	}else{

		toggleDiv.innerHTML = iconHtmlMap[toggleOpenIcon]
		toggleDiv.style = style.ToggleDivCloseLeft + style.ToggleDivStyle

		toggleDiv.onmouseover = function () {
			toggleDiv.style = style.ToggleDivCloseLeft + style.ToggleDivHoverStyle
		}

		toggleDiv.onmouseout = function () {
			toggleDiv.style = style.ToggleDivCloseLeft + style.ToggleDivStyle
		}
	}

	// console.log(sheet)
	// sheetDiv.appendChild(toggleDiv)

	setTimeout(function () {
		toggleDiv.style.opacity = '0.5'
	}, 5000)
}

function renderBar(sliderBarDiv, sheetTableDiv, toggleDiv) {

	var isOpen = toggleDiv.isOpen

	sliderBarDiv.style = style.SliderBarStyle

	if(isOpen){
		sheetTableDiv.style.width = '75%'
		sliderBarDiv.style.width = '25%'
	}else{
		sheetTableDiv.style.width = '100%'
		sliderBarDiv.style.width = '0'
	}
}

function renderPanes(sliderPanes) {

	var sliderPaneConfig = config.SlideBarConfig.sliderPaneConfig

	for(var key in sliderPanes) {
		var isOpen = false
		var paneDiv = sliderPanes[key]
		var paneConfig = key

		var paneTitleDiv = document.createElement('div')
		paneTitleDiv.innerHTML = sliderPaneConfig[paneConfig].title
		paneTitleDiv.id = sliderPaneConfig[paneConfig].titleId
		paneTitleDiv.style = style.PaneTitleStyle
		paneTitleDiv.isOpen = isOpen

		var arrowDiv = document.createElement('div')
		arrowDiv.innerHTML = config.SlideBarConfig.iconHtmlMap[config.SlideBarConfig.arrowIcon]
		arrowDiv.id = sliderPaneConfig[paneConfig].arrowId
		paneTitleDiv.appendChild(arrowDiv)

		var paneContentDiv = document.createElement('div')
		paneContentDiv.id = sliderPaneConfig[paneConfig].contentId

		renderPane(arrowDiv, paneContentDiv, isOpen)

		//第一次写闭包！！！
		paneTitleDiv.onclick = function (arrowDiv, paneContentDiv, paneTitleDiv) {
			return function () {
				paneTitleDiv.isOpen = !paneTitleDiv.isOpen
				renderPane(arrowDiv, paneContentDiv, paneTitleDiv)
			}
		}(arrowDiv, paneContentDiv, paneTitleDiv)
		// paneContentDiv.style = style.PaneContentOpenStyle


		paneDiv.appendChild(paneTitleDiv)
		paneDiv.appendChild(paneContentDiv)
	}
}

function renderPane(arrowDiv, paneContentDiv, paneTitleDiv) {

	var isOpen = paneTitleDiv.isOpen

	if(isOpen){
		paneContentDiv.style = style.PaneContentOpenStyle
		arrowDiv.style = style.ArrowUpStyle

		if(paneContentDiv.firstChild){
			setTimeout(function () {paneContentDiv.firstChild.style.display = 'table'}, 200)
		}
	}else{
		paneContentDiv.style = style.PaneContentCloseStyle
		arrowDiv.style = style.ArrowDownStyle

		if(paneContentDiv.firstChild){paneContentDiv.firstChild.style.display = 'none'}
	}
}

// SliderBarRender.prototype
//
// function addFunc() {
//
// }

module.exports.SliderBarRender = SliderBarRender

module.exports.renderBar = renderBar

module.exports.renderPane = renderPane

module.exports.renderToggle = renderToggle