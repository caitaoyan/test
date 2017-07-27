/**
 * 工具栏渲染
 */
var config = require('config')
var style = require('style')
var ToolRender = function (sheet) {
    this.sheet = sheet
}
ToolRender.prototype.init = function (toolDiv, width, height) {
    var ToolEventBinderModule = require('ToolEventBinder')
    var ToolEventBinder = ToolEventBinderModule.ToolEventBinder
    // var toolEventBinder = new ToolEventBinder(this.sheet)

    toolDiv.style.width = width + 'px'
    toolDiv.style.height = height + 'px'

    renderMenu(toolDiv, this.sheet)

}

function renderMenu(toolDiv, sheet) {
    var menuDiv = document.createElement('div')
    var buttonBox = document.createElement('div')

    buttonBox.style = style.ButtonBoxStyle

    toolDiv.appendChild(menuDiv)
    toolDiv.appendChild(buttonBox)

    var toolMenu = config.ToolConfig.menu

    var styleBox = document.createElement('div')
    styleBox.innerHTML = toolMenu[0]
    var toolBox = document.createElement('div')
    toolBox.innerHTML = toolMenu[1]
    var dataBox = document.createElement('div')
    dataBox.innerHTML = toolMenu[2]

    var boxes = {
        styleBox: {node: styleBox, selected: false},
        toolBox: {node: toolBox, selected: false},
        dataBox: {node: dataBox, selected: false}
    }
    menuDiv.id = 'menuDiv'
    menuDiv.style = style.MeunDivStyle

    changeSelected(boxes, 'styleBox')
    renderTool('styleBox', buttonBox, sheet)

    styleBox.onclick = function () {
        if (!sheet.isPreview) {
            changeSelected(boxes, 'styleBox')
            renderTool('styleBox', buttonBox, sheet)
        }
    }
    toolBox.onclick = function () {
        if (!sheet.isPreview) {
            changeSelected(boxes, 'toolBox')
            renderTool('toolBox', buttonBox, sheet)
        }
    }
    dataBox.onclick = function () {
        if (!sheet.isPreview) {
            changeSelected(boxes, 'dataBox')
            renderTool('dataBox', buttonBox, sheet)
        }
    }

    for (var box in boxes) {
        menuDiv.appendChild(boxes[box].node)
    }

}

function renderTool(menu, buttonBox, sheet) {

    removeAllChild(buttonBox)

    var ToolEventBinderModule = require('ToolEventBinder')
    var ToolEventBinder = ToolEventBinderModule.ToolEventBinder
    var toolEventBinder = new ToolEventBinder(sheet)

    var cmdMap = config.ToolConfig.cmdCodeMap
    var toolHtml = config.ToolConfig.toolHtml
    var styleHtml = config.ToolConfig.styleHtml
    var dataHtml = config.ToolConfig.dataHtml
    var defaultHtml = config.ToolConfig.defaultHtml

    var previewHtml = config.ToolConfig.previewHtml
    var previewDiv = document.createElement('div')
    previewHtml.forEach(function (innerhtml) {

        var buttonDiv = document.createElement('div')
        buttonDiv.id = innerhtml
        buttonDiv.style = style.ButtonDivStyle
        buttonDiv.innerHTML = innerhtml

        buttonDiv.onclick = function () {
            toolEventBinder.buttonClick(innerhtml)
        }
        previewDiv.appendChild(buttonDiv)
    })
    previewDiv.style.display = 'none'
    previewDiv.id = 'previewDiv'
    buttonBox.appendChild(previewDiv)

    var editDiv = document.createElement('div')
    defaultHtml.forEach(function (innerhtml) {
        console.log(innerhtml)
        var buttonDiv = document.createElement('div')
        buttonDiv.id = innerhtml
        buttonDiv.style = style.ButtonDivStyle
        buttonDiv.innerHTML = cmdMap[innerhtml]
        // if(innerhtml==='Init'){
        //     buttonDiv.style.display='none'
        // }
        buttonDiv.onclick = function () {
            toolEventBinder.buttonClick(innerhtml)
        }
        editDiv.appendChild(buttonDiv)
    })

    if (menu === 'styleBox') {

        for (var i = 0; i < styleHtml.length; i++) {

            if (styleHtml[i] === 'font') {
                renderFont(toolEventBinder, editDiv)
                continue
            }

            if (styleHtml[i] === 'align') {
                renderAlign(toolEventBinder, editDiv)
                continue
            }

            if (styleHtml[i] === 'border') {
                renderBorder(toolEventBinder, editDiv)
                continue
            }

            var styleButtonDiv = document.createElement('div')
            styleButtonDiv.id = styleHtml[i]
            styleButtonDiv.style = style.ButtonDivStyle
            styleButtonDiv.innerHTML = cmdMap[styleHtml[i]]

            editDiv.appendChild(styleButtonDiv)

            if (styleHtml[i] === 'bold') {
                toolEventBinder.initFontWeight(styleButtonDiv)
                continue
            }
            if (styleHtml[i] === 'italic') {
                toolEventBinder.initFontStyle(styleButtonDiv)
                continue
            }
            if (styleHtml[i] === 'textColor') {
                renderFontColor(toolEventBinder, styleButtonDiv)
                renderColorSelect(toolEventBinder, editDiv)
                continue
            }
            if (styleHtml[i] === 'fillColor') {
                renderBackgroundColor(toolEventBinder, styleButtonDiv)
                renderColorSelect(toolEventBinder, editDiv)
                continue
            }


            styleButtonDiv.onclick = function (e) {
                toolEventBinder.buttonClick(e.currentTarget.id)
            }
        }
    } else if (menu === 'toolBox') {

        for (var m = 0; m < toolHtml.length; m++) {

            var toolButtonDiv = document.createElement('div')
            toolButtonDiv.id = toolHtml[m]
            toolButtonDiv.style = style.ButtonDivStyle
            toolButtonDiv.innerHTML = cmdMap[toolHtml[m]]

            editDiv.appendChild(toolButtonDiv)

            if (toolHtml[m] === 'mergeCell') {
                toolEventBinder.initMerge(toolButtonDiv, true)
                continue
            }

            if (toolHtml[m] === 'splitCell') {
                toolEventBinder.initMerge(toolButtonDiv, false)
                continue
            }

            toolButtonDiv.onclick = function (e) {
                toolEventBinder.buttonClick(e.currentTarget.id)
            }
        }
    } else {

        for (var n = 0; n < dataHtml.length; n++) {

            var dataButtonDiv = document.createElement('div')
            dataButtonDiv.id = dataHtml[n]
            dataButtonDiv.style = style.ButtonDivStyle
            dataButtonDiv.style.fontSize = '15px'
            dataButtonDiv.innerHTML = dataHtml[n]

            dataButtonDiv.onclick = function (e) {
                console.log(e.currentTarget.id)
                toolEventBinder.buttonClick(e.currentTarget.id)
            }
            editDiv.appendChild(dataButtonDiv)
        }
    }
    editDiv.id = 'editDiv'
    buttonBox.appendChild(editDiv)

}

function changeSelected(boxes, selectedNode) {

    for (var box in boxes) {
        boxes[box].selected = false
    }
    boxes[selectedNode].selected = true

    for (var b in boxes) {
        if (boxes[b].selected) {
            boxes[b].node.style = style.MeunBoxSelectedStyle
        } else {
            boxes[b].node.style = style.MeunBoxStyle
        }
    }
}

function renderFont(toolEventBinder, buttonBox) {
    //字体
    var fontFamilySelect = document.createElement('select')
    fontFamilySelect.style.marginLeft = '10px'
    var fontFamilyOption = ['fontFamily', 'Default', 'Custom', 'Verdana', 'Arial', 'Courier']
    fontFamilyOption.forEach(function (o) {
        var option = document.createElement('option')
        option.innerHTML = o
        fontFamilySelect.appendChild(option)
    })
    toolEventBinder.initFontFamily(fontFamilySelect)
    buttonBox.appendChild(fontFamilySelect)

    //字体大小
    var fontSizeSelect = document.createElement('select')
    fontSizeSelect.style.marginLeft = '10px'
    var fontSizeOption = ['fontSize', 'Default', 'X-Small', 'Small', 'Medium', 'Large',
        'X-Large', '6pt', '7pt', '8pt', '9pt', '10pt', '11pt', '12pt', '14pt', '16pt', '18pt',
        '20pt', '22pt', '24pt', '28pt', '36pt', '48pt', '72pt']
    fontSizeOption.forEach(function (o) {
        var option = document.createElement('option')
        option.innerHTML = o
        fontSizeSelect.appendChild(option)
    })
    toolEventBinder.initFontSize(fontSizeSelect)
    buttonBox.appendChild(fontSizeSelect)
}

function renderBorder(toolEventBinder, buttonBox) {
    var borderHtml = config.ToolConfig.borderHtml
    var cmdMap = config.ToolConfig.cmdCodeMap
    var borderButton = []
    var i

    for (i = 0; i < borderHtml.length; i++) {
        var borderButtonDiv = document.createElement('div')
        borderButtonDiv.id = borderHtml[i]
        borderButtonDiv.style = style.ButtonDivStyle
        borderButtonDiv.innerHTML = cmdMap[borderHtml[i]]

        buttonBox.appendChild(borderButtonDiv)
        borderButton.push(borderButtonDiv)
    }

    for (i = 0; i < borderButton.length; i++) {
        borderButton[i].onclick = function (e) {
            for (var j = 0; j < borderButton.length; j++) {
                borderButton[j].style = style.ButtonDivStyle
            }

            var fontBorderOption = config.ToolConfig.fontBorderOption
            toolEventBinder.initFontBorder(fontBorderOption[e.currentTarget.id])

            e.currentTarget.style = style.ButtonDivSelectedStyle
        }
    }
}

function renderAlign(toolEventBinder, buttonBox) {
    var alignHtml = config.ToolConfig.alignHtml
    var cmdMap = config.ToolConfig.cmdCodeMap
    var alignButton = []
    var i

    for (i = 0; i < alignHtml.length; i++) {
        var alignButtonDiv = document.createElement('div')
        alignButtonDiv.id = alignHtml[i]
        alignButtonDiv.style = style.ButtonDivStyle
        alignButtonDiv.innerHTML = cmdMap[alignHtml[i]]

        buttonBox.appendChild(alignButtonDiv)
        alignButton.push(alignButtonDiv)
    }

    for (i = 0; i < alignButton.length; i++) {
        alignButton[i].onclick = function (e) {
            for (var j = 0; j < alignButton.length; j++) {
                alignButton[j].style = style.ButtonDivStyle
            }

            var alignOption = config.ToolConfig.alignOption
            toolEventBinder.initTextAlign(alignOption[e.currentTarget.id])

            e.currentTarget.style = style.ButtonDivSelectedStyle
        }
    }
}

function renderFontColor(toolEventBinder, styleButtonDiv) {
    //color,字体颜色
    // var colorDiv = document.createElement('div')
    // colorDiv.style = style.ColorDivStyle
    // buttonBox.appendChild(colorDiv)
    toolEventBinder.initColor(styleButtonDiv)
}

function renderBackgroundColor(toolEventBinder, styleButtonDiv) {

    //backgroundColor背景颜色
    // var backgroundColorDiv = document.createElement('div')
    // backgroundColorDiv.style = style.backgroundColorDivStyle
    // buttonBox.appendChild(backgroundColorDiv)
    toolEventBinder.initBackgroundColor(styleButtonDiv)
}

function renderColorSelect(toolEventBinder, buttonBox) {
    var colorSelectDiv = document.createElement("div")
    colorSelectDiv.id = 'colorSelect'
    colorSelectDiv.style = style.ColorSelectDivStyle
    // colorSelectDiv.style.display = 'none'
    // colorSelectDiv.style.position = "absolute"
    // colorSelectDiv.style.zIndex = 100
    // colorSelectDiv.style.backgroundColor = "#FFF"
    // colorSelectDiv.style.border = "1px solid black"
    // colorSelectDiv.style.width = '106px'
    buttonBox.appendChild(colorSelectDiv)

    var mainDiv = document.createElement("div")
    // mainDiv.style.padding = "3px"
    // mainDiv.style.backgroundColor = "#CCC"
    colorSelectDiv.appendChild(mainDiv)

    var mainTable = document.createElement("table")
    mainTable.cellSpacing = 0
    mainTable.cellPadding = 0
    mainTable.style.width = "100px"
    mainDiv.appendChild(mainTable)

    var mainTbody = document.createElement("tbody")
    mainTable.appendChild(mainTbody)
    var steps = [0, 68, 153, 204, 255]
    var commonRgb = ["400", "310", "420", "440", "442", "340", "040", "042", "032", "044", "024", "004", "204", "314", "402", "414"]

    for (var row = 0; row < 16; row++) {
        var tr = document.createElement("tr")
        for (var col = 0; col < 5; col++) {
            var td = document.createElement("td")
            toolEventBinder.initColorSelect(td)
            td.style.fontSize = "1px"
            td.innerHTML = "&nbsp;"
            td.style.height = "10px"
            td.style.lineHeight = '12px'
            if (col <= 1) {
                td.style.width = "17px"
                td.style.borderRight = "3px solid white"
            }
            else {
                td.style.width = "20px"
                td.style.backgroundRepeat = "no-repeat"
            }
            if (col === 0) {
                var x = commonRgb[row]
                td.style.backgroundColor = "rgb(" + steps[x.charAt(0) - 0] + "," + steps[x.charAt(1) - 0] + "," + steps[x.charAt(2) - 0] + ")"
            }
            if (col === 1) {
                td.style.backgroundColor = makeRGB(17 * (15 - row), 17 * (15 - row), 17 * (15 - row))
            }
            if (col === 2) {
                td.style.backgroundColor = makeRGB(17 * (15 - row), 0, 0)
            }
            if (col === 3) {
                td.style.backgroundColor = makeRGB(0, 17 * (15 - row), 0)
            }
            if (col === 4) {
                td.style.backgroundColor = makeRGB(0, 0, 17 * (15 - row))
            }
            tr.appendChild(td)
        }
        mainTbody.appendChild(tr)
    }
    mainTable.appendChild(mainTbody)
}

function makeRGB(r, g, b) {
    return "rgb(" + (r > 0 ? r : 0) + "," + (g > 0 ? g : 0) + "," + (b > 0 ? b : 0) + ")"
}

function removeAllChild(node) {

    while (node.hasChildNodes()) {
        node.removeChild(node.firstChild)
    }
}

module.exports.ToolRender = ToolRender