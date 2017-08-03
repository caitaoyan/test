var Formula = require('Formula').Formula

var FuncListRender =  function () {
    this.formula = new Formula()
}

FuncListRender.prototype.init = function (ele) {
    var formula = this.formula
    var funcTable = document.createElement('table')
    funcTable.style.height = 200 + 'px'
    funcTable.style.width = 500 + 'px'
    funcTable.border = '1px bold #000'
    funcTable.style.backgroundColor = '#DDD'
    ele.appendChild(funcTable)

    var funcTr = document.createElement('tr')
    var funcTdLeft = document.createElement('td')
    var funcTdRight = document.createElement('td')
    var funcSpanLeft = document.createElement('span')
    var funcSpanRight = document.createElement('span')
    var funcTr2 = document.createElement('tr')
    var funcTd2 = document.createElement('td')
    var funcDiv = document.createElement('div')
    var funcDiv2 = document.createElement('div')
    var funcClassSelect = document.createElement('select')
    var funcNameSelect = document.createElement('select')
    funcNameSelect.id = 'funcNameSelect'

    var sumOfClass = 0
    for (name in formula.FunctionClassList) {
        var op = document.createElement('option')
        op.innerHTML = name
        op.id = name
        funcClassSelect.appendChild(op)
        if (sumOfClass === 0) op.selected = true
        sumOfClass++

    }

    funcClassSelect.size = sumOfClass
    funcNameSelect.size = sumOfClass
    var sumOfFname = 0
    for (fname in  formula.FunctionClassList[funcClassSelect.value]) {
        var op = document.createElement('option')
        op.innerHTML = fname
        op.id = fname
        funcNameSelect.appendChild(op)
        if (sumOfFname === 0) op.selected = true
        sumOfFname++

    }
    funcDiv.innerHTML = funcNameSelect.value + '('+ formula.funcInfo['s_farg_' + formula.funcParem[funcNameSelect.value]] + ')'
    funcDiv2.innerHTML = formula.funcInfo['s_fdef_' + funcNameSelect.value]

    funcSpanLeft.innerHTML = 'Category'
    funcTdRight.innerHTML = 'Functions'
    funcTd2.colSpan = 2
    funcClassSelect.style.width = 250 + 'px'
    funcNameSelect.style.width = 250 + 'px'
    funcTr.appendChild(funcTdLeft)
    funcTr.appendChild(funcTdRight)
    funcTdLeft.appendChild(funcSpanLeft)
    funcTdLeft.appendChild(document.createElement('tr'))
    funcTdLeft.appendChild(funcClassSelect)
    funcTdRight.appendChild(funcSpanRight)
    funcTdRight.appendChild(document.createElement('tr'))
    funcTdRight.appendChild(funcNameSelect)
    funcTr2.appendChild(funcTd2)
    funcTd2.appendChild(funcDiv)
    funcTd2.appendChild(document.createElement('tr'))
    funcTd2.appendChild(funcDiv2)

    funcTable.appendChild(funcTr)
    funcTable.appendChild(funcTr2)

    funcClassSelect.onchange = function () {
        funcNameSelect.innerHTML = ''
        funcNameSelect.size = 7
        var sumOfFname =0
        for (fname in  formula.FunctionClassList[this.value]) {
            var op = document.createElement('option')
            op.innerHTML = fname
            op.id = fname
            funcNameSelect.appendChild(op)
            if (sumOfFname === 0) {
                funcDiv.innerHTML = fname + '('+ formula.funcInfo['s_farg_' + formula.funcParem[fname]]+ ')'
                funcDiv2.innerHTML = formula.funcInfo['s_fdef_' + fname]
                op.selected = true
            }
        }
    }
    funcNameSelect.onchange = function () {
        funcDiv.innerHTML = funcNameSelect.value +'('+ formula.funcInfo['s_farg_' + formula.funcParem[funcNameSelect.value]] + ')'
        funcDiv2.innerHTML = formula.funcInfo['s_fdef_' + funcNameSelect.value]
    }
}

module.exports.FuncListRender = FuncListRender