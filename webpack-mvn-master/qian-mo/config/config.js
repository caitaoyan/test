/**
 * 配置文件，主要是个各类属性默认值
 * 整张表命名为：WorkSpace
 * 表格上半部分叫做：Tool
 * 表格叫做：Sheet
 *
 */

/**
 * 整张表-WorkSpace的属性默认值
 *
 */
var WSConfig = {
    isInit: true
}

/**
 * 表头-Tool的属性默认值
 *
 */
var ToolConfig = {
    height: '12%',
    width: '100%',

    cmdCodeMap: {
        borderAll: '&#xe228', borderBottom: '&#xe229', borderClear: '&#xe22a',
        borderLeft: '&#xe22e', borderOuter: '&#xe22f', borderRight: '&#xe230',
        borderTop: "&#xe232", mergeCell: '&#xe252', splitCell: '&#xe0b6',
        copy: '&#xe14d', cut: '&#xe14e', paste: '&#xe14f', find: '&#xe881',
        font: '&#xe167', alignCenter: '&#xe234', alignLeft: '&#xe236', alignRight: '&#xe237',
        bold: '&#xe238', italic: '&#xe23f', fillColor: '&#xe23a', textColor: '&#xe23c',
        brush: '&#xe243', underline: '&#xe249', sigma: '&#xe24a', func: '&#xe940',
        preview: '&#xe8f4', sort: '&#xe164', wrapText: '&#xe25b', undo: '&#xe967', redo: '&#xe968',
        cleanText: '清除文字', cleanStyle: '清除格式', cleanAll: '清除文字和格式', multiLine: '多行输入',
        addCol: '插列', addRow: '插列', menu: '&#xe3c7', close: '&#xe5cd',Init:'Init'
    },
    styleHtml: ['font', 'bold', 'italic', 'textColor', 'fillColor', 'align', 'border', 'brush', 'cleanText',
        'cleanStyle', 'cleanAll'],
    toolHtml: ['copy', 'paste', 'cut', 'sort', 'find', 'wrapText', 'func', 'sigma', 'mergeCell', 'splitCell',
        'multiLine', 'addCol','addRow'],
    dataHtml: ['数据源', '数据集'],
    defaultHtml: ['undo', 'redo', 'preview', 'Init'],
    borderHtml: ['borderAll', 'borderBottom', 'borderTop', 'borderClear', 'borderLeft', 'borderOuter', 'borderRight'],
    alignHtml: ['alignLeft', 'alignCenter', 'alignRight'],
    previewHtml: ['Edit', 'Down'],
    fontBorderOption: {
        borderLeft: 'left',
        borderTop: 'top',
        borderRight: 'right',
        borderBottom: 'bottom',
        borderClear: 'clear',
        borderOuter: 'outer',
        borderAll: 'all'
    },
    alignOption: {
        alignLeft: 'left',
        alignCenter: 'center',
        alignRight: 'right'
    },

    menu: ['样式', '工具', '数据']
}

/**
 * 侧边栏-sliderBar的常量
 *
 * */
var SlideBarConfig = {
    toggleOpenIcon: 'open',
    toggleCloseIcon: 'close',
	arrowIcon: 'arrow',

    id: 'sliderBarDiv',

    iconHtmlMap: {
		open: '&#xe3c7',
        close: '&#xe5cd',
        arrow: '&#xe313'
    },

    sliderPaneTitle: ['单元格属性','表格属性', '公式','数据源','数据集'],
    sliderPaneId: ['cellAttrPane', 'sheetAttrPane','funcPane', 'dataSrcPane', 'dataTarPane'],
    sliderPaneConfig:{
        cellAttr: {
            title: '单元格属性',
            paneId: 'cellAttrPane',
            titleId: 'cellAttrTitle',
            arrowId: 'cellAttrArrow',
            contentId: 'cellAttrContent'
        },
		sheetAttr: {
			title: '表格属性',
			paneId: 'sheetAttrPane',
			titleId: 'sheetAttrTitle',
			arrowId: 'sheetAttrArrow',
			contentId: 'sheetAttrContent'
		},
		func: {
			title: '公式',
			paneId: 'funcPane',
			titleId: 'funcTitle',
			arrowId: 'funcArrow',
			contentId: 'funcContent'
		},
		dataSrc: {
			title: '数据源',
			paneId: 'dataSrcPane',
			titleId: 'dataSrcTitle',
			arrowId: 'dataSrcArrow',
			contentId: 'dataSrcContent'
		},
		dataTar: {
			title: '数据集',
			paneId: 'dataTarPane',
			titleId: 'dataTarTitle',
			arrowId: 'dataTarArrow',
			contentId: 'dataTarContent'
		}
    },

    toggleId: 'toggleDiv'
}

/**
 * 表单-Cell的属性默认值
 *
 */
var CellConfig = {
    height: 20,
    width: 100
}

/**
 * 表单-Sheet的属性默认值
 *
 */
var SheetConfig = {
    headWidth: 30,
    headHeight: 25,

    rowNum: 200,
    colNum: 25
}

var SheetTableDivConfig = {
    id: 'sheetTableDiv'
}

var CellPropConfig = {
    id: 'cellProp'
}

var InputConfig = {
    id: 'input'
}

var ClipBoardConfig = {
    id: 'ta',
    value: ''
}
var keyboardTables = {

    specialKeysCommon: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
        35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
        46: "[del]", 113: "[f2]"
    },

    specialKeysIE: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
        35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
        46: "[del]", 113: "[f2]"
    },

    controlKeysIE: {
        67: "[ctrl-c]",
        83: "[ctrl-s]",
        86: "[ctrl-v]",
        88: "[ctrl-x]",
        90: "[ctrl-z]"
    },

    specialKeysOpera: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
        35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]",
        45: "[ins]", // issues with releases before 9.5 - same as "-" ("-" changed in 9.5)
        46: "[del]", // issues with releases before 9.5 - same as "." ("." changed in 9.5)
        113: "[f2]"
    },

    controlKeysOpera: {
        67: "[ctrl-c]",
        83: "[ctrl-s]",
        86: "[ctrl-v]",
        88: "[ctrl-x]",
        90: "[ctrl-z]"
    },

    specialKeysSafari: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 63232: "[aup]", 63233: "[adown]",
        63234: "[aleft]", 63235: "[aright]", 63272: "[del]", 63273: "[home]", 63275: "[end]", 63276: "[pgup]",
        63277: "[pgdn]", 63237: "[f2]"
    },

    controlKeysSafari: {
        99: "[ctrl-c]",
        115: "[ctrl-s]",
        118: "[ctrl-v]",
        120: "[ctrl-x]",
        122: "[ctrl-z]"
    },

    ignoreKeysSafari: {
        63236: "[f1]", 63238: "[f3]", 63239: "[f4]", 63240: "[f5]", 63241: "[f6]", 63242: "[f7]",
        63243: "[f8]", 63244: "[f9]", 63245: "[f10]", 63246: "[f11]", 63247: "[f12]", 63289: "[numlock]"
    },

    specialKeysFirefox: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
        35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
        46: "[del]", 113: "[f2]"
    },

    controlKeysFirefox: {
        99: "[ctrl-c]",
        115: "[ctrl-s]",
        118: "[ctrl-v]",
        120: "[ctrl-x]",
        122: "[ctrl-z]"
    },

    ignoreKeysFirefox: {
        16: "[shift]", 17: "[ctrl]", 18: "[alt]", 20: "[capslock]", 19: "[pause]", 44: "[printscreen]",
        91: "[windows]", 92: "[windows]", 112: "[f1]", 114: "[f3]", 115: "[f4]", 116: "[f5]",
        117: "[f6]", 118: "[f7]", 119: "[f8]", 120: "[f9]", 121: "[f10]", 122: "[f11]", 123: "[f12]",
        144: "[numlock]", 145: "[scrolllock]", 224: "[cmd]"
    }
}
/**
 * 公式类型
 * @type {[*]}
 */
// var function_classlist= ["all", "stat", "lookup", "datetime", "financial", "test", "math", "text"]
/**
 * 输出模块
 *
 */
module.exports.WSConfig = WSConfig
module.exports.ToolConfig = ToolConfig
module.exports.SheetConfig = SheetConfig
module.exports.CellConfig = CellConfig
module.exports.CellPropConfig = CellPropConfig
module.exports.InputConfig = InputConfig
module.exports.ClipBoardConfig = ClipBoardConfig
module.exports.SlideBarConfig = SlideBarConfig
module.exports.SheetTableDivConfig = SheetTableDivConfig
