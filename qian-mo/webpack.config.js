/**
 * Created by Ian on 17/6/5.
 */
var path = require('path')
var packageJSON = require('./package.json');


module.exports = {
	entry: "./app.js",
	output: {
		path: path.join(__dirname, 'target', 'classes', 'META-INF', 'resources', 'webjars', packageJSON.name, packageJSON.version),
    publicPath: '/assets/',
		filename: "app.js"
	},
	resolve: {
		alias: {
            config: path.resolve(__dirname, "./config/config.js"),
			style: path.resolve(__dirname, "./style/style.js"),
            Cell: path.resolve(__dirname, "./component/Cell.js"),
            CellRender: path.resolve(__dirname, "./component/CellRender.js"),
            Sheet: path.resolve(__dirname, "./component/Sheet.js"),
            SheetRender: path.resolve(__dirname, "./component/SheetRender.js"),
            SheetEventHandler: path.resolve(__dirname, "./component/SheetEventHandler.js"),
            SheetEventBinder: path.resolve(__dirname, "./component/SheetEventBinder.js"),
            Tool: path.resolve(__dirname, "./component/Tool.js"),
            ToolRender: path.resolve(__dirname, "./component/ToolRender.js"),
            ToolEventHandler: path.resolve(__dirname, "./component/ToolEventHandler.js"),
            ToolEventBinder: path.resolve(__dirname, "./component/ToolEventBinder.js"),
			UndoStack: path.resolve(__dirname, "./component/UndoStack.js"),
            Workspace: path.resolve(__dirname, "./component/Workspace.js"),
            WSManager: path.resolve(__dirname, "./component/WSManager.js"),
            WSRender: path.resolve(__dirname, "./component/WSRender.js"),
            Util: path.resolve(__dirname, "./component/Util.js"),
			SliderBarHandler: path.resolve(__dirname, "./component/SliderBarHandler.js"),
			SliderBarRender: path.resolve(__dirname, "./component/SliderBarRender.js")
		}
	},
	devServer: {
		historyApiFallback: {
			rewrites: [
				{ from: /./, to: './app.html' }
			]
		},
		noInfo: false
	},
	devtool: "source-map"
}