/**
 *  堆栈对象
 */
var UndoStack = function () {

    this.tos = -1
    //this.maxRedo = 0
    this.maxUndo = 50
    this.stack = []

}
UndoStack.prototype.pushStep = function (stack) {
    while(this.stack.length > this.tos+1) this.stack.pop()
    this.stack.push(stack)
    if(this.stack.length > this.maxUndo) this.stack.shift()
    else this.tos++
}
UndoStack.prototype.reDo = function () {
    if (this.tos < this.stack.length-1){
        this.tos ++
        var result = this.stack[this.tos]
    }
    else result = null
    return result
}
UndoStack.prototype.unDo = function () {
    if (this.tos > -1){
        var result = this.stack[this.tos]
        this.tos --
    }
    else result = null
    return result
}
module.exports.UndoStack = UndoStack