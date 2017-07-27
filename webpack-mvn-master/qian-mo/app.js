/**
 * Created by Ian on 17/6/5.
 */

var WSManagerModule = require('WSManager')
var WSManager = WSManagerModule.WSManager
var wsManager=new WSManager()

var parentNode = document.getElementById('QianMoApp')

var myUrl=GetQueryString("url")
if(myUrl!=null && myUrl.toString().length>1){
    parentNode.setAttribute('url',myUrl)
}
function GetQueryString(name){
    var reg=new RegExp("(^|&)"+name+"=([^&]*)(&|$)")
    var r=window.location.search.substring(1).match(reg)
    if(r!=null){
        return unescape(r[2])
    }
    return null
}

wsManager.init(parentNode)