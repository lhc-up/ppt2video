// 转场动画类型
const transitions = [
    'fade'       , 
    'wipeleft'    ,
    'wiperight'   ,
    'wipeup'      ,
    'wipedown'    ,
    'slideleft'   ,
    'slideright'  ,
    'slideup'     ,
    'slidedown'   ,
    'circlecrop'  ,
    'rectcrop'    ,
    'distance'    ,
    'fadeblack'   ,
    'fadewhite'   ,
    'radial'      ,
    'smoothleft'  ,
    'smoothright' ,
    'smoothup'    ,
    'smoothdown'  ,
    'circleopen'  ,
    'circleclose' ,
    'vertopen'    ,
    'vertclose'   ,
    'horzopen'    ,
    'horzclose'   ,
    'dissolve'    ,
    'pixelize'    ,
    'diagtl'      ,
    'diagtr'      ,
    'diagbl'      ,
    'diagbr'      ,
];

/**
 * 获取转场动画类型
 * @param {String} type 动画类型，不存在时取第一个
 * @param {Number} index 序号，超出时从头开始，负数倒序
 * @param {Array} typeArr 传入的类型组合，优先使用传入的类型
 * @returns {String} 动画类型
 */
function getTransitionType(type, index, typeArr=[]) {
    typeArr = typeArr.length ? typeArr : transitions;
    typeArr = typeArr.filter(v => transitions.includes(v));
    const len = typeArr.length;
    if (!!type) {
        typeArr = transitions;
        return typeArr.includes(type) ? type : typeArr[0];
    }
    if (Number.isInteger(index)) {
        typeArr = transitions;
        if (index >= 0) {
            return index < len 
            ? typeArr[index] 
            : getTransitionType(type, index - len);
        }
        return getTransitionType(type, index + len);
    }
    const random = Math.floor(Math.random() * len);
    return typeArr[random];
}

module.exports = {
    transitions,
    getTransitionType
}