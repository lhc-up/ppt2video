# ppt2video

ppt(x)转视频，目前仅支持windows

ms office ppt(x)转图片

ffmpeg 图片转视频并添加音频，拼接视频并添加转场动画

## Install

```bash
git clone https://github.com/luohao8023/ppt2video.git

cd ppt2video

npm i
```

## Quick Start

```js
const fs = require('fs');
const path = require('path');
const { PPT2video} = require('./src/ppt2pdf.js');

// 查看当前ppt软件版本
PPT2video.getPPTVersion().then(version => {
    console.log(version);
}).catch(err => {
    // 获取版本失败，未安装或其他错误
});

// 支持的转场动画列表
console.log(PPT2video.transitions);

// 实例化
const converter = new PPT2video({
    pptPath: path.resolve('../doc/test.pptx'),
    animate: {
        use: true,//使用转场动画
        // type: 'vertclose',具体的动画类型
        // type: 'random',随机使用一种动画
        // type: 'randomAll',随机所有动画
        // type: ['fade', 'vertclose'],按顺序使用给定数组中的动画
        type: 'inturn',//按顺序使用所有动画
        duration: 1//动画时长
    },
    tmpdir: path.resolve('./temp'),//默认是os.tmpdir()
    resultFolder: path.resolve('./result'),
    audioFolder: path.resolve('../doc/audio'),
    slideDuration: 3//没有对应音频时，默认每页展示时长
});

// ppt转图片
(async () => {
    await converter.getSlideImgs(pptPath, resultFolder);
})();

// 单个图片转视频
(async () => {
    await converter.img2video(imgPath, audioPath, resultFolder, (percent, curSeconds, totalSeconds) => {});
})();

// 图片批量转视频
(async () => {
    await converter.igetSlideVideos(imgFolder, audioFolder, resultFolder);
})();

// 获取音频时长
(async () => {
    await converter.getDuration(audioPath);
})();

// 拼接视频
(async () => {
    await converter.concatVideos(videoFolder, resultFolder, options, (percent, curSeconds, totalSeconds) => {});
})();

// 查看filter_complex_script
converter.createFilterScript(videoCount, {});

// ppt -> video，一步到位
(async () => {
    const resultPath = await converter.convert();
})();
```

## 环境要求

### windows  
- 安装ms Powerpoint 或 wps Powerpoint
- 安装ffmpeg
- vbs调用Powerpoint(ms/wps)程序，分页转为图片
- ffmpeg把单个图片转为视频并添加音频，无音频时添加静音音轨
- ffmpeg合并多个视频片段，支持31种ffmpeg原生转场动画


### Linux、Mac
* TODO：ppt转图片过程使用Libreoffice