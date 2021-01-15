const { PPT2video } = require('../ppt2pdf.js');
const path = require('path');
const ppt2video = new PPT2video({
    pptPath: path.resolve('../doc/test.pptx'),
    animate: {
        use: true,
        // type: 'vertclose',
        type: 'inturn',
        duration: 1
    },
    tmpdir: path.resolve('./temp'),
    resultFolder: path.resolve('./result'),
    audioFolder: path.resolve('../doc/audio'),
    slideDuration: 3
});

const imgFolder = path.resolve(__dirname, '../doc/img');
const audioFolder = path.resolve(__dirname, '../doc/audio');
const videoFolder = path.resolve(__dirname, '../doc/video');
// ppt2video.getSlideVideos(imgFolder, audioFolder, videoFolder).then(data => {
//     console.log(data)
// }).catch(err => {
//     console.log(err);
// });

ppt2video.concatVideos(videoFolder, path.dirname(videoFolder), {
    resultName: 'hhhhhhhresult.mp4'
}, (a, b, c) =>{
    console.log(a, b, c);
}).then(data => {
    console.log(data)
}).catch(err => {
    console.log(err);
});