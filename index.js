const { PPT2video } = require('./src/ppt2pdf.js');
const path = require('path');
const fs = require('fs');
const { spawnAsync, execAsync } = require('./src/lib.js');

// const ppt2video = new PPT2video({
//     pptPath: path.resolve('./src/doc/test.pptx')
// });
// const a = fs.statSync(path.resolve('./src/doc/audio/tip.mp3'));
// console.log(a);
const audioPath = path.resolve('./src/doc/audio/tip1.mp3');
// getDuration();
// ffprobe -v quiet -print_format json -show_format E:\work\bnu\code\ppt2video\src\doc\audio\tip1.mp3
// ffprobe -v quiet -print_format json -show_format=duration E:\work\bnu\code\ppt2video\src\doc\audio\tip1.mp3
// ffprobe E:\work\bnu\code\ppt2video\src\doc\audio\tip1.mp3
function getTransitionType(type) {
    // 变换效果
    const transitions = [
        "fade"       , 
        "wipeleft"    ,
        "wiperight"   ,
        "wipeup"      ,
        "wipedown"    ,
        "slideleft"   ,
        "slideright"  ,
        "slideup"     ,
        "slidedown"   ,
        "circlecrop"  ,
        "rectcrop"    ,
        "distance"    ,
        "fadeblack"   ,
        "fadewhite"   ,
        "radial"      ,
        "smoothleft"  ,
        "smoothright" ,
        "smoothup"    ,
        "smoothdown"  ,
        "circleopen"  ,
        "circleclose" ,
        "vertopen"    ,
        "vertclose"   ,
        "horzopen"    ,
        "horzclose"   ,
        "dissolve"    ,
        "pixelize"    ,
        "diagtl"      ,
        "diagtr"      ,
        "diagbl"      ,
        "diagbr"      ,
    ];
    if (type && transitions.includes(type)) return type;
    const random = Math.floor(Math.random() * 31);
    return transitions[random];
}
function getSlideVideos() {
    return new Promise(async (resolve, reject) => {
        try {
            const imgFolder = path.resolve('./src/doc/img');
            const audioFolder = path.resolve('./src/doc/audio');
            const imgList = fs.readdirSync(imgFolder);
            const audioList = fs.readdirSync(audioFolder);
            for (let i = 0; i < imgList.length; i++) {
                console.log(`转换第${i}个`);
                const imgPath = path.join(imgFolder, imgList[i]);
                /**audio TODO//////////////////////////////////////// */
                const audioPath = path.join(audioFolder, audioList[i]);
                // const audioPath = audioList[i];
                /**audio TODO//////////////////////////////////////// */
                await img2video(imgPath, audioPath);
            }
            resolve();
        } catch(err) {
            console.log('errrrrrr', err);
            reject(err);
        }
    });
}
function img2video(imgPath, audioPath) {
    return new Promise(async (resolve, reject) => {
        // ffmpeg -hide_banner -loglevel quiet -loop 1 -i 1.PNG -i tip1.mp3 -c:v libx264 -c:a copy -t 5 -pix_fmt yuv420p -y 1.mp4
        try {
            // tip:拼接命令时加空格，放在前一条的末尾处
            let cmd = `ffmpeg -hide_banner -loop 1 -i ${imgPath} `;
            let duration = 0;
            // 有对应音频时，添加该音频，视频时长为音频时长
            if (existPath(audioPath)) {
                duration = await getDuration(audioPath);
                cmd += `-i ${audioPath} -c:a copy `;
            } else {
                duration = 5;
                cmd += '-f lavfi -i anullsrc=channel_layout=stereo:sample_rate=44100 ';
            }
            cmd += `-c:v libx264 -s 1920x1080 -pix_fmt yuv420p -t ${duration} -y `;
            const videoName = path.basename(imgPath, path.extname(imgPath)) + '.mp4';
            cmd += path.join(process.cwd(), videoName);
            // 执行ffmpeg命令
            await spawnAsync(cmd);
            resolve();
        } catch(err) {
            return reject(err);
        }
    });
}
function existPath(path) {
    try {
        return fs.existsSync(path);
    } catch(err) {
        return false;
    }
}
function getDuration(audioPath) {
    return new Promise((resolve, reject) => {
        // spawnAsync('ffmpeg', ['-hide_banner', '-i', audioPath]).then(data => {
        //     console.log(data.match(/^Duration: ([\d|:|\.]+),$/));
        // }).catch(err => {
        //     console.log(err.match(/Duration: ([\d|:|\.]+),/));
        // });
        const cmd = `ffprobe -v quiet -print_format json -show_format ${audioPath}`;
        spawnAsync(cmd).then(data => {
            const time = JSON.parse(data).format.duration;
            resolve(parseFloat(time));
        }).catch(err => {
            reject(err);
        });
    });
}

// 过滤videoFolder中非video文件
// -filter_complex_script 对应表达式
function createFilterScript(videoFolder, animate=false, animateDuration=1) {
    return new Promise((resolve, reject) => {
        let scriptText = '';
        try {
            const videoList = fs.readdirSync(videoFolder).filter(v => path.extname(v) === '.mp4');
            if (!animate) {
                // 没有转场动画，直接拼接
                scriptText += videoList.map((v, i) => `[${i}:v][${i}:a]`).join('');
                scriptText += `concat=n=${videoList.length}:v=1:a=1[video][audio]`;
                return resolve(scriptText);
            }
            // 创建副本，准备用来制作转场动画的片段
            videoList.forEach((v, i) => {
                // 视频流创建2个副本，副本1用来最后的拼接，副本2用来制作转场动画
                scriptText += `[${i}:v]split[v${i}][v${i}copy];`;
                // 取副本2的其中animateDuration(即设置的转场动画时间)秒，这里从开头取，
                // 如果取其他时间段，需要使用setpts=PTS-STARTPTS矫正时间戳
                scriptText += `[v${i}copy]trim=0:${animateDuration}[v${i}1];`;
                if (i > 0 && i < videoList.length - 1) {
                    // 非首尾的视频，创建animateDuration秒的两个副本，分别和前后制作转场动画
                    scriptText += `[v${i}1]split[v${i}10][v${i}11];`;
                }
            });
            // 制作转场动画
            videoList.forEach((v, i) => {
                if (i < videoList.length - 1) {
                    const v1 = `[v${i}1${i > 0 ? 1 : ''}]`;//[v01],[v111],[v211]
                    const v2 = `[v${i + 1}1${i + 1 === videoList.length - 1 ? '' : 0}]`;//[v110],[v210],...[v81]
                    scriptText += `${v1}${v2}xfade=transition=${getTransitionType()}:duration=${animateDuration}:offset=0[vt${i}];`;
                }
            });
            // 合并video和转场视频
            // scriptText += `[v0][vt0][v1][vt1]...[v7][vt7][v8]concat=n=17[video];`;
            scriptText += videoList.map((v, i) => {
                if (i < videoList.length - 1) {
                    return `[v${i}][vt${i}]`;
                } else {
                    return `[v${i}]concat=n=${2 * videoList.length - 1}[video];`;
                }
            }).join('');
            // 合并audio，给转场动画添加空白音频
            // 最后一个输入流是静音文件（即[videoList.length : a]），可供操作
            // 取animateDuration秒静音，并创建videoList.length-1个副本
            scriptText += `[${videoList.length}:a]atrim=0:1[asilent];`;
            // [asilent]asplit=8[asilent0][asilent1]...[asilent8];
            scriptText += `[asilent]asplit=${videoList.length - 1}`;
            scriptText += videoList.slice(0, -1).map((v, i) => `[asilent${i}]`).join('');
            scriptText += ';';
            // 合并音频并插入静音音频，转场动画需要对应音频流，否则正常音频会提前播放，导致音视频不同步
            scriptText += videoList.map((v, i) => {
                if (i < videoList.length - 1) {
                    return `[${i}:a][asilent${i}]`;
                } else {
                    return `[${i}:a]concat=n=${2 * videoList.length - 1}:v=0:a=1[audio]`;
                }
            }).join('');
            resolve(scriptText);
        } catch(err) {
            reject(err);
        }
    });
}

function concatVideos(videoFolder, resultFolder) {
    console.log('===================================================',videoFolder)
    return new Promise(async (resolve, reject) => {
        // -hide_banner
        let cmd = 'ffmpeg ';
        const videoList = fs.readdirSync(videoFolder).filter(v => path.extname(v) === '.mp4');
        console.log(videoList)
        // 输入文件
        const inputFiles = videoList.map(v => `-i ${path.join(videoFolder, v)}`).join(' ');
        cmd += `${inputFiles} `;
        // 静音音频
        cmd += '-f lavfi -i anullsrc=channel_layout=stereo:sample_rate=44100 ';
        /*******-filter_complex_script "file.txt */
        const scriptText = await createFilterScript(videoFolder, true);
        const scriptPath = path.resolve('./src/doc/filter.txt');
        console.log(scriptPath)
        // scriptText += videoList.map((v, i) => {
        //     let char = 0;
        //     if ( i === videoList.length - 1) {
        //         char = '';
        //     }
        //     return `[${i}:a]atrim=0:${char}[a${i}];[${i}:v]split[v${i}00][v${i}10];`;
        // }).join('');
        fs.writeFileSync(scriptPath, scriptText);
        cmd += `-safe 0 -filter_complex_script ${scriptPath} -vsync 0 `;
        /*******-filter_complex_script "file.txt */
        // settings  -map [audio] -movflags +faststart
        cmd += '-map [video] -map [audio] -movflags +faststart -y ';
        // cmd += '-profile:v high -level 3.1 -preset:v veryfast -keyint_min 72 -g 72 -sc_threshold 0 ';
        // cmd += '-b:v 3000k -minrate 3000k -maxrate 6000k -bufsize 6000k ';
        // cmd += '-b:a 128k -avoid_negative_ts make_zero -fflags +genpts -y ';
        cmd += path.join(resultFolder, 'out.mp4');
        spawnAsync(cmd).then(data => {
            resolve(data);
        }).catch(err => {
            reject(err);
        });
        // fs.writeFileSync(path.resolve('./doc/cmd.txt'), cmd);
    });
}
const videoFolder = path.resolve('./src/doc/video');
const resultFolder = path.resolve('./src/doc');
// getSlideVideos().then(data => {
//     console.log(data)
// }).catch(err => {
//     console.log(err)
// });
// concatVideos(process.cwd(), resultFolder).then(data => {
//     console.log(data);
// }).catch(err => {
//     console.log(err);
// })
// getDuration(path.resolve(videoFolder, '../out.mp4')).then(data => {
//     console.log(data);
// }).catch(err => {
//     console.log(err);
// });
// const imgPath = path.resolve('./src/doc/img/1.PNG');
// const audioPathss = path.resolve('./src/doc/audio/tip2.mp3');
// img2video(imgPath, 'ddddd').then(data => {
//     console.log(data);
// }).catch(err => {
//     console.log(err);
// });

const ppt2video = new PPT2video({
    pptPath: path.resolve('./src/doc/test.pptx'),
    animate: {
        use: true,
        // type: 'vertclose',
        type: 'inturn',
        duration: 1
    },
    tmpdir: path.resolve('./temp'),
    resultFolder: path.resolve('./result'),
    audioFolder: path.resolve('./src/doc/audio'),
    slideDuration: 3
});
ppt2video.convert().then(data => {
    console.log(data);
}).catch(err => {
    console.log(err);
});