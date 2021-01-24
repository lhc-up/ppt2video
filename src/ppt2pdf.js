const path = require('path');
const del = require('del');
const os = require('os');
const fs = require('fs');
const { transitions, getTransitionType } = require('./transitions.js');
const {
    spawnAsync,
    spawnFfmpegAsync,
    getVbsPath,
    existPath,
    getPPTVersion,
    isOSValid
} = require('./lib.js');
// ffmpeg下载及使用参考
// https://www.ffmpeg.org/
// https://trac.ffmpeg.org/wiki
class PPT2video {
    constructor(options) {
        // 默认配置
        const config = {
            pptPath: '',
            animate: {
                use: false,
                // 可能的值:
                // inturn，按顺序使用
                // random，随机一种，
                // randomAll,每页都随机,
                // String,指定具体值
                // Array，循环使用数组
                type: getTransitionType('', 0),//转场动画
                duration: 1//动画时间
            },
            tmpdir: os.tmpdir(),
            resultFolder: path.dirname(options.pptPath),//最终video输出目录
            audioFolder: path.resolve('./doc/audio'),//音频文件存放路径
            slideDuration: 5,//没有对应音频时，每页显示时长
        }
        Object.assign(config, options);
        this.config = config;
        // tmpdir
        this.mkCleanTempDir();
    }
    convert(progressCb) {
        return new Promise(async (resolve, reject) => {
            try {
                const { pptPath, audioFolder, resultFolder } = this.config;
                // ppt -> img
                const imgFolder = await this.getSlideImgs(pptPath);
                // img + audio -> video
                const getSlideVideosScale = 0.42;
                const concatVideosScale = 1 - getSlideVideosScale;
                const videoFolder = await this.getSlideVideos(imgFolder, audioFolder, '', (percent, cur, total) => {
                    if (progressCb && progressCb instanceof Function) {
                        progressCb(percent * getSlideVideosScale);
                    }
                });
                // video + animate + video... -> video
                const videoName = path.basename(pptPath, path.extname(pptPath)) + '.mp4';
                const resultPath = await this.concatVideos(videoFolder, resultFolder, {
                    resultName: videoName
                }, (percent, cur, total) => {
                    if (progressCb && progressCb instanceof Function) {
                        progressCb(percent * concatVideosScale + getSlideVideosScale * 100);
                    }
                });
                resolve(resultPath);
            } catch(err) {
                reject(err);
            } 
        });
    }
    // 创建干净的(如已存在，删除其中文件)临时文件夹，用于存放过程文件
    async mkCleanTempDir() {
        const tmpdir = this.config.tmpdir || os.tmpdir();
        const base = path.join(tmpdir, 'ppt2video');//根目录
        const img = path.join(base, 'img');//临时图片
        const video = path.join(base, 'video');//临时视频
        try {
            del.sync(base);
            fs.mkdirSync(tmpdir);
        } catch(err) { /**console.log('目录已存在') */ }
        // node < 10,依次创建目录及文件
        try { fs.mkdirSync(base); } catch(err) { /**console.log('目录已存在') */ }
        try { fs.mkdirSync(img); } catch(err) { /**console.log('目录已存在') */ }
        try { fs.mkdirSync(video); } catch(err) { /**console.log('目录已存在') */ }
        this.tmpdir = { base, img, video };
    }
    /**
     * 通过VBS脚本转换ppt->img
     * @param {String} pptPath 需要转换的ppt路径
     * @param {String} imgFolder 输出目录，可选
     */
    getSlideImgs(pptPath, imgFolder) {
        return new Promise(async (resolve, reject) => {
            if (!isOSValid()) {
                return reject('getSlideImgs方法仅支持windows系统');
            }
            try {
                const version = await getPPTVersion();
                if (!version) return reject('未检测到ppt版本');
            } catch(err) {
                return reject('未检测到ppt版本，未安装office PowerPoint，或不可用');
            }
            if (!existPath(pptPath)) {
                return reject(`getSlideImgs:ppt文件(${pptPath})不存在`);
            }
            
            const vbsPath = getVbsPath('ppt2img');
            if (!existPath(vbsPath)) {
                return reject(`getSlideImgs:vbs脚本(${vbsPath})不存在`);
            }

            if (!imgFolder || !existPath(imgFolder)) {
                imgFolder = this.tmpdir.img;
            }

            // 通过cscript命令执行vbs脚本，报错信息通过终端输出
            const cmd = `cscript //nologo ${vbsPath} ${pptPath} ${imgFolder}`;
            spawnAsync(cmd).then(() => {
                resolve(imgFolder);
            }).catch(err => {
                reject(err);
            });
        });
    }
    /**
     * 幻灯片批量转video，并添加音频
     * 图片、音频可能存在多种格式，这里不做过滤，目录中不要放其他无关文件
     * @param {String} imgFolder 图片目录
     * @param {String} audioFolder 音频目录
     * @param {String} videoFolder 输出目录
     * @returns {String} 输出目录
     */
    getSlideVideos(imgFolder, audioFolder, videoFolder, progressCb) {
        return new Promise(async (resolve, reject) => {
            if (!existPath(imgFolder)) {
                imgFolder = this.tmpdir.img;
            }
            if (!existPath(audioFolder)) {
                audioFolder = this.config.audioFolder;
            }
            if (!existPath(videoFolder)) {
                videoFolder = this.tmpdir.video;
            }
            try {
                const imgList = fs.readdirSync(imgFolder);
                const audioList = fs.readdirSync(audioFolder);
                for (let i = 0; i < imgList.length; i++) {
                    const imgPath = path.join(imgFolder, imgList[i]);
                    /**generate audio TODO//////////////////////////////////////// */
                    const audioPath = path.join(audioFolder, audioList[i] || 'no-match-audio');
                    /**generate audio TODO//////////////////////////////////////// */
                    // TODO:选出跟图片对应的音频
                    await this.img2video(imgPath, audioPath, videoFolder);
                    if (progressCb && progressCb instanceof Function) {
                        const percent = Math.round((i + 1) / imgList.length * 10000) / 100;
                        progressCb(percent, i + 1, imgList.length);
                    }
                }
                resolve(videoFolder);
            } catch(err) {
                reject(err);
            }
        });
    }
    /**
     * img -> video
     * @param {String} imgPath 图片地址
     * @param {String} audioPath 音频地址，可选
     * @param {String} videoFolder 输出目录，可选
     * @returns {String} 输出目录
     */
    img2video(imgPath, audioPath, videoFolder, progressCb) {
        return new Promise(async (resolve, reject) => {
            if (!existPath(imgPath)) {
                return reject(`img2video:输入文件${imgPath}不存在`);
            }

            // output folder
            if (!videoFolder || !existPath(videoFolder)) {
                videoFolder = this.tmpdir.video;
            }

            try {
                // ffmpeg -hide_banner -loglevel quiet -loop 1 -i 1.PNG -i tip1.mp3 \
                // -c:v libx264 -c:a copy -t 5 -pix_fmt yuv420p -y 1.mp4
                // tip:拼接命令时空格在前一条的末尾处
                let cmd = `ffmpeg -hide_banner -loop 1 -i ${imgPath} `;
                let duration = this.config.slideDuration;
                // 有对应音频时，添加该音频，视频时长为音频时长
                if (existPath(audioPath)) {
                    duration = await this.getDuration(audioPath);
                    cmd += `-i ${audioPath} -c:a copy `;
                } else {
                    // 添加静音流
                    cmd += '-f lavfi -i anullsrc=channel_layout=stereo:sample_rate=44100 ';
                }
                cmd += `-c:v libx264 -s 1920x1080 -pix_fmt yuv420p -t ${duration} -y `;
                const videoName = path.basename(imgPath, path.extname(imgPath)) + '.mp4';
                cmd += path.join(videoFolder, videoName);
                // 执行ffmpeg命令
                await spawnFfmpegAsync(cmd, 1, 0, (percent, seconds, totalSeconds) => {
                    if (progressCb && progressCb instanceof Function) {
                        progressCb(percent, seconds, totalSeconds);
                    }
                });
                resolve(videoFolder);
            } catch(err) {
                return reject(err);
            }
        });
    }
    /**
     * 拼接视频
     * @param {String} videoFolder 视频片段目录
     * @param {String} resultFolder 输出目录
     * @param {String} resultName 输出名称
     * @param {Array} exts 合法的后缀名，除此之外的过滤掉
     * @param {Function} sortFn 排序函数
     * @returns {String} 输出文件路径
     */
    concatVideos(videoFolder, resultFolder, {
        resultName='result.mp4',
        exts=['.mp4'],
        sortFn
    }, progressCb) {
        return new Promise(async (resolve, reject) => {
            // 默认以文件名称末尾的数字大小排序（js默认是字典序）
            // 可自定义排序函数
            sortFn = sortFn || ((a, b) => {
                const reg = /\d+/g;
                const aIndex = (a.split('.')[0].match(reg) || [0]).reverse()[0];
                const bIndex = (b.split('.')[0].match(reg) || [0]).reverse()[0];
                return aIndex - bIndex;
            });
            if (!existPath(videoFolder)) {
                videoFolder = this.tmpdir.video;
            }
            if (!existPath(resultFolder)) {
                resultFolder = this.config.resultFolder;
                try {
                    fs.mkdirSync(resultFolder);
                } catch(err) {/**console.log(err) */}
            }
            try {
                let cmd = 'ffmpeg -hide_banner ';
                const videoList = fs.readdirSync(videoFolder).filter(v => {
                    return exts.includes(path.extname(v));
                }).sort(sortFn);

                if (!videoList.length) {
                    return reject('没有视频文件');
                }

                // 输入文件
                const inputFiles = videoList.map(v => {
                    return `-i ${path.join(videoFolder, v)}`;
                }).join(' ');
                cmd += `${inputFiles} `;

                if (this.config.animate.use) {
                    // 使用转场动画时，构造静音音频流输入，用来填补转场动画对应的音频（filter中处理）
                    cmd += '-f lavfi -i anullsrc=channel_layout=stereo:sample_rate=44100 ';
                }
                /*******-filter_complex_script "filterScript.txt */
                const scriptText = this.createFilterScript(videoList.length, this.config.animate);
                const scriptPath = path.join(this.tmpdir.base, 'filterScript.txt');
                fs.writeFileSync(scriptPath, scriptText);
                cmd += `-safe 0 -filter_complex_script ${scriptPath} -vsync 0 `;
                /*******-filter_complex_script "filterScript.txt */
                // settings
                cmd += '-map [video] -map [audio] -movflags +faststart -y ';
                // cmd += '-profile:v high -level 3.1 -preset:v veryfast -keyint_min 72 -g 72 -sc_threshold 0 ';
                // cmd += '-b:v 3000k -minrate 3000k -maxrate 6000k -bufsize 6000k ';
                // cmd += '-b:a 128k -avoid_negative_ts make_zero -fflags +genpts -y ';
                const resultPath = path.resolve(resultFolder, resultName);
                cmd += resultPath;

                const inputCount = videoList.length;
                let animateSeconds = 0;
                // 有动画时需要加上转场动画耗时
                const { use, duration } = this.config.animate;
                if (use) {
                    animateSeconds = (inputCount - 1) * duration;
                }
                await spawnFfmpegAsync(cmd, inputCount, animateSeconds, (percent, seconds, totalSeconds) => {
                    if (progressCb && progressCb instanceof Function) {
                        progressCb(percent, seconds, totalSeconds);
                    }
                });
                resolve(resultPath);
            } catch(err) {
                reject(err);
            }
        });
    }
    // 获取音频时长
    getDuration(audioPath) {
        return new Promise((resolve, reject) => {
            if (!existPath(audioPath)) {
                return reject('音频不存在');
            }
            const cmd = `ffprobe -v quiet -print_format json -show_format ${audioPath}`;
            spawnAsync(cmd).then(data => {
                const time = JSON.parse(data).format.duration;
                resolve(parseFloat(time));
            }).catch(err => {
                reject(err);
            });
        });
    }
    // -filter_complex_script 对应表达式
    createFilterScript(videoCount, animate={}) {
        if (videoCount <= 0) {
            throw new Error('没有视频文件输入!');
        }
        
        let { use=false, type, duration=1 } = animate;
        // 只有一个视频时无法使用转场动画
        if (videoCount === 1) use = false;
        // 构造0-videoCount的数组，方便后面使用
        const countArr = Array(videoCount).fill('').map((v, i) => i);

        if (!use) {
            // 没有转场动画，直接拼接
            const script = countArr.map(i => `[${i}:v][${i}:a]`).join('') + `concat=n=${videoCount}:v=1:a=1[video][audio]`;
            return script;
        }

        // 有转场动画
        let scriptText = '';
        // 创建副本，准备用来制作转场动画的片段
        countArr.forEach(i => {
            // 视频流创建2个副本，副本1用来最后的拼接，副本2用来制作转场动画
            scriptText += `[${i}:v]split[v${i}][v${i}copy];`;
            // 取副本2的其中duration(即设置的转场动画时间)秒，这里从开头取，
            // 如果取其他时间段，需要使用setpts=PTS-STARTPTS矫正时间戳
            scriptText += `[v${i}copy]trim=0:${duration}[v${i}1];`;
            if (i > 0 && i < videoCount - 1) {
                // 非首尾的视频，创建duration秒的两个副本，分别和前后制作转场动画
                scriptText += `[v${i}1]split[v${i}10][v${i}11];`;
            }
        });
        // 制作转场动画
        /***************************************************** */
        /***************************************************** */
        if (type === 'random') {
            type = getTransitionType();
        }
        countArr.forEach(i => {
            if (i < videoCount - 1) {
                const v1 = `[v${i}1${i > 0 ? 1 : ''}]`;//[v01],[v111],[v211]
                const v2 = `[v${i + 1}1${i + 1 === videoCount - 1 ? '' : 0}]`;//[v110],[v210],...[v81]
                let currType;
                if (type === 'inturn') {
                    currType = getTransitionType('', i);
                } else if (type === 'randomAll') {
                    currType = getTransitionType();
                } else if (type instanceof Array) {
                    currType = getTransitionType('', i, type);
                }
                scriptText += `${v1}${v2}xfade=transition=${currType || type}:duration=${duration}:offset=0[vt${i}];`;
            }
        });
        /***************************************************** */
        /***************************************************** */
        // 合并video和转场视频
        // scriptText += `[v0][vt0][v1][vt1]...[v7][vt7][v8]concat=n=17[video];`;
        scriptText += countArr.map(i => {
            if (i < videoCount - 1) {
                return `[v${i}][vt${i}]`;
            } else {
                return `[v${i}]concat=n=${2 * videoCount - 1}[video];`;
            }
        }).join('');
        // 合并audio，给转场动画添加空白音频
        // 最后一个输入流是静音文件（即[videoCount : a]），可供操作
        // 取duration秒静音，并创建videoCount-1个副本
        scriptText += `[${videoCount}:a]atrim=0:1[asilent];`;
        // [asilent]asplit=8[asilent0][asilent1]...[asilent8];
        scriptText += `[asilent]asplit=${videoCount - 1}`;
        scriptText += countArr.slice(0, -1).map(i => `[asilent${i}]`).join('');
        scriptText += ';';
        // 合并音频并插入静音音频，转场动画需要对应音频流，否则正常音频会提前播放，导致音视频不同步
        scriptText += countArr.map(i => {
            if (i < videoCount - 1) {
                return `[${i}:a][asilent${i}]`;
            } else {
                return `[${i}:a]concat=n=${2 * videoCount - 1}:v=0:a=1[audio]`;
            }
        }).join('');
        return scriptText;
    }
}

// statics
PPT2video.transitions = transitions;
PPT2video.getPPTVersion = getPPTVersion;

module.exports = {
    PPT2video
}