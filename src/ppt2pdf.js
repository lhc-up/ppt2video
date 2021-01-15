const path = require('path');
const del = require('del');
const os = require('os');
const fs = require('fs');
const { spawnAsync } = require('./lib.js');
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
                // Array，循环使用数组，TODO
                type: this.getTransitionType('', 0),//转场动画
                duration: 1//动画时间
            },
            tmpdir: os.tmpdir(),
            resultFolder: path.dirname(options.pptPath),//最终video输出目录
            audioFolder: path.resolve('./doc/audio'),//音频文件存放路径
            slideDuration: 5,//没有对应音频时，每页显示时长
        }
        Object.assign(config, options);
        this.config = config;
    }
    convert() {
        return new Promise(async (resolve, reject) => {
            try {
                // tmpdir
                await this.mkCleanTempDir();
                if (!this.existPath(this.config.pptPath)) throw new Error('ppt文件不存在');
                // ppt -> img
                await this.getSlideImgs();
                // img + audio -> video
                await this.getSlideVideos();
                // video + animate + video... -> video
                await this.concatVideos(this.tmpdir.video);
                resolve();
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
            await del(base);
            fs.mkdirSync(tmpdir);
            fs.mkdirSync(base);
            fs.mkdirSync(img);
            fs.mkdirSync(video);
        } catch(err) { /**console.log('目录已存在') */ }
        this.tmpdir = { base, img, video };
    }
    // 通过VBS脚本转换ppt->img
    getSlideImgs() {
        return new Promise((resolve, reject) => {
            const vbsPath = this.getVbsPath('ppt2img');
            if (!this.existPath(vbsPath)) {
                return reject(`getSlideImgs:vbs脚本(${vbsPath})不存在`);
            }
            // 通过cscript命令执行vbs脚本，报错信息通过终端输出
            const cmd = `cscript //nologo ${vbsPath} ${this.config.pptPath} ${this.tmpdir.img}`;
            spawnAsync(cmd).then(() => {
                resolve();
            }).catch(err => {
                reject(err);
            });
        });
    }
    // 每页幻灯片转video，并添加音频
    // video时长为音频时长
    getSlideVideos() {
        return new Promise(async (resolve, reject) => {
            try {
                const imgFolder = this.tmpdir.img;
                const imgList = fs.readdirSync(imgFolder);
                const audioList = fs.readdirSync(this.config.audioFolder);
                for (let i = 0; i < imgList.length; i++) {
                    const imgPath = path.join(imgFolder, imgList[i]);
                    /**generate audio TODO//////////////////////////////////////// */
                    const audioPath = path.join(this.config.audioFolder, audioList[i] || 'no-match-audio');
                    /**generate audio TODO//////////////////////////////////////// */
                    // TODO:选出跟图片对应的音频
                    await this.img2video(imgPath, audioPath);
                }
                resolve();
            } catch(err) {
                reject(err);
            }
        });
    }
    img2video(imgPath, audioPath) {
        return new Promise(async (resolve, reject) => {
            // ffmpeg -hide_banner -loglevel quiet -loop 1 -i 1.PNG -i tip1.mp3 -c:v libx264 -c:a copy -t 5 -pix_fmt yuv420p -y 1.mp4
            try {
                // tip:拼接命令时加空格，放在前一条的末尾处
                let cmd = `ffmpeg -hide_banner -loop 1 -i ${imgPath} `;
                let duration = 0;
                // 有对应音频时，添加该音频，视频时长为音频时长
                if (this.existPath(audioPath)) {
                    duration = await this.getDuration(audioPath);
                    cmd += `-i ${audioPath} -c:a copy `;
                } else {
                    duration = this.config.slideDuration;
                    // 添加静音流
                    cmd += '-f lavfi -i anullsrc=channel_layout=stereo:sample_rate=44100 ';
                }
                cmd += `-c:v libx264 -s 1920x1080 -pix_fmt yuv420p -t ${duration} -y `;
                const videoName = path.basename(imgPath, path.extname(imgPath)) + '.mp4';
                cmd += path.join(this.tmpdir.video, videoName);
                // 执行ffmpeg命令
                await spawnAsync(cmd);
                resolve();
            } catch(err) {
                return reject(err);
            }
        });
    }
    concatVideos(videoFolder) {
        return new Promise(async (resolve, reject) => {
            try {
                let cmd = 'ffmpeg -hide_banner ';
                // TODO:videoList文件顺序，按照序号排序
                const videoList = fs.readdirSync(videoFolder).filter(v => {
                    return path.extname(v) === '.mp4';
                });
                // 输入文件
                const inputFiles = videoList.map(v => `-i ${path.join(videoFolder, v)}`).join(' ');
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
                const pptPath = this.config.pptPath;
                const videoName = path.basename(pptPath, path.extname(pptPath)) + '.mp4';
                cmd += path.resolve(this.config.resultFolder, videoName);

                await spawnAsync(cmd);
                // fs.writeFileSync(path.resolve('./doc/cmd.txt'), cmd);
                resolve();
            } catch(err) {
                reject(err);
            }
        });
    }
    getDuration(audioPath) {
        return new Promise((resolve, reject) => {
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
        // TODO:type支持多种配置
        if (type instanceof Array) {
            // TODO:循环使用数组中的值
            type = this.getTransitionType('', 0);
        }else if (type === 'random') {
            type = this.getTransitionType();
        } else {
            // inturn,randomAll在循环中处理
        }
        /***************************************************** */
        /***************************************************** */
        countArr.forEach(i => {
            if (i < videoCount - 1) {
                const v1 = `[v${i}1${i > 0 ? 1 : ''}]`;//[v01],[v111],[v211]
                const v2 = `[v${i + 1}1${i + 1 === videoCount - 1 ? '' : 0}]`;//[v110],[v210],...[v81]
                let currType;
                if (type === 'inturn') {
                    currType = this.getTransitionType('', i);
                } else if (type === 'randomAll') {
                    currType = this.getTransitionType();
                }
                scriptText += `${v1}${v2}xfade=transition=${currType || type}:duration=${duration}:offset=0[vt${i}];`;
            }
        });
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
    // 转场动画类型
    // type:取具体值
    // index：按序号取值
    getTransitionType(type, index) {
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
        if (index !== undefined) {
            if (index < transitions.length) {
                return transitions[index];
            } else {
                // 超过之后从头开始
                return transitions[index - transitions.length]
            }
        }
        const random = Math.floor(Math.random() * 31);
        return transitions[random];
    }
    // 路径是否存在
    existPath(path) {
        try {
            return fs.existsSync(path);
        } catch(err) {
            return false;
        }
    }
    // 获取ppt版本
    getPPTVersion() {
        return new Promise((resolve, reject) => {
            const vbsPath = this.getVbsPath('pptv');
            if (!this.existPath(vbsPath)) {
                return reject(`getPPTVersion:vbs脚本(${vbsPath})不存在`);
            }
            const cmd = `cscript //nologo ${vbsPath}`;
            spawnAsync(cmd).then(version => {
                resolve(version);
            }).catch(err => {
                reject(err);
            });
        });
    }
    getVbsPath(name) {
        return path.join(process.cwd(), `src/vbs/${name}.vbs`);
    }
}

module.exports = {
    PPT2video
}