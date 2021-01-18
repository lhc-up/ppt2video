const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
function spawnAsync(cmd, progressCb) {
    return new Promise((resolve, reject) => {
        const cmdArr = cmd.split(' ');
        const cp = spawn(cmdArr.shift(), cmdArr);
        const stdout = [], stderr = [];

        cp.on('error', err => reject(err));
        cp.stderr.on('data', data => {
            stderr.push(data);
            if (progressCb && progressCb instanceof Function) {
                progressCb('stderr', data.toString());
            }
        });
        cp.stdout.on('data', data => {
            stdout.push(data);
            if (progressCb && progressCb instanceof Function) {
                progressCb('stdout', data.toString());
            }
        });

        // FFmpeg使用stdout来输出媒体数据，并使用stderr来记录/进度信息
        // 这里使用进程退出码判断成功或失败
        cp.on('close', code => {
            const info = Buffer.concat([...stdout, ...stderr]).toString();
            if (!code) {
                resolve(info);
            } else {
                reject(info);
            }
        });
    });
}
/**
 * 执行ffmpeg命令，需要返回进度的场景
 * @param {String} cmd 命令
 * @param {Number} inputCount 输入文件数
 * @param {Number} tolerance 总时长容差s
 * @param {Function} progressCb 进度回调
 */
function spawnFfmpegAsync(cmd, inputCount, tolerance, progressCb) {
    return new Promise((resolve, reject) => {
        // 各个输入文件的时长
        const durationArr = [];
        spawnAsync(cmd, (stdType, data) => {
            if (stdType !== 'stderr') return false;
            const duration = getSecondsFromStd(data, /Duration: ([\d|:|\.]+)/);
            if (duration) {
                durationArr.push(duration);
            }
            // 获取到所有输入文件的时长时，开始获取当前时间，用于计算进度
            if (durationArr.length < inputCount) return false;
            let totalSeconds = durationArr.reduce((prev, cur) => prev + cur, 0);
            totalSeconds += tolerance;

            const seconds = getSecondsFromStd(data, /time=([\d|:|\.]+)/);
            let percent = 0;
            if (totalSeconds <= 0) {
                percent = 0;
            } else {
                percent = Math.round(seconds / totalSeconds * 10000) / 100;
                percent = percent >= 100 ? 100 : percent;
            }
            if (percent >0 && progressCb && progressCb instanceof Function) {
                progressCb(percent, seconds, totalSeconds);
            }
        }).then(data => {
            resolve(data);
        }).catch(err => {
            reject(err);
        });
    });
}
function getSecondsFromStd(data, reg) {
    const timeArr = data.match(reg);
    if (!Array.isArray(timeArr)) return 0;
    const time = timeArr[1];
    if (!time) return 0;
    // 00:00:07.04
    const [hour, minute, second] = time.split(':').map(v => parseFloat(v));
    return hour * 3600 + minute * 60 + second;
}
// 获取ppt软件版本
function getPPTVersion() {
    return new Promise((resolve, reject) => {
        const vbsPath = getVbsPath('pptv');
        if (!existPath(vbsPath)) {
            return reject(`getPPTVersion:vbs脚本(${vbsPath})不存在`);
        }
        const cmd = `cscript //nologo ${vbsPath}`;
        spawnAsync(cmd).then(v => {
            const version = parseFloat(v);
            if (!!version) {
                resolve(version)
            } else {
                reject(v);
            }
        }).catch(err => {
            reject(err);
        });
    });
}
function getVbsPath(name) {
    return path.join(__dirname, 'vbs', `${name}.vbs`);
}
// 路径是否存在
function existPath(path) {
    try {
        if (path instanceof Array) {
            return path.length && path.every(v => fs.existsSync(v));
        } else {
            return fs.existsSync(path);
        }
    } catch(err) {
        return false;
    }
}

module.exports = {
    spawnAsync,
    spawnFfmpegAsync,
    getVbsPath,
    existPath,
    getPPTVersion
}