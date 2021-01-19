const path = require('path');
const fs = require('fs');
const AdmZip = require('adm-zip');
const xml2js = require('xml2js');
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
// 读取ppt每页备注
function getNotes(pptxPath) {
    return new Promise(async (resolve, reject) => {
        if (!existPath(pptxPath)) {
            return reject('pptx文件不存在');
        }
        console.log('getNotes, 正在读取注释内容...');
        const noteList = [];
        const zip = new AdmZip(pptxPath);
        const zipEntries = zip.getEntries().filter(v => {
            return v.entryName.includes('ppt/notesSlides');
        });
        // 注释,.xml对应的是注释文件
        const noteEntries = zipEntries.filter(v => {
            return path.extname(v.entryName) === '.xml';
        });
        // 关联文件,.xml.rels对应的是该注释文件的关联关系
        const relsEntries = zipEntries.filter(v => {
            return path.extname(v.entryName) === '.rels';
        });
        for (let i = 0; i < noteEntries.length; i++) {
            const entry = noteEntries[i];
            const note = {
                slide: '',//对应幻灯片序号
                text: []//文本内容
            };
            try {
                note.text = await getText(entry);
                // 读取关联xml
                const noteName = path.basename(entry.entryName);
                const relsEntry = relsEntries.find(v => {
                    return v.entryName.includes(noteName);
                });
                note.slide = await getSlide(relsEntry);
                noteList.push(note);
            } catch(err) {
                console.log(err);
            }
        }
        resolve(noteList);
    });
}
function getText(entry) {
    // 读取注释节点
    return new Promise((resolve, reject) => {
        if (!entry || !entry.getData) {
            return reject(new Error('-_-'));
        }
        const xml = entry.getData().toString('utf8');
        xml2js.parseStringPromise(xml).then(result => {
            // 按zip格式解压pptx文件，ppt/notesSlides文件夹下即为注释相关文件
            // notesSlide*.xml是注释内容文件
            // _rels文件夹下为对应的关联文件
            // 解析xml文件，得到注释文本
            // 以下获取文本节点的方式未经详细测试，注释格式复杂时可能出现问题
            const rows = result['p:notes']['p:cSld'][0]['p:spTree'][0]['p:sp'][1]['p:txBody'][0]['a:p'];
            const textArr = rows.map(row => {
                return (row['a:r']||[]).map(v => v['a:t']).flat().join('');
            });
            resolve(textArr);
        }).catch(err => {
            reject(err);
        });
    });
}

module.exports = {
    spawnAsync,
    spawnFfmpegAsync,
    getVbsPath,
    existPath,
    getPPTVersion,
    getNotes
}