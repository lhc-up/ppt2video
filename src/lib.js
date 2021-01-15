const { spawn } = require('child_process');
function spawnAsync(cmd) {
    return new Promise((resolve, reject) => {
        console.log(cmd)
        const cmdArr = cmd.split(' ');
        const cp = spawn(cmdArr.shift(), cmdArr);
        const stdout = [], stderr = [];

        cp.on('error', err => reject(err));

        cp.stderr.on('data', data => stderr.push(data));
        cp.stdout.on('data', data => stdout.push(data));

        // ffmpeg log信息通过stderr输出，这里使用进程退出码判断成功或失败
        cp.on('close', code => {
            console.log(code)
            if (!code) {
                resolve(Buffer.concat(stdout).toString());
            } else {
                const err = Buffer.concat(stderr).toString();
                reject(err);
            }
        });
    });
}

module.exports = {
    spawnAsync
}