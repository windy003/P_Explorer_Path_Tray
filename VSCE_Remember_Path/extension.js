// Cursor / VS Code 扩展:每次启动带有 workspace 的窗口时,把项目路径写入
// 一个 JSON 文件供「Explorer Path Tray」托盘应用读取。
//
// 文件位置固定为托盘 app 项目根目录下的 cursor_recent.json。
// 列表 LRU 排序(最新在最前),最多 MAX_ENTRIES 条,大小写不敏感去重。

const vscode = require('vscode');
const fs = require('fs');
const path = require('path');

const TARGET_FILE = 'D:\\files\\using\\Python\\P_Explorer_Path_Tray\\cursor_recent.json';
const MAX_ENTRIES = 20;

function recordProjectPath(projectPath) {
    let list = [];
    try {
        if (fs.existsSync(TARGET_FILE)) {
            const content = fs.readFileSync(TARGET_FILE, 'utf8');
            const parsed = JSON.parse(content);
            if (Array.isArray(parsed)) {
                list = parsed.filter((p) => typeof p === 'string');
            }
        }
    } catch (e) {
        // 读不出来就当空列表处理
        list = [];
    }

    const lower = projectPath.toLowerCase();
    list = list.filter((p) => p.toLowerCase() !== lower);
    list.unshift(projectPath);
    if (list.length > MAX_ENTRIES) {
        list.length = MAX_ENTRIES;
    }

    const dir = path.dirname(TARGET_FILE);
    try {
        fs.mkdirSync(dir, { recursive: true });
    } catch (e) {
        // 目录已存在或无权限,后面写入会再报
    }

    const payload = JSON.stringify(list, null, 2);
    const tmp = TARGET_FILE + '.tmp';
    try {
        fs.writeFileSync(tmp, payload, 'utf8');
        fs.renameSync(tmp, TARGET_FILE);
    } catch (e) {
        // 原子写失败(比如目标被其他进程占用),退化为直接覆盖
        try {
            fs.writeFileSync(TARGET_FILE, payload, 'utf8');
        } catch (e2) {
            console.error('[vsce-remember-path] write failed:', e2);
        }
    }
}

function pickProjectPath() {
    const folders = vscode.workspace.workspaceFolders;
    if (!folders || folders.length === 0) {
        return null;
    }
    // 多根工作区只取第一个根目录
    const uri = folders[0].uri;
    if (!uri || uri.scheme !== 'file') {
        return null;
    }
    return uri.fsPath || null;
}

function activate(context) {
    try {
        const projectPath = pickProjectPath();
        if (projectPath) {
            recordProjectPath(projectPath);
        }
    } catch (e) {
        console.error('[vsce-remember-path] activate error:', e);
    }
}

function deactivate() {}

module.exports = { activate, deactivate };
