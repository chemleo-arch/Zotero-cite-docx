# -*- coding: utf-8 -*-
"""zotero_paths.py — 跨平台定位本机 Zotero 数据目录 / zotero.sqlite。

Zotero 的数据目录不一定在默认位置（设置 → 高级 → 文件和文件夹 → 数据目录位置
可以改，改过之后 profile 里 useDataDir=true）。因此这里按优先级依次探测：

  1. 显式传入的参数（--db / --sqlite）
  2. 环境变量 ZOTERO_SQLITE       —— 直接指向 zotero.sqlite 文件
  3. 环境变量 ZOTERO_DATA_DIR     —— 指向数据目录，其下取 zotero.sqlite
  4. 常见默认位置（按平台）

用法：
    from zotero_paths import default_db
    con = sqlite3.connect("file:%s?immutable=1" % default_db().replace("\\\\", "/"), uri=True)

注意：SQLite 打开时请用 `?immutable=1` 只读模式 —— Zotero 客户端正在运行时会独占
数据库锁，普通只读模式（mode=ro）会被阻塞，immutable 模式则能读到一致快照。
"""
import os
import sys

SQLITE_NAME = "zotero.sqlite"


def candidates():
    """按优先级产出候选 zotero.sqlite 路径。"""
    env_sqlite = os.environ.get("ZOTERO_SQLITE")
    if env_sqlite:
        yield env_sqlite

    env_dir = os.environ.get("ZOTERO_DATA_DIR")
    if env_dir:
        yield os.path.join(env_dir, SQLITE_NAME)

    home = os.path.expanduser("~")
    if sys.platform.startswith("win"):
        # Windows 上程序与数据目录常被放在非系统盘（D:/E: 等）
        for drive in ("D:", "C:", "E:"):
            yield drive + os.sep + "Zotero" + os.sep + SQLITE_NAME
        yield os.path.join(home, "Zotero", SQLITE_NAME)
    else:
        # macOS / Linux 的 Zotero 默认数据目录均位于用户主目录下
        yield os.path.join(home, "Zotero", SQLITE_NAME)


def default_db():
    """返回首个存在的候选路径；都不存在时返回优先级最高的候选，便于报错提示。"""
    cands = list(candidates())
    for path in cands:
        if os.path.exists(path):
            return path
    return cands[0] if cands else SQLITE_NAME


def sqlite_uri(db_path):
    """把本地路径转成 sqlite3 的 immutable 只读 URI。"""
    return "file:%s?immutable=1" % str(db_path).replace("\\", "/")


if __name__ == "__main__":
    print("候选顺序：")
    for c in candidates():
        print("  %s %s" % ("[存在]" if os.path.exists(c) else "[    ]", c))
    print("解析结果：%s" % default_db())
