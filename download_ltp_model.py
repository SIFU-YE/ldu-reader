#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""下载 LTP/small 模型到 ./LTP/small（优先 hf-mirror.com）。"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(HERE, 'LTP', 'small')


def split_model(local_dir, chunk_mb=90):
    """把下载完成的 pytorch_model.bin 拆分为 <100MB 分片（partN），删除完整 bin。

    Gitee 不支持 git-lfs 且单文件限 100MB，仓库模型按 90MB 分片保存；
    运行时由 core._ensure_model_merged 自动合并回 pytorch_model.bin。
    """
    bin_path = os.path.join(local_dir, "pytorch_model.bin")
    if not os.path.exists(bin_path):
        return False
    size = os.path.getsize(bin_path)
    chunk = chunk_mb * 1024 * 1024
    if size <= chunk:
        return False  # 无需拆分
    with open(bin_path, "rb") as f:
        i = 1
        while True:
            part = f.read(chunk)
            if not part:
                break
            with open(f"{bin_path}.part{i}", "wb") as pf:
                pf.write(part)
            i += 1
    os.remove(bin_path)
    return True


def main():
    if os.path.isdir(TARGET) and 'config.json' in os.listdir(TARGET):
        print('模型已存在:', TARGET)
        return
    os.makedirs(os.path.dirname(TARGET), exist_ok=True)
    os.environ.setdefault('HF_ENDPOINT', 'https://hf-mirror.com')
    os.environ.setdefault('HF_HUB_DISABLE_XET', '1')
    try:
        from huggingface_hub import snapshot_download
    except ImportError:
        print('缺少 huggingface_hub，请先安装：pip install huggingface_hub', file=sys.stderr)
        sys.exit(1)
    print('开始下载 LTP/small（约355MB）...')
    snapshot_download(repo_id='LTP/small', local_dir=TARGET)
    if split_model(TARGET):
        print('已拆分模型为分片（<100MB，可入库 Gitee）:', TARGET)
    else:
        print('已下载到:', TARGET)


if __name__ == '__main__':
    main()
