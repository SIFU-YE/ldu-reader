#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""安装/检查 LTP 运行环境（包 + 模型），用于干净克隆后的首次配置。

【适用场景】仅源码开发/打包机需要；exe 用户无需执行——exe 已经过
collect_all 固化依赖与模型，双击即可运行。

用法:
  python3 install_ltp_env.py            # 检查
  python3 install_ltp_env.py --install  # 安装到 .libs312（使用当前 pip 镜像）
  python3 install_ltp_env.py --model    # 仅下载 LTP/small 模型

已知限制：check_torch 的检查子进程未注入 .libs312 到 PYTHONPATH——安装成功后
子进程仍可能报"缺失/不可用"（误报）；以 .libs312 目录内包文件存在为准。
"""
import argparse, os, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(HERE, '.libs312')
MODEL_DIR = os.path.join(HERE, 'LTP', 'small')


def check_pip():
    try:
        subprocess.run([sys.executable, '-m', 'pip', '--version'],
                       capture_output=True, text=True, check=True)
        return True
    except Exception:
        return False


def check_ltp_package():
    for p in [os.path.join(TARGET, 'ltp', '__init__.py'),
              os.path.join(TARGET, 'ltp_core', '__init__.py'),
              os.path.join(TARGET, 'ltp_extension', '__init__.py')]:
        if not os.path.exists(p):
            return False, p
    return True, ''


def check_model():
    return os.path.exists(os.path.join(MODEL_DIR, 'config.json'))


def check_torch():
    try:
        r = subprocess.run([sys.executable, '-c', 'import torch; print(torch.__version__)'],
                           capture_output=True, text=True, timeout=30)
        return r.returncode == 0, (r.stdout or '').strip()
    except Exception:
        return False, ''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--install', action='store_true', help='安装 LTP 包到 .libs312')
    ap.add_argument('--model', action='store_true', help='仅下载 LTP/small 模型')
    args = ap.parse_args()

    if args.install:
        if not check_pip():
            print('当前 Python 无 pip，请先安装 pip。', file=sys.stderr)
            sys.exit(1)
        print('安装 torch（默认 pip 镜像；无 GPU 建议 CPU 版）...')
        # 尝试从镜像安装 torch；若镜像为 CUDA 版且过慢，可手动安装 CPU 版：
        # python -m pip install torch --index-url https://download.pytorch.org/whl/cpu
        r = subprocess.run([sys.executable, '-m', 'pip', 'install', '--no-cache-dir',
                            '--target', TARGET, 'torch', 'transformers'],
                           capture_output=False)
        if r.returncode != 0:
            print('torch/transformers 安装失败，请手动执行后重试。', file=sys.stderr)
            sys.exit(1)
        print('安装 ltp / ltp-core / ltp-extension ...')
        r = subprocess.run([sys.executable, '-m', 'pip', 'install', '--no-cache-dir',
                            '--no-deps', '--target', TARGET, 'ltp', 'ltp-core', 'ltp-extension'])
        if r.returncode != 0:
            print('ltp 安装失败。', file=sys.stderr)
            sys.exit(1)
        print('LTP 包安装完成。')

    if args.model:
        import download_ltp_model
        download_ltp_model.main()
        return

    # 检查报告
    pkg_ok, miss = check_ltp_package()
    model_ok = check_model()
    torch_ok, torch_ver = check_torch()
    print('=' * 56)
    print('LTP 环境检查')
    print('=' * 56)
    print(f'Python          : {sys.executable}')
    print(f'pip             : {"OK" if check_pip() else "缺失"}')
    print(f'torch           : {"OK " + torch_ver if torch_ok else "缺失/不可用"}')
    print(f'ltp 包          : {"OK (.libs312)" if pkg_ok else "缺失 " + miss}')
    print(f'LTP/small 模型  : {"OK" if model_ok else "缺失"}')
    if not pkg_ok or not model_ok or not torch_ok:
        print('-' * 56)
        print('修复命令:')
        print('  python3 install_ltp_env.py --install')
        print('  python3 install_ltp_env.py --model')
        print('或手动指定已有 LTP 解释器:')
        print('  LTP_PYTHON=/path/to/python python3 batch_report_v4.py ...')
    else:
        print('环境完整，可以运行 LTP 主流程。')


if __name__ == '__main__':
    main()
