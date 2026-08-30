#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fluent_mesh_check.py — 基于 PyFluent 的网格质量检查工具
====================================================================

功能：
  1. 利用 PyFluent 读取 Fluent 网格文件，支持的格式：
        *.msh       （纯网格，文本格式）
        *.msh.h5    （纯网格，HDF5 格式）
        *.cas       （算例，文本格式，含网格与边界）
        *.cas.h5    （算例，HDF5 格式，含网格与边界）
  2. 利用 solver.scheme_eval.scheme_eval 执行 Fluent 的 scheme 命令，
     触发网格质量报告，并通过同一 scheme_eval 通道获取报告文本，
     解析出正交质量、歪斜度、长宽比等关键质量指标。

核心 API 说明：
    solver.scheme 是 SchemeInterpreter 实例（scheme_eval 自 0.32 起已弃用、等价但会告警），提供两种调用：
      - solver.scheme.eval(expr)     执行 scheme 表达式（本工具用于触发报告）
      - solver.scheme.exec(commands) 执行命令序列并返回 TUI 输出字符串（用于获取文本）

网格维度：
    维度检测逻辑已独立到同目录的 fluent_dim_check.py，本文件仅在启动
    Fluent 前调用其 detect_dimension() 获取维度结果：
      - 省略 --dim 时自动调用 fluent_dim_check.detect_dimension() 检测；
      - 也可手动用 --dim 2 / --dim 3 指定，跳过自动检测。
    若只想查看维度而不做质量检查，请直接运行：
        python fluent_dim_check.py <网格文件路径>

用法：
    python fluent_mesh_check.py <网格文件路径> [--dim 2|3] [--cores N] [--precision single|double] [--json FILE]

示例：
    python fluent_mesh_check.py test.cas.h5
    python fluent_mesh_check.py VM02.msh
    python fluent_mesh_check.py PTE.cas.h5 --cores 8 --json result.json
    python fluent_mesh_check.py VM02.msh --dim 2
"""
import argparse
import json
import os
import re
import sys

# 统一标准输出编码，避免中文/特殊字符在 Windows 控制台乱码
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import ansys.fluent.core as pyfluent

# 网格维度判断由独立模块 fluent_dim_check.py 提供
try:
    from fluent_dim_check import detect_dimension
except ImportError:  # 以其他工作目录运行时，确保能找到同目录模块
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from fluent_dim_check import detect_dimension


# --------------------------------------------------------------------------
# 1. 文件类型识别与网格读取
# --------------------------------------------------------------------------
def _read_via_tui(solver, file_type: str, path: str) -> None:
    """优先使用 tui 通道读取（对所有格式通用、最底层）。"""
    if file_type == "mesh":
        solver.tui.file.read_mesh(path)
    else:
        solver.tui.file.read_case(path)


def _read_via_settings(solver, file_type: str, path: str) -> None:
    """回退：使用 settings API 读取。"""
    if file_type == "mesh":
        solver.settings.file.read_mesh(file_name=path)
    else:
        solver.settings.file.read_case(file_name=path)


def read_mesh_file(solver, mesh_file: str) -> str:
    """根据扩展名判断文件类型，调用对应的 PyFluent 读取接口。

    返回描述读取通道的字符串，便于日志打印。
    """
    name = mesh_file.lower()
    path = mesh_file.replace("\\", "/")

    if name.endswith(".msh.h5") or name.endswith(".msh"):
        ftype = "mesh"
    elif name.endswith(".cas.h5") or name.endswith(".cas"):
        ftype = "case"
    else:
        raise ValueError(
            f"不支持的文件类型: {mesh_file}\n"
            f"仅支持: *.msh / *.msh.h5 / *.cas / *.cas.h5"
        )

    # 先尝试 tui 通道，失败再回退 settings 通道
    try:
        _read_via_tui(solver, ftype, path)
    except Exception as exc_tui:
        print(f"[提示] tui 读取失败，尝试 settings 通道: {exc_tui}")
        _read_via_settings(solver, ftype, path)

    label = {
        "mesh": "mesh (.msh / .msh.h5)",
        "case": "case (.cas / .cas.h5)",
    }[ftype]
    return label


# --------------------------------------------------------------------------
# 2. 利用 scheme_eval 获取网格质量信息（核心）
# --------------------------------------------------------------------------
# Fluent TUI 中 /mesh/quality 会打印完整的质量报告，
# /mesh/info 会打印网格规模（单元/面/节点数等）。
_QUALITY_CMD = '(ti-menu-load-string "/mesh/quality")'
_INFO_CMD = '(ti-menu-load-string "/mesh/info")'


def get_mesh_quality(solver) -> str:
    """利用 solver.scheme_eval 通道获取网格质量信息文本。

    步骤：
      ① 调用 solver.scheme_eval.scheme_eval(...) 触发网格质量报告
         （这是本工具按要求使用的核心入口）；
      ② 调用 solver.scheme_eval.exec(...) 获取 TUI 输出字符串。
         /mesh/quality 与 /mesh/info 分开执行：/mesh/info 在部分 Fluent
         版本中不存在（会报 Invalid command），失败时优雅跳过。
    """
    # ① 核心调用：solver.scheme.eval 触发质量报告（推荐通道，避免弃用告警）
    solver.scheme.eval(_QUALITY_CMD)

    # ② 分别执行，/mesh/info 失败时不影响主流程
    parts: list[str] = []
    try:
        parts.append(solver.scheme.exec([_QUALITY_CMD]) or "")
    except Exception as exc:
        print(f"[提示] /mesh/quality 文本捕获异常: {exc}")
    try:
        info = solver.scheme.exec([_INFO_CMD])
        if info:
            parts.append(info)
    except Exception:
        pass  # /mesh/info 在部分 Fluent 版本中不可用，静默忽略
    return "\n".join(p for p in parts if p.strip())


# --------------------------------------------------------------------------
# 3. 报告文本解析
# --------------------------------------------------------------------------
# 正则：匹配 Fluent /mesh/quality 输出中的常见关键指标。
_QUALITY_PATTERNS = {
    "min_orthogonal_quality": r"Minimum Orthogonal Quality\s*=\s*([-\d.eE+]+)",
    "max_orthogonal_quality": r"Maximum Orthogonal Quality\s*=\s*([-\d.eE+]+)",
    "min_aspect_ratio": r"Minimum Aspect Ratio\s*=\s*([-\d.eE+]+)",
    "max_aspect_ratio": r"Maximum Aspect Ratio\s*=\s*([-\d.eE+]+)",
    "max_cell_skewness": r"Maximum Cell Skewness\s*=\s*([-\d.eE+]+)",
    "max_face_skewness": r"Maximum Face Skewness\s*=\s*([-\d.eE+]+)",
    "max_cell_squish": r"Maximum Cell Squish\s*=\s*([-\d.eE+]+)",
    "min_cell_volume": r"Minimum Cell Volume\s*=\s*([-\d.eE+]+)",
}


def parse_quality(output: str) -> dict:
    """从网格质量报告文本中解析关键质量指标，返回 {指标名: 数值}。"""
    metrics: dict = {}
    for key, pat in _QUALITY_PATTERNS.items():
        m = re.search(pat, output)
        if m:
            try:
                metrics[key] = float(m.group(1))
            except ValueError:
                metrics[key] = m.group(1)
    return metrics


def parse_mesh_size(output: str) -> dict:
    """尝试从 /mesh/info 输出中解析网格规模（单元/面/节点数）。"""
    size: dict = {}
    pats = {
        "cells": r"(\d+)\s+cells",
        "faces": r"(\d+)\s+faces",
        "nodes": r"(\d+)\s+nodes",
    }
    for key, pat in pats.items():
        m = re.search(pat, output, re.IGNORECASE)
        if m:
            try:
                size[key] = int(m.group(1))
            except ValueError:
                pass
    return size


# --------------------------------------------------------------------------
# 4. 主流程
# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="基于 PyFluent 的网格质量检查工具"
    )
    parser.add_argument(
        "mesh_file",
        help="网格文件路径 (*.msh / *.msh.h5 / *.cas / *.cas.h5)",
    )
    parser.add_argument(
        "--dim", type=int, choices=[2, 3], default=None,
        help="网格维度（2 或 3）。不指定时由 fluent_dim_check.py 自动检测",
    )
    parser.add_argument(
        "--cores", type=int, default=4,
        help="Fluent 并行核数（默认 4）",
    )
    parser.add_argument(
        "--precision", default="double", choices=["single", "double"],
        help="求解精度（默认 double）",
    )
    parser.add_argument(
        "--json", default=None, metavar="FILE",
        help="质量结果保存为 JSON 的路径（默认: <网格名>_quality.json）",
    )
    args = parser.parse_args()

    if not os.path.exists(args.mesh_file):
        print(f"[错误] 网格文件不存在: {args.mesh_file}")
        sys.exit(1)

    print(f"[启动] 目标网格: {args.mesh_file}")

    # 维度：未显式指定时，调用 fluent_dim_check.detect_dimension() 获取结果
    if args.dim is None:
        try:
            args.dim = detect_dimension(args.mesh_file)
            print(f"[维度] fluent_dim_check 自动检测结果: {args.dim}D")
        except Exception as exc:
            print(f"[错误] 自动检测维度失败: {exc}")
            print("        请手动用 --dim 2 或 --dim 3 指定网格维度。")
            sys.exit(1)
        if args.dim not in (2, 3):
            print(f"[错误] 无法自动判断网格维度: {args.mesh_file}")
            sys.exit(1)
    else:
        print(f"[维度] 使用指定维度: {args.dim}D")

    # 启动 Fluent solver 会话
    launch_kwargs = dict(
        mode="solver",
        precision=args.precision,
        processor_count=args.cores,
        start_timeout=600,
        dimension=args.dim,
    )

    solver = pyfluent.launch_fluent(**launch_kwargs)
    fluent_version = None
    try:
        fluent_version = solver.get_fluent_version()
        print("Fluent version:", fluent_version)

        # ---- 读取网格 ----
        ftype = read_mesh_file(solver, args.mesh_file)
        print(f"[读取] 已载入（{ftype}）")

        # ---- 网格完整性检查（可选）----
        try:
            solver.settings.mesh.check()
            print("[检查] mesh.check() 完成")
        except Exception as exc:
            print(f"[提示] mesh.check() 跳过: {exc}")

        # ---- 利用 scheme_eval 获取网格质量信息 ----
        report = get_mesh_quality(solver)

        # ---- 输出 ----
        print("\n" + "=" * 64)
        print("网格质量报告（经 solver.scheme_eval.scheme_eval 获取）")
        print("=" * 64)
        if report.strip():
            print(report)
        else:
            print("[提示] 未捕获到文本输出，请直接查看 Fluent 控制台中的 "
                  "Mesh Quality 信息。")

        # ---- 解析关键指标 ----
        metrics = parse_quality(report)
        size = parse_mesh_size(report)
        if metrics:
            print("\n--- 解析得到的关键质量指标 ---")
            for k, v in metrics.items():
                print(f"  {k:28s}: {v}")
        if size:
            print("\n--- 网格规模 ---")
            for k, v in size.items():
                print(f"  {k:8s}: {v}")

        # ---- 保存网格质量结果到 JSON ----
        # fluent_version 可能为非 JSON 可序列化对象，统一转成字符串
        result = {
            "path": os.path.abspath(args.mesh_file),
            "dimension": f"{args.dim}d",
            "fluent_version": str(fluent_version) if fluent_version is not None else None,
            "min_orthogonal_quality": metrics.get("min_orthogonal_quality"),
            "max_aspect_ratio": metrics.get("max_aspect_ratio"),
            "extra_metrics": {
                k: v for k, v in metrics.items()
                if k not in ("min_orthogonal_quality", "max_aspect_ratio")
            },
            "mesh_size": size,
        }
        if args.json:
            json_path = args.json
        else:
            base = os.path.splitext(os.path.abspath(args.mesh_file))[0]
            json_path = base + "_quality.json"
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"\n[输出] 网格质量结果已保存: {json_path}")
        except Exception as exc:
            print(f"[错误] 保存 JSON 失败: {exc}")

    finally:
        solver.exit()


if __name__ == "__main__":
    main()
