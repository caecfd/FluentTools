#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fluent_mesh_quality_check.py — 基于 PyFluent 的网格质量检查工具
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
    普通模式：省略 --dim 时自动检测（基于文件内容判断，无需 pyvista）；
    也可手动用 --dim 2 / --dim 3 指定。
    仅检测模式：加 --d 参数可只输出网格维度（如 "2d"）后退出，
    不启动 Fluent、不写入 JSON，适合快速查看维度。

用法：
    python fluent_mesh_quality_check.py <网格文件路径> [--dim 2|3] [--cores N] [--precision single|double] [--json FILE] [--d]

示例：
    python fluent_mesh_quality_check.py test.cas.h5
    python fluent_mesh_quality_check.py VM02.msh
    python fluent_mesh_quality_check.py PTE.cas.h5 --cores 8 --json result.json
    python fluent_mesh_quality_check.py VM02.msh --dim 2
    python fluent_mesh_quality_check.py VM02.msh --d        # 仅输出维度：当前网格 / 网格维度：2d
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


# --------------------------------------------------------------------------
# 0. 网格维度自动检测（移植自 grid_detect.py 的 --d 功能）
# --------------------------------------------------------------------------
def _detect_dim_from_text(file_path: str):
    """从传统文本网格文件 (.cas / .msh) 解析维度，无需 pyvista / 无需启动 Fluent。

    Fluent 文本文件开头通常包含维度声明：
      - .cas:  (0 "Dimension:") 换行 (2 2)  或  (2 3)
      - .msh:  直接以 (2 2) / (2 3) 出现在文件前部（无 Dimension 头）
    解析成功返回 2 / 3，失败返回 None。
    """
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as fh:
            found_header = False
            for lineno, line in enumerate(fh, 1):
                s = line.strip()
                if not s:
                    continue
                if found_header:
                    # 已找到 Dimension 头，下一行即为维度值
                    if re.search(r'\(2\s+2\s*\)', s):
                        return 2
                    if re.search(r'\(2\s+3\s*\)', s):
                        return 3
                    # 遇到下一个块头仍未找到，停止扫描
                    if s.startswith("(0"):
                        break
                    continue
                # 兼容 "Dimension:" 与 "Dimensions:" 头
                if re.search(r'\(0\s*"dimensions?', line, re.IGNORECASE):
                    found_header = True
                    continue
                # 无 Dimension 头时（如 .msh），在文件前部直接查找 (2 2)/(2 3)
                if lineno <= 20:
                    if re.search(r'\(2\s+2\s*\)', s):
                        return 2
                    if re.search(r'\(2\s+3\s*\)', s):
                        return 3
    except Exception:
        return None
    return None


def _detect_dim_from_h5(file_path: str):
    """从 Fluent HDF5 网格文件 (.cas.h5 / .msh.h5) 读取维度，无需 pyvista。

    维度信息存储在 /meshes/<id> 组的 'dimension' 属性中（如 [2] 或 [3]）。
    解析成功返回 2 / 3，失败返回 None。
    """
    try:
        import h5py
    except Exception:
        return None
    try:
        with h5py.File(file_path, "r") as f:
            if "meshes" not in f:
                return None
            meshes = f["meshes"]
            for name in meshes:
                grp = meshes[name]
                if isinstance(grp, h5py.Group) and "dimension" in grp.attrs:
                    val = grp.attrs["dimension"]
                    try:
                        if hasattr(val, "__len__"):
                            return int(val[0])
                        return int(val)
                    except Exception:
                        return None
    except Exception:
        return None
    return None


def detect_dimension(file_path: str) -> int:
    """自动判断网格维度，返回 2 或 3（移植自 grid_detect.py 的 mesh_dimension）。

    对于传统 .cas（非 HDF5）文本文件，直接读取文件内容判断维度，
    无需 pyvista；其余格式仍通过 pyvista 检查单元类型判断。
    """
    # 传统文本格式 (.cas / .msh)：从文件头解析维度，避免依赖 pyvista
    ext = file_path.lower()
    if ext.endswith(".cas") or ext.endswith(".msh"):
        dim = _detect_dim_from_text(file_path)
        if dim is not None:
            return dim
        # 解析失败则回退到下面的 pyvista 流程
    elif ext.endswith(".h5"):
        # .cas.h5 / .msh.h5：从 HDF5 属性直接读取维度，避免依赖 pyvista
        dim = _detect_dim_from_h5(file_path)
        if dim is not None:
            return dim
        # 解析失败则回退到下面的 pyvista 流程
    try:
        import vtk
        # 回退路径使用 pyvista/vtk 读取，可能打印无关 ERROR 日志，予以静音
        vtk.vtkLogger.SetStderrVerbosity(vtk.vtkLogger.VERBOSITY_OFF)
    except Exception:
        pass
    import numpy as np
    import pyvista as pv

    reader = pv.get_reader(file_path)
    data = reader.read()
    blocks = list(data) if isinstance(data, pv.MultiBlock) else [data]

    # 三维单元类型的 VTK id 集合（与 grid_detect.py 一致）
    _3d_ids = {
        int(pv.CellType.TETRA),
        int(pv.CellType.PYRAMID),
        int(pv.CellType.WEDGE),
        int(pv.CellType.HEXAHEDRON),
        int(pv.CellType.POLYHEDRON),
    }
    for block in blocks:
        if isinstance(block, pv.UnstructuredGrid):
            types = np.asarray(block.celltypes)
            if np.any(np.isin(types, list(_3d_ids))):
                return 3
    return 2


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
        help="网格维度（2 或 3）。不指定时自动检测（需 pyvista）",
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
    parser.add_argument(
        "--d", action="store_true",
        help="仅检测并输出网格维度（2 或 3）后退出，不启动 Fluent、不写入 JSON",
    )
    args = parser.parse_args()

    if not os.path.exists(args.mesh_file):
        print(f"[错误] 网格文件不存在: {args.mesh_file}")
        sys.exit(1)

    # 仅检测维度模式（--d）：不启动 Fluent，不写入 JSON，仅输出维度信息
    if args.d:
        dim = detect_dimension(args.mesh_file)
        if dim not in (2, 3):
            print(f"[错误] 无法自动判断网格维度: {args.mesh_file}")
            sys.exit(1)
        print(f"当前网格：{os.path.abspath(args.mesh_file)}")
        print(f"网格维度：{dim}d")
        sys.exit(0)

    print(f"[启动] 目标网格: {args.mesh_file}")

    # 维度：未显式指定时自动检测（移植自 grid_detect.py 的 --d 功能）
    if args.dim is None:
        try:
            args.dim = detect_dimension(args.mesh_file)
            print(f"[维度] 自动检测结果: {args.dim}D")
        except Exception as exc:
            print(f"[错误] 自动检测维度失败: {exc}")
            print("        请手动用 --dim 2 或 --dim 3 指定网格维度。")
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
