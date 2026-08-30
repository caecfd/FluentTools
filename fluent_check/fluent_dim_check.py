#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fluent_dim_check.py — Fluent 网格维度检测工具
====================================================================

功能：
    不启动 Fluent、不依赖 PyFluent，直接判断网格文件的维度（2D / 3D）。
    支持的格式：
        *.msh       （纯网格，文本格式）
        *.msh.h5    （纯网格，HDF5 格式）
        *.cas       （算例，文本格式）
        *.cas.h5    （算例，HDF5 格式）

检测策略（按优先级）：
    1. 文本格式 (.cas / .msh)：解析文件头中的维度声明 (2 2) / (2 3)；
    2. HDF5 格式 (*.h5)：读取 /meshes/<id> 组的 'dimension' 属性（需 h5py）；
    3. 上述均失败时回退 pyvista/vtk，按单元类型判断。

对外接口：
    detect_dimension(file_path) -> int      返回 2 或 3，失败抛异常
    可被 fluent_mesh_quality_check.py 等模块直接导入调用。

用法（命令行）：
    python fluent_dim_check.py <网格文件路径>

示例：
    python fluent_dim_check.py VM02.msh
        当前网格：C:\\...\\VM02.msh
        网格维度：2d
"""
import argparse
import os
import re
import sys

# 统一标准输出编码，避免中文/特殊字符在 Windows 控制台乱码
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# --------------------------------------------------------------------------
# 1. 文本格式 (.cas / .msh) 维度解析
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


# --------------------------------------------------------------------------
# 2. HDF5 格式 (*.cas.h5 / *.msh.h5) 维度解析
# --------------------------------------------------------------------------
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


# --------------------------------------------------------------------------
# 3. 对外统一入口
# --------------------------------------------------------------------------
def detect_dimension(file_path: str) -> int:
    """自动判断网格维度，返回 2 或 3。

    对于传统 .cas / .msh 文本文件与 *.h5 文件，直接读取文件内容判断维度，
    无需 pyvista；解析失败时回退 pyvista 检查单元类型判断。
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

    # 三维单元类型的 VTK id 集合
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
# 4. 命令行入口
# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Fluent 网格维度检测工具（不启动 Fluent）"
    )
    parser.add_argument(
        "mesh_file",
        help="网格文件路径 (*.msh / *.msh.h5 / *.cas / *.cas.h5)",
    )
    args = parser.parse_args()

    if not os.path.exists(args.mesh_file):
        print(f"[错误] 网格文件不存在: {args.mesh_file}")
        sys.exit(1)

    try:
        dim = detect_dimension(args.mesh_file)
    except Exception as exc:
        print(f"[错误] 自动检测维度失败: {exc}")
        sys.exit(1)

    if dim not in (2, 3):
        print(f"[错误] 无法自动判断网格维度: {args.mesh_file}")
        sys.exit(1)

    print(f"当前网格：{os.path.abspath(args.mesh_file)}")
    print(f"网格维度：{dim}d")


if __name__ == "__main__":
    main()
