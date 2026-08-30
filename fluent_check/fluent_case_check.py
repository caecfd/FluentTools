#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fluent_case_check.py — 基于 PyFluent 读取 *.cas / *.cas.h5 并输出用户设置参数
================================================================================

功能：
    1. 启动 Fluent solver 会话，读入 Fluent 算例文件（*.cas / *.cas.h5）；
    2. 提取并输出用户在算例中设置的内容，分为两部分：
       （一）参数化参数（Parameters & Customization）：
        a. 输入参数 (Input Parameters)
        b. 输出参数 (Output Parameters)
        c. 命名表达式 (Named Expressions)
       （二）求解器设置（用户配置的模型与边界等）：
        d. 边界条件 (Boundary Conditions) —— 各区域类型与设置
        e. 材料 (Materials) —— fluid / solid / mixture 的属性
        f. 求解模型 (Models) —— viscous / energy / multiphase 等模型选择
        g. 求解设置 (Solution) —— 求解方法 / 求解控制 / 初始化
    3. 结果可打印到控制台，也可保存为 JSON；指定 --json 时会同时输出同名 .toml 文件。

核心 API 说明：
    - 维度判定复用 fluent_dim_check.detect_dimension()（不依赖 pyvista 时优先，
      失败回退 pyvista）。
    - 参数读取走 settings API（不同 Fluent 版本路径略有差异，已做多路径兼容）：
        * 输入/输出参数（Fluent 2026 起为 Group 容器，下分 expression /
          scheme_proc / udf_side 等子组，已兼容）：
            solver.settings.parameters.input_parameters
        * 命名表达式：
            solver.setup.named_expressions  /  solver.settings.setup.named_expressions
      每个输入/输出参数对象提供 .value（数值）、.units（单位）等属性；
      每个命名表达式对象提供 .definition（表达式字符串）、.input_parameter、
      .output_parameter（是否作为参数）等属性。
    - 求解器设置（边界条件/材料/模型/求解）通过各 settings 对象的 get_state()
      读取，并递归序列化为 JSON 友好的结构（Quantity -> "值 [单位]" 等）。

用法：
    python fluent_case_check.py <算例文件路径> [--dim 2|3] [--cores N] [--precision single|double] [--json FILE]

示例：
    python fluent_case_check.py VM02.cas
    python fluent_case_check.py elbow_param.cas.h5 --cores 8 --json params.json
    python fluent_case_check.py VM02.cas --dim 2
"""
import argparse
import json
import os
import sys

# 统一标准输出编码，避免中文/特殊字符在 Windows 控制台乱码
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import ansys.fluent.core as pyfluent

# 网格维度判断由独立模块 fluent_dim_check.py 提供（与 fluent_mesh_quality_check.py 一致）
try:
    from fluent_dim_check import detect_dimension
except ImportError:  # 以其他工作目录运行时，确保能找到同目录模块
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from fluent_dim_check import detect_dimension


# --------------------------------------------------------------------------
# 1. 文件类型识别与算例读取
# --------------------------------------------------------------------------
def _case_file_type(mesh_file: str) -> str:
    """根据扩展名判断是否为受支持的算例文件，返回 "case"。

    仅支持 *.cas / *.cas.h5，其它扩展名抛 ValueError。
    """
    name = mesh_file.lower()
    if name.endswith(".cas.h5") or name.endswith(".cas"):
        return "case"
    raise ValueError(
        f"不支持的文件类型: {mesh_file}\n"
        f"仅支持: *.cas / *.cas.h5"
    )


def read_case_file(solver, case_file: str) -> None:
    """调用 PyFluent 读入算例文件（优先 tui 通道，失败回退 settings 通道）。"""
    path = case_file.replace("\\", "/")
    try:
        solver.tui.file.read_case(path)
    except Exception as exc_tui:
        print(f"[提示] tui 读取失败，尝试 settings 通道: {exc_tui}")
        solver.settings.file.read_case(file_name=path)


# --------------------------------------------------------------------------
# 2. settings 对象访问辅助
# --------------------------------------------------------------------------
def _resolve_group(solver, candidates: list):
    """依次尝试候选路径，返回第一个可访问的非空 settings 组；都失败返回 None。"""
    for cand in candidates:
        try:
            grp = cand(solver)
            if grp is not None:
                return grp
        except Exception:
            continue
    return None


def _iter_named_object(obj):
    """遍历一个 PyFluent 参数容器，返回 [(名称, 子对象), ...]。

    兼容两种结构：
      * 直接的 NamedObject（如 named_expressions、旧版 input_parameters）：
        可直接 list(obj) 得到子项名称；
      * Fluent 2026 的 Group 容器（如 input_parameters / output_parameters）：
        其本身就是 Group，子项按类别分在子组里（input_parameters 下有
        expression / scheme_proc / udf_side；output_parameters 下有
        report_definitions / named_expressions），需先遍历子组再枚举。
    名称统一用 "子组/参数名" 形式，便于在 Group 结构下区分来源。
    """
    # 先尝试直接枚举（NamedObject）
    try:
        names = list(obj)
        if names:
            return [(n, obj[n]) for n in names]
    except Exception:
        pass
    # 否则当作 Group 处理：遍历其子组（child_names），再枚举每个子组
    try:
        subs = obj.child_names
    except Exception:
        return []
    items = []
    for sub in subs:
        try:
            subobj = getattr(obj, sub)
        except Exception:
            continue
        try:
            for n in list(subobj):
                try:
                    items.append((f"{sub}/{n}", subobj[n]))
                except Exception:
                    continue
        except Exception:
            continue
    return items


def _get_attr(child, *attrs):
    """依次尝试读取 child 的若干属性/方法，返回第一个成功的值；都失败返回 None。"""
    for a in attrs:
        try:
            v = getattr(child, a)
            if callable(v):
                try:
                    v = v()
                except Exception:
                    continue
            return v
        except Exception:
            continue
    return None


def _stringify(v):
    """将 PyFluent 返回的值（可能为 Real / Quantity / 字符串 / 布尔）转为可序列化内容。"""
    if v is None:
        return None
    # 优先尝试数值（含 Quantity 的量值）
    try:
        return float(v)
    except (TypeError, ValueError):
        pass
    # Quantity 风格（含 .magnitude / .units）
    try:
        return f"{v.magnitude} [{v.units}]"
    except Exception:
        return str(v)


def _to_serializable(v):
    """递归地将 PyFluent 的 settings 状态（get_state 返回值）转为 JSON 可序列化结构。

    处理：dict / list / tuple 递归；Quantity -> "值 [单位]"；Real -> float；
    枚举 -> 字符串；其余 -> str。
    """
    import numbers
    if v is None:
        return None
    if isinstance(v, dict):
        return {str(k): _to_serializable(val) for k, val in v.items()}
    if isinstance(v, (list, tuple)):
        return [_to_serializable(x) for x in v]
    # Quantity 风格（含 .magnitude / .units）
    try:
        return f"{v.magnitude} [{v.units}]"
    except Exception:
        pass
    if isinstance(v, bool):
        return v
    if isinstance(v, numbers.Real):
        try:
            return float(v)
        except Exception:
            return str(v)
    if isinstance(v, str):
        return v
    try:
        return str(v)
    except Exception:
        return repr(v)


def _safe_state(obj):
    """对某 settings 对象调用 get_state() 并递归序列化；失败返回错误描述字符串。"""
    try:
        return _to_serializable(obj.get_state())
    except Exception as exc:
        return f"<无法读取状态: {type(exc).__name__}: {exc}>"


# --------------------------------------------------------------------------
# 3. 参数提取
# --------------------------------------------------------------------------
def get_input_parameters(solver) -> dict:
    """读取输入参数，返回 {参数名: 描述字典}。"""
    grp = _resolve_group(solver, [
        lambda s: s.settings.parameters.input_parameters,
        lambda s: s.settings.parameter_workspace.parameters.input_parameters,
        lambda s: s.setup.parameters.input_parameters,
        lambda s: s.parameters.input_parameters,
    ])
    result = {}
    if grp is None:
        return result
    for name, child in _iter_named_object(grp):
        val = _get_attr(child, "value", "definition", "expression")
        units = _get_attr(child, "units")
        result[name] = {
            "value": _stringify(val),
            "units": _stringify(units) if units is not None else None,
        }
    return result


def get_output_parameters(solver) -> dict:
    """读取输出参数，返回 {参数名: 描述字典}。"""
    grp = _resolve_group(solver, [
        lambda s: s.settings.parameters.output_parameters,
        lambda s: s.settings.parameter_workspace.parameters.output_parameters,
        lambda s: s.setup.parameters.output_parameters,
        lambda s: s.parameters.output_parameters,
    ])
    result = {}
    if grp is None:
        return result
    for name, child in _iter_named_object(grp):
        val = _get_attr(child, "value", "definition", "expression")
        units = _get_attr(child, "units")
        result[name] = {
            "value": _stringify(val),
            "units": _stringify(units) if units is not None else None,
        }
    return result


def get_named_expressions(solver) -> dict:
    """读取命名表达式，返回 {名称: 描述字典}。"""
    grp = _resolve_group(solver, [
        lambda s: s.settings.setup.named_expressions,
        lambda s: s.setup.named_expressions,
        lambda s: s.settings.named_expressions,
    ])
    result = {}
    if grp is None:
        return result
    for name, child in _iter_named_object(grp):
        definition = _get_attr(child, "definition")
        is_input = _get_attr(child, "input_parameter")
        is_output = _get_attr(child, "output_parameter")
        result[name] = {
            "definition": definition if definition is not None else None,
            "input_parameter": bool(is_input) if is_input is not None else None,
            "output_parameter": bool(is_output) if is_output is not None else None,
        }
    return result


# --------------------------------------------------------------------------
# 3b. 求解器设置导出（边界条件 / 材料 / 模型 / 求解方法）
# --------------------------------------------------------------------------
def get_boundary_conditions(solver) -> dict:
    """导出所有边界条件区域及其设置（按边界类型分组）。

    返回 {边界类型: {区域名: 设置字典}}。某类无区域或不可枚举时自动跳过。
    """
    bc = _resolve_group(solver, [
        lambda s: s.settings.setup.boundary_conditions,
        lambda s: s.setup.boundary_conditions,
    ])
    result = {}
    if bc is None:
        return result
    for btype in getattr(bc, "child_names", []) or []:
        try:
            grp = getattr(bc, btype)
            zones = list(grp)
        except Exception:
            continue
        if not zones:
            continue
        zones_out = {}
        for z in zones:
            state = _safe_state(grp[z])
            # 在每个区域的设置中显式标注其边界类型
            if isinstance(state, dict):
                state = {"boundary_type": btype, **state}
            zones_out[z] = state
        result[btype] = zones_out
    return result


def get_materials(solver) -> dict:
    """导出材料属性（fluid / solid / mixture）。

    返回 {材料类别: {材料名: 属性字典}}。
    """
    mat = _resolve_group(solver, [
        lambda s: s.settings.setup.materials,
        lambda s: s.setup.materials,
    ])
    result = {}
    if mat is None:
        return result
    for mtype in ("fluid", "solid", "mixture"):
        try:
            grp = getattr(mat, mtype)
            items = list(grp)
        except Exception:
            continue
        if items:
            result[mtype] = {n: _safe_state(grp[n]) for n in items}
    return result


def get_models(solver) -> dict:
    """导出启用的求解模型选择（viscous / energy / multiphase 等）。"""
    models = _resolve_group(solver, [
        lambda s: s.settings.setup.models,
        lambda s: s.setup.models,
    ])
    if models is None:
        return {}
    return _safe_state(models)


def get_solution_settings(solver) -> dict:
    """导出求解设置（求解方法 / 求解控制 / 初始化 / 迭代设置）。"""
    out = {}
    sol = _resolve_group(solver, [
        lambda s: s.settings.solution,
        lambda s: s.solution,
    ])
    if sol is None:
        return out
    for key, attr in (("methods", "methods"), ("controls", "controls"),
                      ("initialization", "initialization"),
                      ("run_calculation", "run_calculation")):
        try:
            obj = getattr(sol, attr)
            out[key] = _safe_state(obj)
        except Exception:
            continue
    return out


# --------------------------------------------------------------------------
# 4. 主流程
# --------------------------------------------------------------------------
def _sanitize_for_toml(value):
    """递归清理为 TOML 可序列化结构：丢弃 None，转换 numpy / 未知类型为字符串。

    - 列表若同时含 dict 与标量，则把标量项包成 {"value": ...} 以兼容 TOML 数组表；
    - 非有限浮点（nan / inf）归零，避免写出非法 TOML。
    """
    import math
    if isinstance(value, dict):
        out = {}
        for k, v in value.items():
            sv = _sanitize_for_toml(v)
            if sv is None:
                continue
            out[k] = sv
        return out
    if isinstance(value, (list, tuple)):
        if not value:
            return []
        if any(isinstance(x, (dict, list)) for x in value):
            items = []
            for x in value:
                sx = _sanitize_for_toml(x)
                if isinstance(sx, dict):
                    items.append(sx)
                else:
                    items.append({"value": sx})
            return items
        return [_sanitize_for_toml(x) for x in value]
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, str)):
        if isinstance(value, float) and not isinstance(value, bool):
            return value if math.isfinite(value) else 0.0
        return value
    # numpy / 其他数值类型
    try:
        import numbers
        if isinstance(value, numbers.Integral):
            return int(value)
        if isinstance(value, numbers.Real):
            f = float(value)
            return f if math.isfinite(f) else 0.0
    except Exception:
        pass
    return str(value)


def save_toml_file(path: str, data: dict) -> bool:
    """将结果以同名 TOML 写出（需要 tomli_w；未安装时优雅跳过）。"""
    try:
        import tomli_w
    except ImportError:
        print("[警告] 未安装 tomli_w，跳过 TOML 输出（可运行：pip install tomli_w）")
        return False
    clean = _sanitize_for_toml(data)
    try:
        with open(path, "wb") as f:
            tomli_w.dump(clean, f)
        return True
    except Exception as exc:
        print(f"[警告] 写入 TOML 失败: {exc}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description="基于 PyFluent 读取算例并输出用户设置参数"
    )
    parser.add_argument(
        "case_file",
        help="算例文件路径 (*.cas / *.cas.h5)",
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
        help="参数结果保存为 JSON 的路径（默认不保存）",
    )
    args = parser.parse_args()

    if not os.path.exists(args.case_file):
        print(f"[错误] 算例文件不存在: {args.case_file}")
        sys.exit(1)

    # 文件类型校验（仅 *.cas / *.cas.h5）
    try:
        _case_file_type(args.case_file)
    except ValueError as exc:
        print(f"[错误] {exc}")
        sys.exit(1)

    print(f"[启动] 目标算例: {args.case_file}")

    # 维度：未显式指定时，调用 fluent_dim_check.detect_dimension() 获取结果
    if args.dim is None:
        try:
            args.dim = detect_dimension(args.case_file)
            print(f"[维度] fluent_dim_check 自动检测结果: {args.dim}D")
        except Exception as exc:
            print(f"[错误] 自动检测维度失败: {exc}")
            print("        请手动用 --dim 2 或 --dim 3 指定网格维度。")
            sys.exit(1)
        if args.dim not in (2, 3):
            print(f"[错误] 无法自动判断网格维度: {args.case_file}")
            sys.exit(1)
    else:
        print(f"[维度] 使用指定维度: {args.dim}D")

    # 启动 Fluent solver 会话（复用与 fluent_mesh_quality_check.py 一致的启动方式）
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

        # ---- 读入算例 ----
        read_case_file(solver, args.case_file)
        print(f"[读取] 已载入算例: {args.case_file}")

        # ---- 提取用户输入参数 ----
        input_params = get_input_parameters(solver)
        output_params = get_output_parameters(solver)
        named_exprs = get_named_expressions(solver)

        # ---- 提取求解器设置（边界条件 / 材料 / 模型 / 求解方法）----
        boundary_conditions = get_boundary_conditions(solver)
        materials = get_materials(solver)
        models = get_models(solver)
        solution_settings = get_solution_settings(solver)

        # ---- 控制台输出 ----
        print("\n" + "=" * 64)
        print("算例中用户设置的参数（经 solver settings API 读取）")
        print("=" * 64)

        print("\n--- 输入参数 (Input Parameters) ---")
        if input_params:
            for k, v in input_params.items():
                unit = f" {v['units']}" if v.get("units") else ""
                print(f"  {k:28s}: {v['value']}{unit}")
        else:
            print("  （无 / 未能读取）")

        print("\n--- 输出参数 (Output Parameters) ---")
        if output_params:
            for k, v in output_params.items():
                unit = f" {v['units']}" if v.get("units") else ""
                print(f"  {k:28s}: {v['value']}{unit}")
        else:
            print("  （无 / 未能读取）")

        print("\n--- 命名表达式 (Named Expressions) ---")
        if named_exprs:
            for k, v in named_exprs.items():
                tags = []
                if v.get("input_parameter"):
                    tags.append("input")
                if v.get("output_parameter"):
                    tags.append("output")
                tag_str = f" [{'/'.join(tags)}]" if tags else ""
                print(f"  {k:28s}{tag_str}: {v['definition']}")
        else:
            print("  （无 / 未能读取）")

        # ---- 边界条件 ----
        print("\n--- 边界条件 (Boundary Conditions) ---")
        if boundary_conditions:
            for btype, zones in boundary_conditions.items():
                print(f"  [{btype}]")
                for zname, zstate in zones.items():
                    print(f"    - {zname}: "
                          + json.dumps(zstate, ensure_ascii=False))
        else:
            print("  （无 / 未能读取）")

        # ---- 材料 ----
        print("\n--- 材料 (Materials) ---")
        if materials:
            for mtype, items in materials.items():
                for mname, mstate in items.items():
                    print(f"  - {mtype}/{mname}: "
                          + json.dumps(mstate, ensure_ascii=False))
        else:
            print("  （无 / 未能读取）")

        # ---- 求解模型 ----
        print("\n--- 求解模型 (Models) ---")
        if models:
            print("  " + json.dumps(models, ensure_ascii=False))
        else:
            print("  （无 / 未能读取）")

        # ---- 求解设置 ----
        print("\n--- 求解设置 (Solution: methods/controls/initialization) ---")
        if solution_settings:
            print("  " + json.dumps(solution_settings, ensure_ascii=False))
        else:
            print("  （无 / 未能读取）")

        # ---- 保存参数结果到 JSON（如有需要）----
        result = {
            "path": os.path.abspath(args.case_file),
            "dimension": f"{args.dim}d",
            "fluent_version": str(fluent_version) if fluent_version is not None else None,
            "input_parameters": input_params,
            "output_parameters": output_params,
            "named_expressions": named_exprs,
            "boundary_conditions": boundary_conditions,
            "materials": materials,
            "models": models,
            "solution_settings": solution_settings,
        }
        if args.json:
            try:
                with open(args.json, "w", encoding="utf-8") as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                print(f"\n[输出] 参数结果已保存: {args.json}")
            except Exception as exc:
                print(f"[错误] 保存 JSON 失败: {exc}")
            # 输出同名 TOML 文件
            toml_path = os.path.splitext(args.json)[0] + ".toml"
            if save_toml_file(toml_path, result):
                print(f"[输出] 参数结果已保存: {toml_path}")

    finally:
        solver.exit()


if __name__ == "__main__":
    main()
