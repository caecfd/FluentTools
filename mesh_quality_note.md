# mesh_quality.py — Fluent 网格质量分析工具

基于 [PyVista](https://dev.pyvista.org/) 的命令行工具，直接读取 ANSYS Fluent
网格文件（CFF 格式 `.cas.h5` / `.msh.h5`，以及传统 `.cas` / `.msh`），计算单元
质量指标并输出报告。无需安装 Fluent。

除 VTK/Verdict 内置指标外，工具自行实现了两个 **Fluent 定义的指标**
（VTK 未提供），数值与 Fluent 网格检查结果一致：

| 指标 | 含义 | 取值范围 |
|---|---|---|
| `fluent_orthogonal` | 正交质量（Orthogonal Quality）：各面上「面法向与质心连线」夹角余弦的最小值（含邻居单元项），2D 用边 | 0（最差）~ 1（完美） |
| `fluent_aspect_ratio` | 长宽比（Aspect Ratio）：「质心到面/边质心的法向距离」与「质心到节点距离」中的最大值 / 最小值 | ≥ √2（正方形）/ √3（立方体），越大越差 |

## 环境要求

- Python ≥ 3.9
- [pyvista](https://dev.pyvista.org/) ≥ 0.45（开发环境：0.48.4 + VTK 9.6.2；
  直接读取 `.cas.h5` 需要 VTK ≥ 9.6 的 `FLUENTCFFReader`）
- numpy
- matplotlib（仅 `--hist` 需要）

```bash
pip install pyvista numpy matplotlib
```

## 快速开始

```bash
# Fluent 网格检查：最小面面积 / 最小正交质量 / 最大长宽比
python mesh_quality.py --check

# 默认指标的质量报告（文件缺省为 test.cas.h5）
python mesh_quality.py

# 指定文件与指标
python mesh_quality.py case.cas.h5 --measures skew shape fluent_aspect_ratio
```

`--check` 输出示例（test.cas.h5，与 Fluent 2026 R1 网格检查结果一致）：

```text
==========================================================================
 Fluent mesh check: test.cas.h5  (reader: FLUENTCFFReader)
==========================================================================

 Block: block_0   Cells: 1600   2D
   Minimum face area            = 6.663859e-03
   Minimum orthogonal quality   = 1.000000e+00
   Maximum aspect ratio         = 6.670067e+00
```

默认质量报告输出示例：

```text
File: test.cas.h5
Reader: FLUENTCFFReader   Blocks: 1

==========================================================================
 Block: block_0
==========================================================================
 Cells: 1600   Points: 1681   Type(s): QUAD x1600   Dimension: 2D
 Total area: 1   Total volume: 0

 measure                      min         max        mean         std         p99       outside
 ----------------------------------------------------------------------------------------------
 fluent_orthogonal              1           1           1    6.13e-11           1     0 (0.00%)
 fluent_aspect_ratio        1.417        6.67       2.378       1.055       5.872     0 (0.00%)
 skew                           0   2.371e-05   2.275e-06   3.008e-06    1.32e-05     0 (0.00%)
 ...
 worst [fluent_aspect_ratio]: #1580=6.67, #1579=6.67, #1581=6.584, ...
```

其中 `outside` 列为超出「可接受范围」的单元数及占比；`worst [...]` 列出该
指标最差的 5 个单元（全部合格时显示 `all cells within acceptable range`，
或按「偏离理想值程度」给出极端单元）。

## 命令行参数

| 参数 | 说明 |
|---|---|
| `file_path` | 网格文件路径，支持 `.cas.h5` / `.msh.h5` / `.cas`；缺省为 `test.cas.h5` |
| `--check` | 只输出 Fluent 网格检查三项指标（最小面面积、最小正交质量、最大长宽比）后退出 |
| `--measures M [M ...]` | 要计算的质量指标，见下文指标列表；缺省为 `fluent_orthogonal fluent_aspect_ratio skew scaled_jacobian aspect_ratio shape area` |
| `--all` | 自动计算该网格所有可用的指标（含两个 Fluent 自定义指标） |
| `--csv [FILE]` | 逐单元质量值写入 CSV（缺省文件名 `mesh_quality.csv`） |
| `--json [FILE]` | 汇总统计写入 JSON（缺省文件名 `mesh_quality.json`） |
| `--hist [FILE]` | 各指标直方图保存为 PNG（缺省文件名 `mesh_quality.png`），并标注可接受范围 |
| `--show` | 打开 3D 视图，超出可接受范围的单元标红 |

示例：

```bash
python mesh_quality.py --all --csv quality.csv --json quality.json --hist quality.png
python mesh_quality.py case.cas.h5 --measures fluent_orthogonal skew --show
```

## 质量指标

### Fluent 定义指标（本工具实现）

- **`fluent_orthogonal`**：对单元的每个面 *f*，取外法向面积向量 `A_f`、
  单元质心指向面质心的向量 `c_f`、面质心指向邻居单元质心的向量 `r_f`，
  单元值 = 各面 `min(A_f·c_f/|A_f||c_f|, A_f·r_f/|A_f||r_f|)` 的最小值
  （边界面只用第一项）。可接受范围 `[0.15, 1]`（Fluent 经验：最低正交质量
  应大于约 0.15）。
- **`fluent_aspect_ratio`**：质心到面/边质心的法向距离、质心到节点的距离
  两类距离中的最大值与最小值之比。可接受范围 `[1, 10]`（工程惯例）。
  注意完美单元的定义下限不为 1：正方形 √2、立方体 √3、正三角形 2、
  正四面体 3。

### VTK/Verdict 指标（`cell_quality`）

`--measures` 可选的 Verdict 指标：`area`、`aspect_frobenius`、`aspect_gamma`、
`aspect_ratio`、`collapse_ratio`、`condition`、`diagonal`、`dimension`、
`distortion`、`jacobian`、`max_angle`、`max_aspect_frobenius`、
`max_edge_ratio`、`med_aspect_frobenius`、`min_angle`、`oddy`、
`radius_ratio`、`relative_size_squared`、`scaled_jacobian`、`shape`、
`shape_and_size`、`shear`、`shear_and_size`、`skew`、`stretch`、`taper`、
`volume`、`warpage`。

「可接受范围」来自 `pyvista.cell_quality`（Verdict 标准值，如 quad 的
`skew ∈ [0, 0.5]`、`aspect_ratio ∈ [1, 1.3]`）。单元类型不支持某指标时
（如 quad 的 `volume`），相应指标自动跳过并提示。

### 注意：不同口径的 "aspect ratio" 数值不同

以矩形（边长比 r）为例，三种口径结果不同，均非计算错误：

| r | Fluent 定义（`fluent_aspect_ratio`）= √(r²+1) | Verdict `aspect_ratio` = (r+1)/2 | Verdict `max_edge_ratio` = r |
|---|---|---|---|
| 1 | 1.414 | 1.0 | 1.0 |
| 2 | 2.236 | 1.5 | 2.0 |
| 6.5946 | **6.670**（= Fluent 显示值） | 3.797 | 6.595 |

需要与 Fluent 对齐时请使用 `fluent_aspect_ratio`。同理，VTK 的 `skew` 也
不等价于 Fluent 的 skewness。

## 输出文件

- **CSV**：每行一个单元，列为 `block, cell_id, cell_type, <各指标>`；
  不适用的指标留空。
- **JSON**：`file` / `reader` / `blocks`，每个 block 含单元统计
  （min/max/mean/std/p99、可接受范围外计数）与 worst 单元列表。
- **直方图 PNG**：每指标一个子图，橙色虚线为可接受范围边界。
- **`--show`**：PyVista 3D 窗口，蓝色 = 全部指标在可接受范围内，红色 = 至少
  一项超差。

## Python API

```python
from mesh_quality import (load_mesh_blocks, fluent_mesh_check,
                          fluent_orthogonal, fluent_aspect_ratio,
                          compute_quality)

# Fluent 网格检查：读取文件，打印报告并返回 dict
report = fluent_mesh_check('test.cas.h5')
report['minimum_face_area']             # 0.006663859...
report['minimum_orthogonal_quality']    # 0.9999999994...
report['maximum_aspect_ratio']          # 6.670067...
report['blocks']['block_0']             # 分块明细

# 逐单元 Fluent 指标（支持 tri/quad/tet/pyramid/wedge/hexa/多面体）
_, blocks = load_mesh_blocks('test.cas.h5')
mesh = blocks[0][1]
oq = fluent_orthogonal(mesh)            # (n_cells,) 0~1，不支持的类型为 NaN
ar = fluent_aspect_ratio(mesh)          # (n_cells,)

# 批量计算（VTK + 自定义），结果可挂回网格用于可视化
arrays, diag = compute_quality(mesh, ['fluent_orthogonal',
                                      'fluent_aspect_ratio', 'skew'])
mesh.cell_data['AR'] = arrays['fluent_aspect_ratio']
mesh.plot(scalars='AR', show_edges=True)
```

## 与 ANSYS Fluent 的一致性

在 test.cas.h5（1600 个 quad 的 2D 网格）上与 Fluent 2026 R1 的
`/mesh/check`（Report Quality）实测对比：

| 指标 | 本工具 | Fluent |
|---|---|---|
| Minimum face area | 6.663859e-03 | 6.663859e-03 |
| Minimum Orthogonal Quality | 1.000000e+00 | 1.00000e+00 |
| Maximum Aspect Ratio | 6.670067e+00 | 6.67000e+00 |

最大长宽比所在单元位置也与 Fluent 一致（质心 (0.52197, 0.99667)）。

实现细节与注意事项：

- 读取 `.cas.h5` 使用 VTK 的 `FLUENTCFFReader`，多 cell zone 文件按块分别
  统计；
- 2D 网格的面为边，"minimum face area" 按惯例输出最小**边长**（与 Fluent
  的 Face area statistics 一致）；
- 自定义指标支持多面体（从 `mesh.cells` 嵌套格式解析面表），不支持的单元
  类型（线、点等）输出 NaN 并跳过；
- 连接性解析为一次 O(n_cells) 的 Python 遍历，其余计算全部向量化；
- CFF 网格节点为 float32，完美网格的指标会带有 ~1e-10 量级的数值噪声，
  报告的容差与排名逻辑已按此处理。
