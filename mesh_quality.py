"""Mesh quality report for Ansys Fluent CFF files based on pyvista.

Reads a Fluent mesh file (``.cas.h5``/``.msh.h5`` via the VTK
``FLUENTCFFReader``, ``.cas`` via ``FLUENTReader``), computes cell quality
measures with :meth:`pyvista.DataSetFilters.cell_quality` and prints a
summary report. Fluent's orthogonal quality and aspect ratio are not
available in VTK, so both are computed directly from the mesh geometry
(:func:`fluent_orthogonal`, :func:`fluent_aspect_ratio`) following the
Fluent definitions. Optional outputs: per-cell CSV, JSON summary, a
histogram image and a 3D view. The acceptable ranges used for the
evaluation come from :mod:`pyvista.cell_quality` (Verdict-based).

Usage::

    python mesh_quality.py                          # test.cas.h5, default measures
    python mesh_quality.py path/to/case.cas.h5
    python mesh_quality.py --check                  # Fluent-style mesh check
    python mesh_quality.py --all                    # every measure valid for the mesh
    python mesh_quality.py --measures skew aspect_ratio --csv quality.csv
    python mesh_quality.py --hist quality.png --json quality.json
    python mesh_quality.py --show                   # 3D view of cells outside range
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
import pyvista as pv

try:
    from pyvista.core.utilities.cell_quality import cell_quality_info
except ImportError:  # pragma: no cover - fallback for other pyvista layouts
    from pyvista import cell_quality as _cq
    cell_quality_info = _cq.cell_quality_info

#: Every quality measure name accepted by ``DataSet.cell_quality``.
ALL_MEASURES = [
    'area',
    'aspect_frobenius',
    'aspect_gamma',
    'aspect_ratio',
    'collapse_ratio',
    'condition',
    'diagonal',
    'dimension',
    'distortion',
    'jacobian',
    'max_angle',
    'max_aspect_frobenius',
    'max_edge_ratio',
    'med_aspect_frobenius',
    'min_angle',
    'oddy',
    'radius_ratio',
    'relative_size_squared',
    'scaled_jacobian',
    'shape',
    'shape_and_size',
    'shear',
    'shear_and_size',
    'skew',
    'stretch',
    'taper',
    'volume',
    'warpage',
]

#: Fluent-like default selection: orthogonal quality, aspect ratio, skewness
#: and other shape/size metrics. ``fluent_orthogonal`` and
#: ``fluent_aspect_ratio`` are computed by this module (Fluent definitions)
#: because VTK/Verdict does not provide them; Verdict's own ``aspect_ratio``
#: uses a different, milder formula.
DEFAULT_MEASURES = ['fluent_orthogonal', 'fluent_aspect_ratio', 'skew',
                    'scaled_jacobian', 'aspect_ratio', 'shape', 'area']

#: Custom measures implemented in this module instead of by ``cell_quality``.
CUSTOM_MEASURES = ('fluent_orthogonal', 'fluent_aspect_ratio')

#: Fluent's rule of thumb: the minimum orthogonal quality should stay above
#: roughly 0.15 (0 = worst, 1 = perfect).
ORTHOGONAL_ACCEPTABLE = (0.15, 1.0)

#: Practical CFD guideline for the Fluent aspect ratio (1 is the ideal
#: direction; boundary-layer cells up to ~10 are usually acceptable).
ASPECT_RATIO_ACCEPTABLE = (1.0, 10.0)

#: Aspect ratio of the canonical unit cell of each type (the definitional
#: floor: sqrt(2) for a square, sqrt(3) for a cube, 2 for an equilateral
#: triangle, 3 for a regular tetrahedron, ...).
ASPECT_RATIO_UNIT_VALUES = {
    pv.CellType.TRIANGLE: 2.0,
    pv.CellType.QUAD: np.sqrt(2.0),
    pv.CellType.TETRA: 3.0,
    pv.CellType.PYRAMID: 7.1411,  # unit pyramid: base 1x1, height 0.5
    pv.CellType.WEDGE: np.sqrt(7.0),  # unit prism: edge 1, height 1
    pv.CellType.HEXAHEDRON: np.sqrt(3.0),
}

#: Faces of each supported cell type, as node-index tuples into the cell's
#: own point list (VTK node ordering; orientation is fixed geometrically).
_FACES_2D = {
    pv.CellType.TRIANGLE: ((0, 1), (1, 2), (2, 0)),
    pv.CellType.QUAD: ((0, 1), (1, 2), (2, 3), (3, 0)),
}
_FACES_3D = {
    pv.CellType.TETRA: ((0, 1, 2), (0, 1, 3), (0, 2, 3), (1, 2, 3)),
    pv.CellType.PYRAMID: ((0, 1, 2, 3), (0, 1, 4), (1, 2, 4), (2, 3, 4),
                          (3, 0, 4)),
    pv.CellType.WEDGE: ((0, 1, 2), (3, 4, 5), (0, 1, 4, 3), (1, 2, 5, 4),
                        (2, 0, 3, 5)),
    pv.CellType.HEXAHEDRON: ((0, 1, 2, 3), (4, 5, 6, 7), (0, 1, 5, 4),
                             (1, 2, 6, 5), (2, 3, 7, 6), (3, 0, 4, 7)),
}


class _CustomInfo:
    """Duck-typed stand-in for ``CellQualityInfo`` for custom measures."""

    def __init__(self, measure: str, cell_type: pv.CellType,
                 acceptable_range: tuple[float, float],
                 unit_cell_value: float) -> None:
        self.quality_measure = measure
        self.cell_type = cell_type
        self.acceptable_range = acceptable_range
        self.unit_cell_value = unit_cell_value
        if measure == 'fluent_orthogonal':
            self.normal_range = (0.0, 1.0)
            self.full_range = (0.0, 1.0)
        else:  # fluent_aspect_ratio
            self.normal_range = (1.0, 100.0)
            self.full_range = (0.0, np.inf)


def quality_infos(mesh: pv.DataSet, measure: str,
                  valid: dict[str, list] | None = None) -> list:
    """``(CellType, info)`` pairs for a measure, including custom measures.

    Mirrors the entries built by :func:`measures_valid_for` but also covers
    the custom measures implemented in this module.

    Parameters
    ----------
    mesh : pv.DataSet
        Mesh whose cell types are inspected.
    measure : str
        Quality measure name.
    valid : dict[str, list], optional
        Precomputed result of :func:`measures_valid_for`.

    Returns
    -------
    list
        ``(CellType, info)`` pairs; empty if the measure is unavailable.
    """
    if measure == 'fluent_orthogonal':
        return [(pv.CellType(int(t)),
                 _CustomInfo(measure, pv.CellType(int(t)),
                             ORTHOGONAL_ACCEPTABLE, 1.0))
                for t in np.unique(mesh.celltypes)
                if int(t) in {int(ct) for ct in (*_FACES_2D, *_FACES_3D)}]
    if measure == 'fluent_aspect_ratio':
        return [(pv.CellType(int(t)),
                 _CustomInfo(measure, pv.CellType(int(t)),
                             ASPECT_RATIO_ACCEPTABLE,
                             ASPECT_RATIO_UNIT_VALUES.get(pv.CellType(int(t)),
                                                          1.0)))
                for t in np.unique(mesh.celltypes)
                if int(t) in {int(ct) for ct in (*_FACES_2D, *_FACES_3D)}]
    if valid is None:
        valid = measures_valid_for(mesh)
    return valid.get(measure, [])


def load_mesh_blocks(file_path: str) -> tuple[str, list[tuple[str, pv.DataSet]]]:
    """Read a Fluent mesh file and return (reader_name, [(block_name, mesh)]).

    Parameters
    ----------
    file_path : str
        Path to a ``.cas.h5``, ``.msh.h5``, ``.cas`` or ``.msh`` mesh file.

    Returns
    -------
    tuple[str, list[tuple[str, pv.DataSet]]]
        Name of the VTK reader class and the list of named mesh blocks
        (one block per cell zone in the file).
    """
    reader = pv.get_reader(file_path)
    data = reader.read()
    if isinstance(data, pv.MultiBlock):
        blocks = []
        for i, (name, block) in enumerate(zip(data.keys(), data)):
            if not isinstance(block, pv.UnstructuredGrid):
                block = block.cast_to_unstructured_grid()
            blocks.append((name or f'block_{i}', block))
    else:
        mesh = data if isinstance(data, pv.UnstructuredGrid) else data.cast_to_unstructured_grid()
        blocks = [(Path(file_path).stem, mesh)]
    return type(reader).__name__, blocks


def measures_valid_for(mesh: pv.DataSet) -> dict[str, list]:
    """Map each quality measure to the cell types it is defined for.

    Parameters
    ----------
    mesh : pv.DataSet
        Mesh whose cell types are inspected.

    Returns
    -------
    dict[str, list]
        Measure name -> list of ``(CellType, CellQualityInfo)`` for every
        cell type present in the mesh for which the measure is defined.
    """
    valid: dict[str, list] = {}
    for measure in ALL_MEASURES:
        infos = []
        for raw_type in np.unique(mesh.celltypes):
            cell_type = pv.CellType(int(raw_type))
            try:
                info = cell_quality_info(cell_type, measure)
            except Exception:
                continue  # measure not defined for this cell type
            infos.append((cell_type, info))
        if infos:
            valid[measure] = infos
    for measure in CUSTOM_MEASURES:
        custom_infos = quality_infos(mesh, measure)
        if custom_infos:
            valid[measure] = custom_infos
    return valid


def _connectivity_offsets(mesh: pv.UnstructuredGrid):
    """Locate every cell's connectivity inside ``mesh.cells``.

    Parameters
    ----------
    mesh : pv.UnstructuredGrid
        Mesh to inspect.

    Returns
    -------
    tuple
        ``(id_starts, n_nodes, poly_faces)``: ``id_starts[i]`` is the offset
        of cell *i*'s point ids inside the flat ``mesh.cells`` array,
        ``n_nodes[i]`` the number of cell nodes, and ``poly_faces`` maps the
        indices of polyhedron cells to their list of faces (arrays of global
        point ids). Regular cells have ``n_nodes > 0``; polyhedra have 0.
    """
    faces = np.asarray(mesh.cells)
    n_cells = mesh.n_cells
    id_starts = np.empty(n_cells, dtype=np.int64)
    n_nodes = np.empty(n_cells, dtype=np.int64)
    poly_faces: dict[int, list[np.ndarray]] = {}
    pos = 0
    for i in range(n_cells):
        n = int(faces[pos])
        id_starts[i] = pos + 1
        if n >= 0:  # regular cell: count followed by node ids
            n_nodes[i] = n
            pos += 1 + n
        else:  # VTK_POLYHEDRON: nested [n_faces, (n, ids...) * n_faces]
            n_faces = -n
            pos += 1
            cell_faces = []
            for _ in range(n_faces):
                k = int(faces[pos])
                pos += 1
                cell_faces.append(faces[pos:pos + k])
                pos += k
            poly_faces[i] = cell_faces
            n_nodes[i] = 0
    return id_starts, n_nodes, poly_faces


def _cos_rows(a: np.ndarray, b: np.ndarray) -> np.ndarray:
    """Row-wise cos(angle) between two stacked vectors (0 if undefined)."""
    norm = np.linalg.norm(a, axis=1) * np.linalg.norm(b, axis=1)
    out = np.zeros(len(a))
    ok = norm > 0.0
    out[ok] = np.einsum('ij,ij->i', a, b)[ok] / norm[ok]
    return out


def _polygon_areas_and_centroids(ids: np.ndarray, counts: np.ndarray,
                                 points: np.ndarray):
    """Newell area vectors and centroids for padded polygon rows.

    Parameters
    ----------
    ids : np.ndarray
        ``(n, k_max)`` global point ids per polygon, padded with -1.
    counts : np.ndarray
        ``(n,)`` number of valid vertices per row (>= 3 for Newell).
    points : np.ndarray
        ``(n_points, 3)`` coordinates.

    Returns
    -------
    tuple
        ``(areas, centroids)``: ``(n, 3)`` area vectors (magnitude = area
        for planar polygons) and ``(n, 3)`` vertex means.
    """
    coords = points[np.clip(ids, 0, None)]
    valid = np.arange(ids.shape[1])[None, :] < counts[:, None]
    nxt = (np.arange(ids.shape[1])[None, :] + 1) % np.maximum(counts[:, None], 1)
    p_next = np.take_along_axis(coords, nxt[:, :, None], axis=1)
    cross = np.cross(coords, p_next) * valid[:, :, None]
    areas = 0.5 * cross.sum(axis=1)
    centroids = (coords * valid[:, :, None]).sum(axis=1) / np.maximum(counts, 1)[:, None]
    return areas, centroids


def _face_geometry(mesh: pv.UnstructuredGrid):
    """Shared per-face/per-cell geometry for the custom Fluent measures.

    Builds every face of every supported cell (2D cells contribute their
    edges), computes outward-oriented area vectors and centroids, and the
    per-cell centroids and node tables.

    Parameters
    ----------
    mesh : pv.UnstructuredGrid
        Mesh to inspect.

    Returns
    -------
    tuple
        ``(face_nodes, face_ks, owner, areas, face_centroids,
        cell_centroids, cell_ids, node_counts)``: padded ``(n_faces, k_max)``
        global point ids per face, the number of face nodes, the owning cell
        of each face, outward ``(n_faces, 3)`` area vectors, face and cell
        centroids, and the padded ``(n_cells, k_max)`` global point ids of
        every cell with its node count (``-1``/0 for cells without nodes).
    """
    points = np.asarray(mesh.points)
    celltypes = np.asarray(mesh.celltypes)
    id_starts, n_nodes, poly_faces = _connectivity_offsets(mesh)
    conn = np.asarray(mesh.cells)

    # --- build every face as a row of global point ids -------------------
    face_cells: list[np.ndarray] = []
    face_ids: list[np.ndarray] = []
    for raw_type in np.unique(celltypes):
        t = int(raw_type)
        rows = np.where(celltypes == raw_type)[0]
        if t == int(pv.CellType.POLYHEDRON):
            for i in rows:
                for f in poly_faces.get(int(i), []):
                    face_cells.append(np.array([i]))
                    face_ids.append(np.asarray(f)[None, :])
            continue
        table = _FACES_2D.get(t, _FACES_3D.get(t))
        if table is None:
            continue  # unsupported type: owns no faces
        ftab = np.asarray(table)
        node_ids = np.stack([conn[s:s + n] for s, n in zip(id_starts[rows],
                                                           n_nodes[rows])])
        faces_global = node_ids[:, ftab]  # (C, F, k)
        face_cells.append(np.repeat(rows, len(table)))
        face_ids.append(faces_global.reshape(-1, ftab.shape[1]))

    k_max = max(a.shape[1] for a in face_ids)
    n_faces = sum(len(a) for a in face_ids)
    face_nodes = np.full((n_faces, k_max), -1, dtype=np.int64)
    face_ks = np.empty(n_faces, dtype=np.int64)
    owner = np.empty(n_faces, dtype=np.int64)
    pos = 0
    for cells_part, ids_part in zip(face_cells, face_ids):
        m = len(ids_part)
        face_nodes[pos:pos + m, :ids_part.shape[1]] = ids_part
        face_ks[pos:pos + m] = ids_part.shape[1]
        owner[pos:pos + m] = cells_part
        pos += m

    # --- padded per-cell node tables and centroids ------------------------
    regular = n_nodes > 0
    width = int(n_nodes[regular].max()) if np.any(regular) else 1
    for faces_list in poly_faces.values():  # polyhedra: unique face nodes
        width = max(width, len(np.unique(np.concatenate(faces_list))))
    cell_ids = np.full((mesh.n_cells, width), -1, dtype=np.int64)
    cell_centroids = np.full((mesh.n_cells, 3), np.nan)
    if np.any(regular):
        cols = np.arange(width)[None, :]
        idx = id_starts[regular, None] + np.minimum(cols,
                                                    n_nodes[regular, None] - 1)
        cell_ids[np.where(regular)[0]] = np.where(
            cols < n_nodes[regular, None], conn[idx], -1)
        _, centroids = _polygon_areas_and_centroids(
            cell_ids[regular], n_nodes[regular], points)
        cell_centroids[np.where(regular)[0]] = centroids
    for i, faces_list in poly_faces.items():
        nodes = np.unique(np.concatenate(faces_list))
        cell_ids[i, :len(nodes)] = nodes
        cell_centroids[i] = points[nodes].mean(axis=0)

    # --- face area vectors, outward orientation and centroids -------------
    areas, face_centroids = _polygon_areas_and_centroids(face_nodes, face_ks,
                                                         points)
    is_edge = face_ks == 2
    if np.any(is_edge):  # 2D cells: edge area vector = length * outward
        plane_vecs, _ = _polygon_areas_and_centroids(cell_ids[regular],
                                                     n_nodes[regular], points)
        plane_normals = np.full((mesh.n_cells, 3), np.nan)
        with np.errstate(invalid='ignore', divide='ignore'):
            plane_normals[regular] = (
                plane_vecs.T / np.linalg.norm(plane_vecs, axis=1)).T
        edge_rows = np.where(is_edge)[0]
        p0 = points[face_nodes[edge_rows, 0]]
        p1 = points[face_nodes[edge_rows, 1]]
        areas[edge_rows] = np.cross(p1 - p0, plane_normals[owner[edge_rows]])

    to_face = face_centroids - cell_centroids[owner]
    sign = np.sign(np.einsum('ij,ij->i', areas, to_face))
    sign[sign == 0] = 1.0
    areas *= sign[:, None]  # orient outward from the owning cell

    return face_nodes, face_ks, owner, areas, face_centroids, \
        cell_centroids, cell_ids, n_nodes


def fluent_orthogonal(mesh: pv.UnstructuredGrid) -> np.ndarray:
    """Compute the Fluent-style orthogonal quality of every cell.

    Follows the ANSYS Fluent definition: for each face *f* of a cell,
    with outward area vector ``A_f``, the vector ``c_f`` from the cell
    centroid to the face centroid and the vector ``r_f`` from the face
    centroid to the centroid of the cell sharing the face, the cell value
    is ``min over faces of min(A_f.c_f/|A_f||c_f|, A_f.r_f/|A_f||r_f|)``
    (boundary faces use only the first term). 2D cells use edges instead
    of faces. Values range 0 (worst) to 1 (perfect); NaN for cells whose
    type is not supported (only triangle, quad, tetra, pyramid, wedge,
    hexahedron and polyhedron are).

    Parameters
    ----------
    mesh : pv.UnstructuredGrid
        Mesh to evaluate.

    Returns
    -------
    np.ndarray
        ``(n_cells,)`` orthogonal quality values.
    """
    if mesh.n_cells == 0:
        return np.zeros(0)
    face_nodes, _, owner, areas, face_centroids, cell_centroids, _, _ = \
        _face_geometry(mesh)

    oq = np.full(mesh.n_cells, np.inf)
    to_face = face_centroids - cell_centroids[owner]
    np.minimum.at(oq, owner, _cos_rows(areas, to_face))

    keys = np.sort(face_nodes, axis=1)
    _, inverse = np.unique(keys, axis=0, return_inverse=True)
    n_shared = np.bincount(inverse)
    order = np.argsort(inverse, kind='stable')
    grouped = order[np.isin(inverse[order], np.where(n_shared == 2)[0])]
    if len(grouped) >= 2:  # interior faces contribute the neighbour term
        r1, r2 = grouped[0::2], grouped[1::2]
        c1, c2 = owner[r1], owner[r2]
        np.minimum.at(oq, c1, _cos_rows(areas[r1],
                                        cell_centroids[c2] - face_centroids[r1]))
        np.minimum.at(oq, c2, _cos_rows(-areas[r1],
                                        cell_centroids[c1] - face_centroids[r1]))

    oq = np.clip(oq, 0.0, 1.0)
    oq[np.isinf(oq)] = np.nan  # cells that own no face are unsupported
    return oq


def fluent_aspect_ratio(mesh: pv.UnstructuredGrid) -> np.ndarray:
    """Compute the ANSYS Fluent cell aspect ratio of every cell.

    Fluent defines the aspect ratio as the ratio of the maximum to the
    minimum of these distances: the normal distances between the cell
    centroid and the face centroids (the projection of the centroid-to-face
    vector onto the face normal) and the distances between the cell centroid
    and its nodes. 2D cells use edges instead of faces. Perfect cells have a
    characteristic floor value (sqrt(2) for squares, sqrt(3) for cubes, 2
    for equilateral triangles, 3 for regular tetrahedra); NaN for cells
    whose type is not supported (only triangle, quad, tetra, pyramid, wedge,
    hexahedron and polyhedron are).

    Parameters
    ----------
    mesh : pv.UnstructuredGrid
        Mesh to evaluate.

    Returns
    -------
    np.ndarray
        ``(n_cells,)`` aspect ratio values.
    """
    if mesh.n_cells == 0:
        return np.zeros(0)
    _, _, owner, areas, face_centroids, cell_centroids, cell_ids, _ = \
        _face_geometry(mesh)
    points = np.asarray(mesh.points)

    with np.errstate(divide='ignore', invalid='ignore'):
        # normal distance from the cell centroid to each face centroid
        normal_dist = np.abs(np.einsum(
            'ij,ij->i', areas,
            face_centroids - cell_centroids[owner])) / np.linalg.norm(areas,
                                                                      axis=1)
    normal_dist = np.where(np.isfinite(normal_dist), normal_dist, 0.0)

    valid = cell_ids >= 0
    dist = np.linalg.norm(points[np.clip(cell_ids, 0, None)]
                          - cell_centroids[:, None, :], axis=2)
    finite_cells = np.isfinite(cell_centroids[:, 0])[:, None]
    node_max = np.where(valid, dist, -np.inf).max(axis=1)
    node_min = np.where(valid & finite_cells, dist, np.inf).min(axis=1)

    ar = np.full(mesh.n_cells, np.nan)
    c_max, c_min = node_max.copy(), node_min.copy()
    np.maximum.at(c_max, owner, normal_dist)
    np.minimum.at(c_min, owner, normal_dist)
    has_face = np.zeros(mesh.n_cells, dtype=bool)
    has_face[owner] = True
    ok = has_face & (c_min > 0.0)
    ar[ok] = c_max[ok] / c_min[ok]
    return ar


def fluent_mesh_check(file_path: str, verbose: bool = True) -> dict:
    """Fluent-style mesh check: the headline statistics of a mesh file.

    Reports the three quality numbers Fluent's mesh check prints:

    * ``minimum_face_area`` - the smallest face area in the mesh; for 2D
      meshes the "faces" are edges, so - exactly like Fluent's face-area
      statistics - the smallest edge length is reported,
    * ``minimum_orthogonal_quality`` - Fluent definition, 0 (worst) to
      1 (perfect), see :func:`fluent_orthogonal`,
    * ``maximum_aspect_ratio`` - Fluent definition, see
      :func:`fluent_aspect_ratio`.

    Parameters
    ----------
    file_path : str
        Mesh file (``.cas.h5``/``.msh.h5``/``.cas``/``.msh``).
    verbose : bool
        When True, print a Fluent-style report.

    Returns
    -------
    dict
        Per-block statistics plus the overall ``minimum_face_area``,
        ``minimum_orthogonal_quality`` and ``maximum_aspect_ratio``.
    """
    try:  # silence VTK notices such as "No data file (.dat.h5) found"
        import vtk
        vtk.vtkLogger.SetStderrVerbosity(vtk.vtkLogger.VERBOSITY_ERROR)
    except Exception:
        pass
    reader_name, blocks = load_mesh_blocks(file_path)
    report: dict = {'file': str(file_path), 'reader': reader_name,
                    'minimum_face_area': np.inf,
                    'minimum_orthogonal_quality': np.inf,
                    'maximum_aspect_ratio': -np.inf, 'blocks': {}}
    if verbose:
        line = '=' * 74
        print(line)
        print(f" Fluent mesh check: {file_path}  (reader: {reader_name})")
        print(line)

    for name, mesh in blocks:
        sized = mesh.compute_cell_sizes(length=False, area=True, volume=True)
        is_2d = not np.any(np.asarray(sized.cell_data['Volume'],
                                      dtype=float) > 0.0)

        # Fluent face-area statistics: the smallest face area in 3D, the
        # smallest edge length in 2D (there the "faces" are edges).
        _, _, _, areas, _, _, _, _ = _face_geometry(mesh)
        face_area = np.linalg.norm(areas, axis=1)
        min_face = float(face_area.min()) if face_area.size else np.nan

        oq = fluent_orthogonal(mesh)
        ar = fluent_aspect_ratio(mesh)
        oq_valid = oq[~np.isnan(oq)]
        ar_valid = ar[~np.isnan(ar)]
        stats = {
            'n_cells': int(mesh.n_cells),
            'is_2d': bool(is_2d),
            'minimum_face_area': min_face,
            'minimum_orthogonal_quality': (float(oq_valid.min())
                                           if oq_valid.size else float('nan')),
            'maximum_aspect_ratio': (float(ar_valid.max())
                                     if ar_valid.size else float('nan')),
        }
        report['blocks'][name] = stats
        report['minimum_face_area'] = min(report['minimum_face_area'],
                                          min_face)
        report['minimum_orthogonal_quality'] = min(
            report['minimum_orthogonal_quality'],
            stats['minimum_orthogonal_quality'])
        report['maximum_aspect_ratio'] = max(
            report['maximum_aspect_ratio'],
            stats['maximum_aspect_ratio'])

        if verbose:
            dims = '2D' if is_2d else '3D'
            print(f"\n Block: {name}   Cells: {stats['n_cells']}   {dims}")
            print(f"   Minimum face area            = {min_face:.6e}")
            print(f"   Minimum orthogonal quality   = "
                  f"{stats['minimum_orthogonal_quality']:.6e}")
            print(f"   Maximum aspect ratio         = "
                  f"{stats['maximum_aspect_ratio']:.6e}")

    if verbose and len(blocks) > 1:
        print(f"\n Overall ({len(blocks)} blocks)")
        print(f"   Minimum face area            = "
              f"{report['minimum_face_area']:.6e}")
        print(f"   Minimum orthogonal quality   = "
              f"{report['minimum_orthogonal_quality']:.6e}")
        print(f"   Maximum aspect ratio         = "
              f"{report['maximum_aspect_ratio']:.6e}")
    return report


def compute_quality(mesh: pv.DataSet, measures: list[str]) -> tuple[dict[str, np.ndarray], dict]:
    """Compute quality measures for a mesh, cleaning VTK sentinel values.

    Parameters
    ----------
    mesh : pv.DataSet
        Mesh to evaluate.
    measures : list[str]
        Quality measure names for :meth:`cell_quality`.

    Returns
    -------
    tuple[dict[str, np.ndarray], dict]
        Per-measure cell arrays (unsupported cell types masked as NaN) and
        diagnostics from :meth:`compute_cell_sizes` (cell areas, volumes,
        whether the mesh is effectively 2D) plus the list of skipped
        measures.
    """
    sized = mesh.compute_cell_sizes(length=False, area=True, volume=True)
    cell_area = np.asarray(sized.cell_data['Area'], dtype=float)
    cell_volume = np.asarray(sized.cell_data['Volume'], dtype=float)
    diagnostics = {
        'cell_area': cell_area,
        'cell_volume': cell_volume,
        'is_2d': not np.any(cell_volume > 0.0),
    }

    valid = measures_valid_for(mesh)
    computable = [m for m in measures if m in valid]
    skipped = [m for m in measures if m not in valid]
    diagnostics['skipped'] = list(skipped)

    custom_functions = {
        'fluent_orthogonal': fluent_orthogonal,
        'fluent_aspect_ratio': fluent_aspect_ratio,
    }
    custom = [m for m in computable if m in CUSTOM_MEASURES]
    vtk_measures = [m for m in computable if m not in CUSTOM_MEASURES]
    quality = mesh.cell_quality(vtk_measures) if vtk_measures else None

    arrays: dict[str, np.ndarray] = {}
    for measure in vtk_measures:
        values = np.asarray(quality.cell_data[measure], dtype=float).copy()
        # VTK fills cells whose type does not support a measure with the
        # sentinel -1; mask them so statistics only use defined values.
        supported = [int(ct) for ct, _ in valid[measure]]
        unsupported = ~np.isin(mesh.celltypes, supported)
        values[unsupported & (values == -1.0)] = np.nan
        if np.all(np.isnan(values)):
            diagnostics.setdefault('skipped', []).append(measure)
            continue  # nothing usable (e.g. 'volume' on a quad mesh)
        arrays[measure] = values
    for measure in custom:
        arrays[measure] = custom_functions[measure](mesh)
    return arrays, diagnostics


def _range_tolerance(lo: float, hi: float) -> float:
    """Tolerance for range comparisons so float round-off never flags
    perfect cells (e.g. ``1.0 + 1e-12`` against a bound of ``1.0``)."""
    return 1e-9 * max(1.0, abs(lo) if np.isfinite(lo) else 1.0,
                      abs(hi) if np.isfinite(hi) else 1.0)


def acceptable_mask(values: np.ndarray, celltypes: np.ndarray,
                    infos: list) -> tuple[np.ndarray, np.ndarray]:
    """Split cells into inside/outside the Verdict acceptable range.

    Parameters
    ----------
    values : np.ndarray
        Quality values, NaN for undefined cells.
    celltypes : np.ndarray
        Per-cell VTK cell type ids.
    infos : list
        ``(CellType, CellQualityInfo)`` pairs from :func:`measures_valid_for`.

    Returns
    -------
    tuple[np.ndarray, np.ndarray]
        Booleans ``(evaluated, ok)``; both False for cells with no range.
    """
    evaluated = np.zeros(values.shape, dtype=bool)
    ok = np.zeros(values.shape, dtype=bool)
    for cell_type, info in infos:
        lo, hi = info.acceptable_range
        tol = _range_tolerance(lo, hi)
        sel = celltypes == int(cell_type)
        evaluated[sel & ~np.isnan(values)] = True
        ok[sel] = (values[sel] >= lo - tol) & (values[sel] <= hi + tol)
    return evaluated, ok & evaluated


def badness(values: np.ndarray, lo: float, hi: float) -> np.ndarray:
    """Distance outside the acceptable range, normalized by its width.

    Parameters
    ----------
    values : np.ndarray
        Quality values (NaN allowed).
    lo, hi : float
        Acceptable range bounds.

    Returns
    -------
    np.ndarray
        0 for cells inside the range, increasing the further out a cell is.
    """
    tol = _range_tolerance(lo, hi)
    width = hi - lo
    if width > 0:
        bad = np.maximum(np.maximum(lo - tol - values, values - hi - tol), 0.0) / width
    else:  # degenerate range such as (0, 0)
        bad = np.abs(values - lo)
    bad[np.isnan(values)] = np.nan
    return bad


def worst_ranking(values: np.ndarray, celltypes: np.ndarray,
                  infos: list) -> np.ndarray:
    """Rank cells from the most to the least problematic.

    Cells outside their acceptable range rank first by normalized distance;
    when no cell is outside, cells are ranked by distance from the ideal
    (unit-cell) value instead, so the list always names the worst cells.

    Parameters
    ----------
    values : np.ndarray
        Quality values (NaN allowed).
    celltypes : np.ndarray
        Per-cell VTK cell type ids.
    infos : list
        ``(CellType, CellQualityInfo)`` pairs from :func:`measures_valid_for`.

    Returns
    -------
    np.ndarray
        Ranking score; ``-inf`` for cells the measure is undefined for.
    """
    rank = np.zeros(values.shape, dtype=float)
    any_outside = False
    for cell_type, info in infos:
        lo, hi = info.acceptable_range
        sel = celltypes == int(cell_type)
        rank[sel] = badness(values[sel], lo, hi)
        finite = rank[sel][~np.isnan(rank[sel])]
        any_outside = any_outside or bool(np.any(finite > 0.0))
    if not any_outside:
        for cell_type, info in infos:
            lo, hi = info.acceptable_range
            width = hi - lo
            width = width if 0.0 < width < np.inf else 1.0
            sel = celltypes == int(cell_type)
            rank[sel] = np.abs(values[sel] - info.unit_cell_value) / width
    rank[np.isnan(values)] = -np.inf  # never rank undefined cells
    return rank


def summarize(name: str, mesh: pv.DataSet, arrays: dict[str, np.ndarray],
              diagnostics: dict) -> dict:
    """Build the per-block report data (stats, acceptable %, worst cells).

    Parameters
    ----------
    name : str
        Block (zone) name.
    mesh : pv.DataSet
        Mesh the arrays were computed on.
    arrays : dict[str, np.ndarray]
        Quality arrays from :func:`compute_quality`.
    diagnostics : dict
        Diagnostics part of the :func:`compute_quality` result.

    Returns
    -------
    dict
        JSON-serializable summary for one block.
    """
    celltypes = np.asarray(mesh.celltypes)
    counts = {
        pv.CellType(int(t)).name: int(c)
        for t, c in zip(*np.unique(celltypes, return_counts=True))
    }
    valid = measures_valid_for(mesh)

    measures_summary = {}
    for measure, values in arrays.items():
        finite = values[~np.isnan(values)]
        infos = quality_infos(mesh, measure, valid)
        evaluated, ok = acceptable_mask(values, celltypes, infos)
        measures_summary[measure] = {
            'min': float(np.min(finite)),
            'max': float(np.max(finite)),
            'mean': float(np.mean(finite)),
            'std': float(np.std(finite)),
            'p99': float(np.percentile(finite, 99)),
            'n_evaluated': int(evaluated.sum()),
            'n_outside_acceptable': int((evaluated & ~ok).sum()),
        }

    worst = {}
    for measure, values in arrays.items():
        rank = worst_ranking(values, celltypes,
                             quality_infos(mesh, measure, valid))
        finite = rank[np.isfinite(rank)]
        if finite.size == 0 or np.max(finite) <= 1e-8:
            # every cell sits on the ideal value within float32 round-off
            worst[measure] = []
        else:
            order = [int(i) for i in np.argsort(rank)[::-1]
                     if np.isfinite(rank[i])][:5]
            worst[measure] = [
                {'cell_id': int(i), 'value': float(values[i])} for i in order
            ]

    return {
        'name': name,
        'n_cells': int(mesh.n_cells),
        'n_points': int(mesh.n_points),
        'cell_types': counts,
        'is_2d': diagnostics['is_2d'],
        'total_area': float(np.sum(diagnostics['cell_area'])),
        'total_volume': float(np.sum(diagnostics['cell_volume'])),
        'measures': measures_summary,
        'worst_cells': worst,
    }


def print_report(summary: dict) -> None:
    """Print one block's summary as a formatted console report.

    Parameters
    ----------
    summary : dict
        Per-block summary produced by :func:`summarize`.
    """
    line = '=' * 74
    print(line)
    print(f" Block: {summary['name']}")
    print(line)
    dims = '2D' if summary['is_2d'] else '3D'
    types = ', '.join(f'{t} x{n}' for t, n in summary['cell_types'].items())
    print(f" Cells: {summary['n_cells']}   Points: {summary['n_points']}   "
          f"Type(s): {types}   Dimension: {dims}")
    print(f" Total area: {summary['total_area']:.6g}   "
          f"Total volume: {summary['total_volume']:.6g}")
    print()
    header = (f" {'measure':<20s} {'min':>11s} {'max':>11s} {'mean':>11s} "
              f"{'std':>11s} {'p99':>11s} {'outside':>13s}")
    print(header)
    print(' ' + '-' * (len(header) - 1))
    for measure, stats in summary['measures'].items():
        n_eval = max(stats['n_evaluated'], 1)
        outside = (f"{stats['n_outside_acceptable']} "
                   f"({100 * stats['n_outside_acceptable'] / n_eval:.2f}%)")
        print(f" {measure:<20s} {stats['min']:>11.4g} {stats['max']:>11.4g} "
              f"{stats['mean']:>11.4g} {stats['std']:>11.4g} {stats['p99']:>11.4g} "
              f"{outside:>13s}")
    for measure, cells in summary['worst_cells'].items():
        if not cells:
            print(f" worst [{measure}]: all cells within acceptable range")
            continue
        items = ', '.join(f"#{c['cell_id']}={c['value']:.4g}" for c in cells[:5])
        print(f" worst [{measure}]: {items}")
    print()


def save_histogram(blocks: list[tuple[str, pv.DataSet]], summaries: list[dict],
                   arrays_per_block: list[dict[str, np.ndarray]], out_path: str) -> None:
    """Save a histogram image (one subplot per measure) with acceptable ranges.

    Parameters
    ----------
    blocks : list[tuple[str, pv.DataSet]]
        ``(name, mesh)`` pairs from :func:`load_mesh_blocks`.
    summaries : list[dict]
        Per-block summaries.
    arrays_per_block : list[dict[str, np.ndarray]]
        Per-block quality arrays from :func:`compute_quality`.
    out_path : str
        Image file path (e.g. ``quality.png``).
    """
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    measures = list(arrays_per_block[0].keys())
    ncols = min(3, len(measures))
    nrows = (len(measures) + ncols - 1) // ncols
    fig, axes = plt.subplots(nrows, ncols, figsize=(6 * ncols, 3.8 * nrows),
                             squeeze=False)

    # acceptable range of the first cell type defining each measure
    for idx, measure in enumerate(measures):
        ax = axes.ravel()[idx]
        all_values = np.concatenate([
            a[measure][~np.isnan(a[measure])] for a in arrays_per_block
        ])
        ax.hist(all_values, bins=50, color='#4C72B0', alpha=0.85)
        lo, hi = quality_infos(blocks[0][1], measure)[0][1].acceptable_range
        edges = [bound for bound in (lo, hi) if np.isfinite(bound)]
        label = f'acceptable [{lo:g}, {hi:g}]'
        for edge_idx, edge in enumerate(edges):
            ax.axvline(edge, color='darkorange', ls='--', lw=1.5,
                       label=label if edge_idx == 0 else None)
        if edges:
            ax.legend(fontsize=8)
        ax.set_title(measure)
        ax.set_xlabel('quality value')
        ax.set_ylabel('cells')
        ax.grid(alpha=0.3)
    for idx in range(len(measures), len(axes.ravel())):
        axes.ravel()[idx].axis('off')
    fig.suptitle('Cell quality histograms', fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    print(f"Histogram saved to: {out_path}")


def save_csv(blocks: list[tuple[str, pv.DataSet]], arrays_per_block: list[dict[str, np.ndarray]],
             out_path: str) -> None:
    """Save per-cell quality values to a CSV file.

    Parameters
    ----------
    blocks : list[tuple[str, pv.DataSet]]
        ``(name, mesh)`` pairs from :func:`load_mesh_blocks`.
    arrays_per_block : list[dict[str, np.ndarray]]
        Per-block quality arrays from :func:`compute_quality`.
    out_path : str
        CSV file path.
    """
    measures = list(arrays_per_block[0].keys())
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        f.write('block,cell_id,cell_type,' + ','.join(measures) + '\n')
        for (name, mesh), arrays in zip(blocks, arrays_per_block):
            celltypes = np.asarray(mesh.celltypes)
            for cell_id in range(mesh.n_cells):
                row = [name, str(cell_id), pv.CellType(int(celltypes[cell_id])).name]
                row += ['' if np.isnan(arrays[m][cell_id]) else f'{arrays[m][cell_id]:.6g}'
                        for m in measures]
                f.write(','.join(row) + '\n')
    print(f"CSV saved to: {out_path}")


def show_plot(mesh: pv.DataSet, arrays: dict[str, np.ndarray]) -> None:
    """Open a 3D view with cells outside their acceptable range in red.

    Parameters
    ----------
    mesh : pv.DataSet
        First block mesh to display.
    arrays : dict[str, np.ndarray]
        Quality arrays of that mesh.
    """
    valid = measures_valid_for(mesh)
    celltypes = np.asarray(mesh.celltypes)
    outside = np.zeros(mesh.n_cells, dtype=bool)
    for measure, values in arrays.items():
        for cell_type, info in quality_infos(mesh, measure, valid):
            lo, hi = info.acceptable_range
            sel = celltypes == int(cell_type)
            outside[sel] |= ~np.isnan(values[sel]) & ((values[sel] < lo) | (values[sel] > hi))
    mesh.cell_data['outside_acceptable'] = outside.astype(float)
    plotter = pv.Plotter(title='cffview - mesh quality')
    plotter.add_mesh(mesh, scalars='outside_acceptable', cmap='bwr',
                     clim=[0, 1], show_edges=True,
                     annotations={0.0: 'acceptable', 1.0: 'outside range'})
    plotter.add_axes()
    plotter.show()


def main() -> None:
    """Parse arguments, compute quality and emit the report/outputs."""
    try:  # silence VTK notices such as "No data file (.dat.h5) found"
        import vtk
        vtk.vtkLogger.SetStderrVerbosity(vtk.vtkLogger.VERBOSITY_ERROR)
    except Exception:
        pass
    parser = argparse.ArgumentParser(
        prog='mesh_quality',
        description='Mesh quality report for Fluent .cas.h5 files (pyvista cell_quality).',
    )
    parser.add_argument('file_path', nargs='?', default='test.cas.h5',
                        help='mesh file (.cas.h5/.msh.h5/.cas); default: test.cas.h5')
    parser.add_argument('--check', action='store_true',
                        help='print the Fluent mesh check (minimum face area, '
                             'minimum orthogonal quality, maximum aspect ratio) '
                             'and exit')
    parser.add_argument('--measures', nargs='+',
                        choices=ALL_MEASURES + list(CUSTOM_MEASURES),
                        default=DEFAULT_MEASURES, metavar='MEASURE',
                        help=f"quality measures to compute (default: {' '.join(DEFAULT_MEASURES)})")
    parser.add_argument('--all', action='store_true',
                        help='compute every measure valid for the cell types in the mesh')
    parser.add_argument('--csv', nargs='?', const='mesh_quality.csv', default=None,
                        metavar='FILE', help='save per-cell values to CSV')
    parser.add_argument('--json', nargs='?', const='mesh_quality.json', default=None,
                        metavar='FILE', help='save summary to JSON')
    parser.add_argument('--hist', nargs='?', const='mesh_quality.png', default=None,
                        metavar='FILE', help='save histogram image')
    parser.add_argument('--show', action='store_true',
                        help='open a 3D view with out-of-range cells in red')
    args = parser.parse_args()

    if not Path(args.file_path).is_file():
        parser.error(f"file not found: {args.file_path}")

    if args.check:
        fluent_mesh_check(args.file_path)
        return

    reader_name, blocks = load_mesh_blocks(args.file_path)
    print(f"File: {args.file_path}")
    print(f"Reader: {reader_name}   Blocks: {len(blocks)}")
    print()

    summaries, arrays_per_block = [], []
    for name, mesh in blocks:
        measures = sorted(measures_valid_for(mesh)) if args.all else args.measures
        arrays, diagnostics = compute_quality(mesh, measures)
        summaries.append(summarize(name, mesh, arrays, diagnostics))
        arrays_per_block.append(arrays)
        print_report(summaries[-1])
        for skipped in diagnostics.get('skipped', []):
            print(f" note: '{skipped}' is not defined for this mesh and was skipped.")
        if not arrays:
            print(" note: none of the requested measures could be computed.")
            print()

    if not any(arrays_per_block):
        return

    if args.json:
        payload = {'file': str(args.file_path), 'reader': reader_name,
                   'blocks': summaries}
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print(f"JSON summary saved to: {args.json}")
    if args.csv:
        save_csv(blocks, arrays_per_block, args.csv)
    if args.hist:
        save_histogram(blocks, summaries, arrays_per_block, args.hist)
    if args.show:
        show_plot(blocks[0][1], arrays_per_block[0])


if __name__ == '__main__':
    main()
