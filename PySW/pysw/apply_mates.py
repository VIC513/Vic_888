#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import sys
import time
import traceback
from dataclasses import dataclass
from pathlib import Path

import pythoncom
import win32com.client


TARGET_DOC_KEYWORD = "MyMechanism"
TARGET_ASM_TITLE = "MyMechanism"


@dataclass(frozen=True)
class MateSpec:
    name: str
    feat1: str
    comp1_key: str
    feat2: str
    comp2_key: str
    mate_type: int
    mate_align: int


def _connect_solidworks():
    pythoncom.CoInitialize()
    try:
        sw_app = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        sw_app = win32com.client.Dispatch("SldWorks.Application")
    try:
        sw_app.Visible = True
    except Exception:
        pass
    return sw_app


def _get_active_model(sw_app):
    try:
        model = sw_app.ActiveDoc
    except Exception:
        model = None
    if model is None:
        return None
    return model


def _find_model_by_documents(sw_app, keyword: str):
    try:
        docs = sw_app.GetDocuments()
    except Exception:
        docs = None
    if not docs:
        return None

    key = (keyword or "").lower()
    for d in docs:
        try:
            title = str(d.GetTitle() or "")
        except Exception:
            title = ""
        if not title:
            continue
        if key and key in title.lower():
            return d
    return None


def _get_model(sw_app, keyword: str):
    model = _get_active_model(sw_app)
    if model is not None:
        return model
    model = _find_model_by_documents(sw_app, keyword)
    if model is not None:
        return model
    raise RuntimeError("找不到可用文档：请确认 SolidWorks 已打开并加载了目标装配体")


def _activate_doc(sw_app, model) -> None:
    try:
        title = str(model.GetTitle() or "")
    except Exception:
        title = ""
    if not title:
        return
    try:
        sw_app.ActivateDoc2(title, False, 0)
    except Exception:
        pass


def _asm_title_stem(model) -> str:
    try:
        title = str(model.GetTitle() or "")
    except Exception:
        title = ""
    return title.split(".")[0] if title else ""


def _component_name_variants(name2: str) -> list[str]:
    n = (name2 or "").strip()
    if not n:
        return []

    out: list[str] = []
    out.append(n)

    if "@" in n:
        out.append(n.split("@", 1)[0])

    if "-" in n:
        base, tail = n.rsplit("-", 1)
        if tail.isdigit():
            out.append(f"{base}<{tail}>")
    if "<" in n and ">" in n:
        try:
            base = n.split("<", 1)[0]
            tail = n.split("<", 1)[1].split(">", 1)[0]
            if tail.isdigit():
                out.append(f"{base}-{tail}")
        except Exception:
            pass

    uniq: list[str] = []
    for s in out:
        if s and s not in uniq:
            uniq.append(s)
    return uniq


def _path_stem(path: str) -> str:
    try:
        return Path(path).stem
    except Exception:
        return ""


def _iter_component_tree(comp):
    yield comp
    try:
        children = comp.GetChildren()
    except Exception:
        children = None
    if not children:
        return
    for ch in children:
        yield from _iter_component_tree(ch)


def _iter_components(asm):
    comps = None
    try:
        comps = asm.GetComponents(False)
    except Exception:
        comps = None
    if not comps:
        try:
            comps = asm.GetComponents(True)
        except Exception:
            comps = None
    if not comps:
        return
    for c in comps:
        yield from _iter_component_tree(c)


def _get_comp_name2(comp) -> str:
    try:
        return str(comp.Name2 or "")
    except Exception:
        return ""


def _get_comp_path_stem(comp) -> str:
    try:
        path = str(comp.GetPathName() or "")
    except Exception:
        path = ""
    return _path_stem(path) if path else ""


def _print_all_component_name2(asm) -> list[object]:
    comps = []
    try:
        raw = asm.GetComponents(False)
        if raw:
            comps = list(raw)
    except Exception:
        comps = []

    if not comps:
        try:
            raw = asm.GetComponents(True)
            if raw:
                comps = list(raw)
        except Exception:
            comps = []

    print("[组件] 当前装配体组件 Name2 列表：")
    for c in comps:
        try:
            print(f"  - {str(c.Name2 or '')}")
        except Exception:
            print("  - (Name2 读取失败)")
    return comps


def get_comp_by_fuzzy_name(comps: list[object], target_short_name: str):
    key = (target_short_name or "").lower().strip()
    if not key:
        return None
    for comp in comps:
        try:
            name2 = str(comp.Name2 or "")
        except Exception:
            name2 = ""
        if key in name2.lower():
            return comp
    return None


def _resolve_components(asm) -> dict[str, object]:
    comps = _print_all_component_name2(asm)
    return {
        "Base": get_comp_by_fuzzy_name(comps, "Base"),
        "Crank": get_comp_by_fuzzy_name(comps, "Crank"),
        "ConnectingRod": get_comp_by_fuzzy_name(comps, "ConnectingRod"),
        "Slider": get_comp_by_fuzzy_name(comps, "Slider"),
    }


def _select(model, sel: str, sel_type: str, append: bool) -> bool:
    try:
        return bool(model.Extension.SelectByID2(str(sel), str(sel_type), 0.0, 0.0, 0.0, bool(append), 0, None, 0))
    except Exception:
        return False


def _select_face_try(model, sel_candidates: list[str], append: bool) -> tuple[bool, str]:
    for s in sel_candidates:
        ok = _select(model, s, "FACE", append)
        if ok:
            return True, s
    return False, sel_candidates[-1] if sel_candidates else ""


def _clear_sel(model) -> None:
    try:
        model.ClearSelection2(True)
    except Exception:
        pass


def _transform_point(arr: list[float], p: tuple[float, float, float]) -> tuple[float, float, float]:
    x, y, z = p
    return (
        arr[0] * x + arr[1] * y + arr[2] * z + arr[9],
        arr[3] * x + arr[4] * y + arr[5] * z + arr[10],
        arr[6] * x + arr[7] * y + arr[8] * z + arr[11],
    )


def _transform_vector(arr: list[float], v: tuple[float, float, float]) -> tuple[float, float, float]:
    x, y, z = v
    return (
        arr[0] * x + arr[1] * y + arr[2] * z,
        arr[3] * x + arr[4] * y + arr[5] * z,
        arr[6] * x + arr[7] * y + arr[8] * z,
    )


def _collect_named_faces_from_component(comp, keywords: list[str]) -> list[str]:
    key_low = [k.lower() for k in keywords if k]
    try:
        part = comp.GetModelDoc2()
    except Exception:
        part = None
    if part is None:
        return []

    names: list[str] = []
    try:
        bodies = part.GetBodies2(0, False)
    except Exception:
        bodies = None
    if not bodies:
        return []

    for body in bodies:
        try:
            faces = body.GetFaces()
        except Exception:
            faces = None
        if not faces:
            continue
        for face in faces:
            try:
                nm = part.Extension.GetEntityName(face)
            except Exception:
                nm = ""
            nm = str(nm or "").strip()
            if not nm:
                continue
            if not key_low:
                names.append(nm)
            else:
                nlow = nm.lower()
                if any(k in nlow for k in key_low):
                    names.append(nm)

    uniq: list[str] = []
    for n in names:
        if n and n not in uniq:
            uniq.append(n)
    return uniq


def _select_largest_face_by_ray(model, comp, append: bool) -> bool:
    try:
        part = comp.GetModelDoc2()
    except Exception:
        part = None
    if part is None:
        return False

    best_face = None
    best_area = None
    try:
        bodies = part.GetBodies2(0, False)
    except Exception:
        bodies = None
    if not bodies:
        return False

    for body in bodies:
        try:
            faces = body.GetFaces()
        except Exception:
            faces = None
        if not faces:
            continue
        for face in faces:
            try:
                area = float(face.GetArea())
            except Exception:
                continue
            if best_area is None or area > best_area:
                best_area = area
                best_face = face

    if best_face is None:
        return False

    try:
        box = list(best_face.GetBox())
        cx = (float(box[0]) + float(box[3])) / 2.0
        cy = (float(box[1]) + float(box[4])) / 2.0
        cz = (float(box[2]) + float(box[5])) / 2.0
    except Exception:
        return False

    try:
        n = best_face.Normal
        nx, ny, nz = float(n[0]), float(n[1]), float(n[2])
    except Exception:
        nx, ny, nz = 0.0, 0.0, 1.0

    try:
        tf = comp.Transform2
        arr = list(tf.ArrayData) if tf is not None else [1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0]
        if len(arr) < 12:
            return False
    except Exception:
        return False

    ox, oy, oz = _transform_point(arr, (cx, cy, cz))
    dx, dy, dz = _transform_vector(arr, (nx, ny, nz))

    try:
        return bool(model.Extension.SelectByRay(float(ox), float(oy), float(oz), float(dx), float(dy), float(dz), 0.001, 2, bool(append), 0, 0))
    except Exception:
        return False


def _iter_part_features(part) -> list[object]:
    try:
        fm = part.FeatureManager
    except Exception:
        fm = None
    if fm is not None:
        try:
            feats = fm.GetFeatures(True)
            if feats:
                return list(feats)
        except Exception:
            pass

    try:
        feats = part.GetFeatures()
        if feats:
            return list(feats)
    except Exception:
        pass

    out: list[object] = []
    try:
        feat = part.FirstFeature()
    except Exception:
        feat = None
    while feat is not None:
        out.append(feat)
        try:
            feat = feat.GetNextFeature()
        except Exception:
            break
    return out


def _print_part_feature_names(comp, limit: int = 10) -> None:
    name2 = _get_comp_name2(comp)
    stem = _get_comp_path_stem(comp)
    try:
        part = comp.GetModelDoc2()
    except Exception:
        part = None
    if part is None:
        print(f"[特征] {name2} ({stem}): 无法获取零件文档")
        return

    feats = _iter_part_features(part)
    names: list[str] = []
    for f in feats[:limit]:
        try:
            nm = str(f.Name or "")
        except Exception:
            nm = ""
        if nm:
            names.append(nm)
    print(f"[特征] {name2} ({stem}) 前{limit}个: {', '.join(names) if names else '(空)'}")


def _find_feature_by_keywords(part, keywords: list[str]):
    keys = [k.lower() for k in keywords if k]
    feats = _iter_part_features(part)
    for f in feats:
        try:
            nm = str(f.Name or "")
        except Exception:
            nm = ""
        nlow = nm.lower()
        if any(k in nlow for k in keys):
            return f
    return None


def _select_entity(ent, append: bool) -> bool:
    try:
        return bool(ent.Select4(bool(append), None))
    except Exception:
        pass
    try:
        return bool(ent.Select2(bool(append), 0))
    except Exception:
        return False


def _get_corresponding_entity(comp, ent):
    try:
        return comp.GetCorrespondingEntity(ent)
    except Exception:
        return None


def _select_first_face_of_feature(model, comp, feat, append: bool) -> tuple[bool, str]:
    try:
        faces = feat.GetFaces()
    except Exception:
        faces = None
    if not faces:
        return False, "feature.GetFaces() 为空"

    face0 = faces[0]
    ent = _get_corresponding_entity(comp, face0) or face0
    ok = _select_entity(ent, append)
    return ok, "GetFaces()[0]"


def _select_largest_face_entity(model, comp, append: bool) -> tuple[bool, str]:
    try:
        part = comp.GetModelDoc2()
    except Exception:
        part = None
    if part is None:
        return False, "无零件文档"

    best_face = None
    best_area = None
    try:
        bodies = part.GetBodies2(0, False)
    except Exception:
        bodies = None
    if bodies:
        for body in bodies:
            try:
                faces = body.GetFaces()
            except Exception:
                faces = None
            if not faces:
                continue
            for face in faces:
                try:
                    area = float(face.GetArea())
                except Exception:
                    continue
                if best_area is None or area > best_area:
                    best_area = area
                    best_face = face

    if best_face is None:
        return False, "找不到面"

    ent = _get_corresponding_entity(comp, best_face) or best_face
    ok = _select_entity(ent, append)
    return ok, "largest-face"


def _add_mate_closest(asm, mate_type: int) -> tuple[object | None, int | None]:
    err = None
    try:
        err = pythoncom.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    except Exception:
        err = None

    try:
        mate = asm.AddMate5(
            int(mate_type),
            0,
            False,
            0.0,
            0.0,
            0.0,
            1.0,
            1.0,
            0.0,
            0.0,
            0.0,
            False,
            False,
            0,
            err if err is not None else 0,
        )
    except Exception:
        mate = None

    if err is None:
        return mate, None
    try:
        return mate, int(err.value)
    except Exception:
        return mate, None


def _get_open_doc(sw_app, name: str):
    try:
        return sw_app.GetOpenDocumentByName(str(name))
    except Exception:
        return None


def _get_part_doc(sw_app, comp):
    try:
        part = comp.GetModelDoc2()
    except Exception:
        part = None
    if part is not None:
        return part

    candidates: list[str] = []
    try:
        path = str(comp.GetPathName() or "")
    except Exception:
        path = ""
    if path:
        candidates.append(path)
        candidates.append(str(Path(path).name))
        candidates.append(str(Path(path).stem))
        candidates.append(str(Path(path).stem) + ".sldprt")

    name2 = _get_comp_name2(comp)
    if name2:
        candidates.append(name2)
        candidates.append(name2 + ".sldprt")
        candidates.append(name2.split("@", 1)[0])
        candidates.append(name2.split("@", 1)[0] + ".sldprt")

    uniq: list[str] = []
    for c in candidates:
        if c and c not in uniq:
            uniq.append(c)

    for c in uniq:
        d = _get_open_doc(sw_app, c)
        if d is not None:
            return d
    return None


def _component_name_variants_for_select(name2: str, asm_title: str) -> list[str]:
    base_vars = _component_name_variants(name2)
    out: list[str] = []
    for v in base_vars:
        out.append(v)
        if v.split("@", 1)[0] != v:
            out.append(v.split("@", 1)[0])

        b = v
        if "-" in b:
            b = b.split("-", 1)[0]
        if "<" in b and ">" in b:
            b = b.split("<", 1)[0]
        out.append(b)
        if asm_title:
            out.append(f"{b}@{asm_title}")

    uniq: list[str] = []
    for s in out:
        s = (s or "").strip()
        if s and s not in uniq:
            uniq.append(s)
    return uniq


def _select_by_id2_candidates(model, feat_candidates: list[str], comp_name_candidates: list[str], asm_title: str, append: bool) -> tuple[bool, str]:
    for feat in feat_candidates:
        for comp_name in comp_name_candidates:
            if not append:
                _clear_sel(model)
            sel = f"{feat}@{comp_name}@{asm_title}" if asm_title else f"{feat}@{comp_name}"
            print(f"  - try(id2): {sel}")
            if _select(model, sel, "FACE", append):
                return True, sel
    return False, ""


def _select_by_ray_near_component(model, comp, append: bool) -> bool:
    try:
        tf = comp.Transform2
        arr = list(tf.ArrayData) if tf is not None else None
    except Exception:
        arr = None
    if not arr or len(arr) < 12:
        return False

    ox, oy, oz = float(arr[9]), float(arr[10]), float(arr[11])
    dirs = [(0.0, 0.0, 1.0), (0.0, 0.0, -1.0), (1.0, 0.0, 0.0), (-1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, -1.0, 0.0)]
    for dx, dy, dz in dirs:
        try:
            ok = bool(model.Extension.SelectByRay(ox, oy, oz, dx, dy, dz, 0.02, 2, bool(append), 0, 0))
        except Exception:
            ok = False
        if ok:
            return True
    return False


def _get_feature_from_component(sw_app, comp, feature_name: str):
    try:
        path = str(comp.GetPathName() or "")
    except Exception:
        path = ""
    if not path:
        return None

    part = _get_open_doc(sw_app, path)
    if part is None:
        try:
            err = pythoncom.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warn = pythoncom.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        except Exception:
            err = 0
            warn = 0
        try:
            sw_app.OpenDoc6(path, 1, 1, "", err, warn)
        except Exception:
            pass
        part = _get_open_doc(sw_app, path)
        if part is None:
            return None

    try:
        title = str(part.GetTitle() or "")
        if title:
            try:
                sw_app.ActivateDoc3(title, True, 0)
            except Exception:
                try:
                    sw_app.ActivateDoc2(title, False, 0)
                except Exception:
                    pass
    except Exception:
        pass

    try:
        return part.FeatureByName(str(feature_name))
    except Exception:
        return None


def _find_named_face_in_part(part, keyword: str):
    key = (keyword or "").strip().lower()
    if not key:
        return None
    try:
        bodies = part.GetBodies2(0, False)
    except Exception:
        bodies = None
    if not bodies:
        return None
    for body in bodies:
        try:
            faces = body.GetFaces()
        except Exception:
            faces = None
        if not faces:
            continue
        for face in faces:
            try:
                nm = part.Extension.GetEntityName(face)
            except Exception:
                nm = ""
            nm = str(nm or "").strip()
            if nm and key in nm.lower():
                return face
    return None


def _get_component_by_name(sw_assy, comp_name: str):
    try:
        return sw_assy.GetComponentByName(str(comp_name))
    except Exception:
        return None


def _iter_part_faces(part):
    try:
        bodies = part.GetBodies2(0, False)
    except Exception:
        bodies = None
    if not bodies:
        return
    for body in bodies:
        try:
            faces = body.GetFaces()
        except Exception:
            faces = None
        if not faces:
            continue
        for face in faces:
            yield face


def _ensure_part_doc_open_and_active(sw_app, comp):
    try:
        part = comp.GetModelDoc2()
    except Exception:
        part = None
    if part is not None:
        try:
            title = str(part.GetTitle() or "")
            if title:
                try:
                    sw_app.ActivateDoc3(title, True, 0)
                except Exception:
                    try:
                        sw_app.ActivateDoc2(title, False, 0)
                    except Exception:
                        pass
        except Exception:
            pass
        return part

    try:
        path = str(comp.GetPathName() or "")
    except Exception:
        path = ""
    if not path:
        return None

    part = _get_open_doc(sw_app, path)
    if part is None:
        try:
            err = pythoncom.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warn = pythoncom.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        except Exception:
            err = 0
            warn = 0
        try:
            sw_app.OpenDoc6(path, 1, 1, "", err, warn)
        except Exception:
            pass
        part = _get_open_doc(sw_app, path)

    if part is None:
        return None

    try:
        title = str(part.GetTitle() or "")
        if title:
            try:
                sw_app.ActivateDoc3(title, True, 0)
            except Exception:
                try:
                    sw_app.ActivateDoc2(title, False, 0)
                except Exception:
                    pass
    except Exception:
        pass
    return part


def _select_named_face_in_component(sw_app, sw_model, asm_title: str, comp_name: str, keyword: str, append: bool) -> tuple[bool, str]:
    comp = _get_component_by_name(sw_model, comp_name)
    if comp is None:
        return False, f"GetComponentByName('{comp_name}')=None"

    key = (keyword or "").strip().lower()
    if not key:
        return False, "keyword empty"

    try:
        sel = f"{keyword}@{comp_name}@{asm_title}"
        ok = bool(sw_model.Extension.SelectByID2(sel, "FACE", 0.0, 0.0, 0.0, bool(append), 0, None, 0))
    except Exception:
        ok = False
    if ok:
        return True, f"id2-face:{sel}"

    part = _ensure_part_doc_open_and_active(sw_app, comp)
    if part is None:
        return False, "part-doc-open-failed"

    for face in _iter_part_faces(part):
        try:
            nm = part.Extension.GetEntityName(face)
        except Exception:
            nm = ""
        nm = str(nm or "").strip()
        if not nm:
            continue
        if key not in nm.lower():
            continue
        try:
            ent = comp.GetCorrespondingEntity(face)
        except Exception:
            ent = None
        ent = ent or face
        try:
            ok = bool(ent.Select4(True, None))
        except Exception:
            ok = _select_entity(ent, append)
        if ok:
            return True, f"{comp_name}:{nm}"

    return False, "no named face match"


def _apply_one_mate(sw_app, model, asm, asm_title: str, comp_map: dict[str, object], spec: MateSpec) -> None:
    try:
        model.ClearSelection2(True)
    except Exception:
        pass

    try:
        model.EditRebuild3()
    except Exception:
        pass

    comp_name_map = {
        "Base": "Base-1",
        "Crank": "Crank-1",
        "ConnectingRod": "ConnectingRod-1",
        "Slider": "Slider-1",
    }

    comp1_name = comp_name_map.get(spec.comp1_key, "")
    comp2_name = comp_name_map.get(spec.comp2_key, "")
    if not comp1_name or not comp2_name:
        print(f"[失败] {spec.name}: 组件名映射缺失")
        return

    print(f"[选择] {spec.name}:")
    _clear_sel(model)
    ok1, used1 = _select_named_face_in_component(sw_app, model, asm_title, comp1_name, spec.feat1, False)
    print(f"  - try(face-scan): {spec.feat1}@{comp1_name}@{asm_title} -> {used1} ({ok1})")
    ok2, used2 = _select_named_face_in_component(sw_app, model, asm_title, comp2_name, spec.feat2, True) if ok1 else (False, "")
    if ok1:
        print(f"  - try(face-scan): {spec.feat2}@{comp2_name}@{asm_title} -> {used2} ({ok2})")

    if not ok1 or not ok2:
        which = []
        if not ok1:
            which.append("第一对象选择失败")
        if not ok2:
            which.append("第二对象选择失败")
        print(f"[失败] {spec.name}: 找不到特征/选择失败 -> {', '.join(which)}（已用：{used1}{', ' + used2 if used2 else ''}）")
        return

    mate_type = int(spec.mate_type)
    if "同轴" in spec.name:
        mate_type = 0

    mate, err_code = _add_mate_closest(asm, mate_type)
    if mate is None:
        if err_code is None:
            print(f"[失败] {spec.name}: AddMate5 返回 None（未知错误码）")
        else:
            print(f"[失败] {spec.name}: AddMate5 返回 None（ErrorStatus={err_code}）")
        return

    print(f"[成功] {spec.name}")
    try:
        model.EditRebuild3()
    except Exception:
        pass


def main() -> int:
    try:
        time.sleep(3.0)
        sw_app = _connect_solidworks()
        model = _get_model(sw_app, TARGET_DOC_KEYWORD)
        _activate_doc(sw_app, model)
        asm = model

        try:
            model.ShowNamedView2("", -1)
        except Exception:
            pass

        try:
            model.ForceRebuild3(False)
        except Exception:
            pass
        asm_title = TARGET_ASM_TITLE.strip() if TARGET_ASM_TITLE else _asm_title_stem(model)
        if not asm_title:
            print("无法确定装配体标题后缀")
            return 1

        comp_map = _resolve_components(asm)

        if not comp_map:
            print("未发现组件（当前活动文档可能不是装配体，或装配体为空）")
            return 1

        try:
            from win32com.client import constants
            sw_mate_concentric = int(constants.swMateCONCENTRIC)
            sw_mate_coincident = int(constants.swMateCOINCIDENT)
            sw_align_aligned = int(constants.swMateAlignALIGNED)
        except Exception:
            sw_mate_concentric = 0
            sw_mate_coincident = 1
            sw_align_aligned = 1

        mate_specs = [
            MateSpec("配合1 Fixed_Hole ↔ Crank_In", "Fixed_Hole", "Base", "Crank_In", "Crank", sw_mate_concentric, sw_align_aligned),
            MateSpec("配合2 Crank_Out ↔ Link_In", "Crank_Out", "Crank", "Link_In", "ConnectingRod", sw_mate_concentric, sw_align_aligned),
            MateSpec("配合3 Link_Out ↔ Slider_Joint", "Link_Out", "ConnectingRod", "Slider_Joint", "Slider", sw_mate_concentric, sw_align_aligned),
            MateSpec("配合4 Slider_Bottom ↔ Slide_Way", "Slider_Bottom", "Slider", "Slide_Way", "Base", sw_mate_coincident, sw_align_aligned),
        ]

        for spec in mate_specs:
            try:
                _apply_one_mate(sw_app, model, asm, asm_title, comp_map, spec)
            except Exception as e:
                print(f"[异常] {spec.name}: {e}")

        try:
            model.ForceRebuild3(False)
        except Exception:
            pass
        return 0
    except Exception as e:
        print(str(e))
        sys.stderr.write(traceback.format_exc() + "\n")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
