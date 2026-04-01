#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import sys
import time
import traceback
from dataclasses import dataclass
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except Exception as e:
    raise SystemExit(
        "缺少依赖：pywin32。请先在你的 sw_env 环境中安装：pip install pywin32\n"
        f"原始错误：{e}"
    )

MM_TO_M = 0.001
USE_CHINESE_NAME = True


class SolidWorksAutomationError(RuntimeError):
    pass


@dataclass(frozen=True)
class MechanismParams:
    R_mm: float = 50.0
    L_mm: float = 150.0
    slider_size_mm: float = 30.0
    bar_width_mm: float = 10.0
    bar_thickness_mm: float = 10.0
    hole_d_mm: float = 10.0
    end_margin_mm: float = 10.0
    slider_thickness_mm: float = 20.0

    base_len_mm: float = 300.0
    base_width_mm: float = 60.0
    base_thickness_mm: float = 10.0
    slot_width_mm: float = 32.0
    slot_len_mm: float = 220.0
    pivot_x_mm: float = 50.0

    pin_d_mm: float = 8.0
    pin_len_mm: float = 25.0

    @property
    def R(self) -> float:
        return self.R_mm * MM_TO_M

    @property
    def L(self) -> float:
        return self.L_mm * MM_TO_M

    @property
    def slider_size(self) -> float:
        return self.slider_size_mm * MM_TO_M

    @property
    def bar_width(self) -> float:
        return self.bar_width_mm * MM_TO_M

    @property
    def bar_thickness(self) -> float:
        return self.bar_thickness_mm * MM_TO_M

    @property
    def hole_d(self) -> float:
        return self.hole_d_mm * MM_TO_M

    @property
    def end_margin(self) -> float:
        return self.end_margin_mm * MM_TO_M

    @property
    def slider_thickness(self) -> float:
        return self.slider_thickness_mm * MM_TO_M

    @property
    def base_len(self) -> float:
        return self.base_len_mm * MM_TO_M

    @property
    def base_width(self) -> float:
        return self.base_width_mm * MM_TO_M

    @property
    def base_thickness(self) -> float:
        return self.base_thickness_mm * MM_TO_M

    @property
    def slot_width(self) -> float:
        return self.slot_width_mm * MM_TO_M

    @property
    def slot_len(self) -> float:
        return self.slot_len_mm * MM_TO_M

    @property
    def pivot_x(self) -> float:
        return self.pivot_x_mm * MM_TO_M

    @property
    def pin_d(self) -> float:
        return self.pin_d_mm * MM_TO_M

    @property
    def pin_len(self) -> float:
        return self.pin_len_mm * MM_TO_M


def _send_sw_msg(sw_app, text: str) -> None:
    try:
        sw_app.SendMsgToUser2(text, 0, 0)
    except Exception:
        pass


FEATURE_NAME_MAP = {
    "Front Plane": "前视基准面",
    "Top Plane": "上视基准面",
    "Right Plane": "右视基准面",
    "Origin": "原点",
}




def _select_plane(model, plane_name_en: str) -> None:
    candidates: list[str] = []
    if USE_CHINESE_NAME:
        candidates.append(FEATURE_NAME_MAP.get(plane_name_en, plane_name_en))
    candidates.append(plane_name_en)

    last_error: str | None = None
    for name in candidates:
        if not name:
            continue
        try:
            feat = model.FeatureByName(str(name))
        except Exception as e:
            last_error = str(e)
            sys.stderr.write(f"警告：FeatureByName('{name}') 调用失败：{e}\n")
            continue

        if feat is None:
            sys.stderr.write(f"警告：FeatureByName 未找到：'{name}'\n")
            continue

        try:
            ok = bool(feat.Select2(False, 0))
        except Exception:
            try:
                ok = bool(feat.Select2(False))
            except Exception as e:
                last_error = str(e)
                sys.stderr.write(f"警告：Select2 失败（'{name}'）：{e}\n")
                ok = False

        if ok:
            return

    if plane_name_en == "Origin" and USE_CHINESE_NAME and "Origin" not in candidates:
        try:
            feat = model.FeatureByName("Origin")
            if feat is not None and bool(feat.Select2(False, 0)):
                return
        except Exception as e:
            last_error = str(e)

    raise SolidWorksAutomationError(
        f"选择失败：{plane_name_en}（候选：{', '.join(candidates)}）"
        + (f"；最后错误：{last_error}" if last_error else "")
    )


def _find_template(sw_app, template_type: str) -> str:
    template_type = template_type.lower().strip()

    pref_id_map = {
        "part": 8,
        "assembly": 9,
        "drawing": 10,
    }
    pref_id = pref_id_map.get(template_type)
    if pref_id is None:
        raise ValueError(f"未知 template_type: {template_type}")

    try:
        tpl = sw_app.GetUserPreferenceStringValue(pref_id)
        if tpl and Path(tpl).exists():
            return tpl
    except Exception:
        pass

    ext_map = {"part": ".prtdot", "assembly": ".asmdot", "drawing": ".drwdot"}
    ext = ext_map[template_type]

    candidates: list[Path] = []
    program_data = os.environ.get("PROGRAMDATA", r"C:\ProgramData")
    candidates += list(Path(program_data).glob(rf"SolidWorks\SOLIDWORKS 2025\templates\*{ext}"))
    candidates += list(Path(program_data).glob(rf"SOLIDWORKS\SOLIDWORKS 2025\templates\*{ext}"))

    for p in candidates:
        if p.exists():
            return str(p)

    raise SolidWorksAutomationError(
        f"找不到 SolidWorks 默认模板（{ext}）。\n"
        "请在 SolidWorks 中设置默认模板（工具 -> 选项 -> 默认模板），或在脚本中手动指定模板路径。"
    )


def _connect_solidworks(visible: bool = True):
    pythoncom.CoInitialize()
    try:
        sw_app = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        sw_app = win32com.client.Dispatch("SldWorks.Application")
    sw_app.Visible = bool(visible)
    return sw_app


def _new_document(sw_app, template_path: str, doc_type: int):
    model = sw_app.NewDocument(template_path, doc_type, 0.0, 0.0)
    if model is None:
        raise SolidWorksAutomationError(f"NewDocument 失败：{template_path}")
    try:
        if str(template_path).lower().endswith(".prtdot"):
            time.sleep(3.0)
    except Exception:
        pass
    return model


def _enter_sketch(model) -> None:
    model.SketchManager.InsertSketch(True)


def _exit_sketch(model) -> None:
    model.SketchManager.InsertSketch(True)


def _boss_extrude(model, depth_m: float, merge: bool = True) -> None:
    feat = model.FeatureManager.FeatureExtrusion2(
        True,
        False,
        False,
        0,
        0,
        float(depth_m),
        0.0,
        False,
        False,
        False,
        False,
        0.0,
        0.0,
        False,
        False,
        False,
        False,
        bool(merge),
        True,
        True,
        0,
        0.0,
        False,
    )
    if feat is None:
        raise SolidWorksAutomationError("拉伸失败（FeatureExtrusion2 返回 None）")


def _safe_set_entity_name(model, ent, name: str) -> bool:
    try:
        return bool(model.Extension.SetEntityName(ent, str(name)))
    except Exception:
        return False


def _safe_set_feature_name(feat, name: str) -> bool:
    try:
        feat.Name = str(name)
        return True
    except Exception:
        return False


def _name_cylindrical_face_from_feature(model, feat, entity_name: str) -> bool:
    try:
        faces = feat.GetFaces()
    except Exception:
        faces = None
    if not faces:
        return False

    for f in faces:
        try:
            surf = f.GetSurface()
        except Exception:
            surf = None
        if surf is None:
            continue
        try:
            if bool(surf.IsCylinder()):
                return _safe_set_entity_name(model, f, entity_name)
        except Exception:
            continue
    return False


def _face_box_z(face) -> tuple[float, float] | None:
    try:
        b = list(face.GetBox())
        if len(b) >= 6:
            return float(b[2]), float(b[5])
    except Exception:
        pass
    return None


def _name_planar_face_from_feature_by_z(model, feat, entity_name: str, pick: str) -> bool:
    try:
        faces = feat.GetFaces()
    except Exception:
        faces = None
    if not faces:
        return False

    best_face = None
    best_val = None
    for f in faces:
        try:
            surf = f.GetSurface()
        except Exception:
            surf = None
        if surf is None:
            continue
        try:
            if not bool(surf.IsPlane()):
                continue
        except Exception:
            continue

        zbox = _face_box_z(f)
        if zbox is None:
            continue
        zmin, zmax = zbox
        val = zmax if pick == "max" else zmin
        if best_val is None or (val > best_val if pick == "max" else val < best_val):
            best_val = val
            best_face = f

    if best_face is None:
        return False
    return _safe_set_entity_name(model, best_face, entity_name)


def _select_feature_by_name(model, candidates: list[str]) -> str:
    for name in candidates:
        if not name:
            continue
        try:
            feat = model.FeatureByName(str(name))
        except Exception:
            feat = None
        if feat is None:
            continue
        try:
            if bool(feat.Select2(False, 0)):
                return str(name)
        except Exception:
            try:
                if bool(feat.Select2(False)):
                    return str(name)
            except Exception:
                pass
    raise SolidWorksAutomationError(f"找不到/无法选中草图特征（候选：{', '.join(candidates)}）")


def _cut_extrude(model, depth_m: float, thickness_m: float, sketch_candidates: list[str]):
    t = abs(float(thickness_m))
    d_over = (t * 1.5) if t > 0.0 else abs(float(depth_m))
    if d_over <= 0.0:
        raise SolidWorksAutomationError(f"切除深度无效：depth={depth_m}, thickness={thickness_m}")

    try:
        model.ClearSelection2(True)
    except Exception:
        pass 

    sketch_name = _select_feature_by_name(model, sketch_candidates)
    print(f"切除前：草图='{sketch_name}'，深度(溢出)={d_over} m")

    attempts = [(False, False), (True, False), (False, True), (True, True)]
    last_ok = None
    for flip, direction in attempts:
        try:
            model.ClearSelection2(True)
        except Exception:
            pass
        _select_feature_by_name(model, sketch_candidates)
        print(f"尝试切除(ThroughAll)：FlipSide={flip}，Dir={direction}")
        try:
            feat = model.FeatureManager.FeatureCut3(
                True,
                bool(flip),
                bool(direction),
                1,
                0,
                float(d_over),
                0.0,
                False,
                False,
                False,
                False,
                0.0,
                0.0,
                False,
                False,
                False,
                False,
                False,
                False,
                True,
                False,
                False,
                False,
                0,
                0.0,
                False,
            )
            if feat is not None:
                return feat
            last_ok = False
        except Exception as e:
            last_ok = e

    for flip, direction in attempts:
        try:
            model.ClearSelection2(True)
        except Exception:
            pass
        _select_feature_by_name(model, sketch_candidates)
        print(f"尝试切除(Blind)：FlipSide={flip}，Dir={direction}，深度(溢出)={d_over} m")
        try:
            feat = model.FeatureManager.FeatureCut3(
                True,
                bool(flip),
                bool(direction),
                0,
                0,
                float(d_over),
                0.0,
                False,
                False,
                False,
                False,
                0.0,
                0.0,
                False,
                False,
                False,
                False,
                False,
                False,
                True,
                False,
                False,
                False,
                0,
                0.0,
                False,
            )
            if feat is not None:
                return feat
            last_ok = False
        except Exception as e:
            last_ok = e

    for flip, direction in attempts:
        try:
            model.ClearSelection2(True)
        except Exception:
            pass
        _select_feature_by_name(model, sketch_candidates)
        print(f"尝试切除(FeatureCut2 ThroughAll)：FlipSide={flip}，Dir={direction}")
        try:
            feat = model.FeatureManager.FeatureCut2(
                True,
                bool(flip),
                bool(direction),
                1,
                0,
                float(d_over),
                0.0,
                False,
                False,
                False,
                False,
                0.0,
                0.0,
                False,
                False,
                False,
                False,
                False,
                False,
                True,
                False,
                False,
                False,
            )
            if feat is not None:
                return feat
            last_ok = False
        except Exception as e:
            last_ok = e

    for flip, direction in attempts:
        try:
            model.ClearSelection2(True)
        except Exception:
            pass
        _select_feature_by_name(model, sketch_candidates)
        print(f"尝试切除(FeatureCut2 Blind)：FlipSide={flip}，Dir={direction}，深度(溢出)={d_over} m")
        try:
            feat = model.FeatureManager.FeatureCut2(
                True,
                bool(flip),
                bool(direction),
                0,
                0,
                float(d_over),
                0.0,
                False,
                False,
                False,
                False,
                0.0,
                0.0,
                False,
                False,
                False,
                False,
                False,
                False,
                True,
                False,
                False,
                False,
            )
            if feat is not None:
                return feat
            last_ok = False
        except Exception as e:
            last_ok = e

    raise SolidWorksAutomationError(f"切除失败（FeatureCut3 返回 None / 异常：{last_ok}）")


def _save_silent(sw_app, model, file_path: Path) -> Path:
    file_path.parent.mkdir(parents=True, exist_ok=True)
    
    if file_path.exists():
        try:
            os.remove(file_path)
        except Exception as e:
            raise SolidWorksAutomationError(f"无法删除旧文件以覆盖保存：{file_path}；原因：{e}")

    ok = False
    try:
        ok = bool(model.SaveAs3(str(file_path), 0, 1))
    except Exception as e:
        sys.stderr.write(f"警告：SaveAs3 发生异常：{e}\n")

    if (not ok) and file_path.exists():
        sys.stderr.write(f"提示：SaveAs3 返回失败，但磁盘已存在文件，视为成功：{file_path}\n")
        ok = True

    if not ok:
        raise SolidWorksAutomationError(f"保存失败：{file_path}")
        
    try:
        title = model.GetTitle()
        sw_app.CloseDoc(title)
    except Exception as e:
        sys.stderr.write(f"警告：关闭文档失败：{e}\n")

    return file_path


def _create_crank_part(sw_app, part_template: str, out_path: Path, p: MechanismParams) -> Path:
    model = _new_document(sw_app, part_template, 0)
    model.ClearSelection2(True)

    _select_plane(model, "Front Plane")
    _enter_sketch(model)

    x0 = -p.end_margin
    x1 = p.R + p.end_margin
    half_w = p.bar_width / 2.0
    model.SketchManager.CreateCornerRectangle(x0, -half_w, 0.0, x1, half_w, 0.0)
    _exit_sketch(model)

    _boss_extrude(model, p.bar_thickness, merge=True)
    try:
        boss = model.FeatureByName("Boss-Extrude1")
    except Exception:
        boss = None
    if boss is not None:
        _name_planar_face_from_feature_by_z(model, boss, "Crank_Side", "max")
        _name_planar_face_from_feature_by_z(model, boss, "Crank_Side_Back", "min")
    model.ClearSelection2(True)

    r = p.hole_d / 2.0
    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    model.SketchManager.CreateCircleByRadius(0.0, 0.0, 0.0, r)
    _exit_sketch(model)
    cut1 = _cut_extrude(model, p.bar_thickness, p.bar_thickness, ["Sketch2", "草图2"])
    _safe_set_feature_name(cut1, "Crank_In")
    _name_cylindrical_face_from_feature(model, cut1, "Crank_In")

    model.ClearSelection2(True)
    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    model.SketchManager.CreateCircleByRadius(p.R, 0.0, 0.0, r)
    _exit_sketch(model)
    cut2 = _cut_extrude(model, p.bar_thickness, p.bar_thickness, ["Sketch3", "草图3"])
    _safe_set_feature_name(cut2, "Crank_Out")
    _name_cylindrical_face_from_feature(model, cut2, "Crank_Out")

    return _save_silent(sw_app, model, out_path)


def _create_rod_part(sw_app, part_template: str, out_path: Path, p: MechanismParams) -> Path:
    model = _new_document(sw_app, part_template, 0)
    model.ClearSelection2(True)

    _select_plane(model, "Front Plane")
    _enter_sketch(model)

    x0 = -p.end_margin
    x1 = p.L + p.end_margin
    half_w = p.bar_width / 2.0
    model.SketchManager.CreateCornerRectangle(x0, -half_w, 0.0, x1, half_w, 0.0)
    _exit_sketch(model)

    _boss_extrude(model, p.bar_thickness, merge=True)
    try:
        boss = model.FeatureByName("Boss-Extrude1")
    except Exception:
        boss = None
    if boss is not None:
        _name_planar_face_from_feature_by_z(model, boss, "Link_Side", "max")
        _name_planar_face_from_feature_by_z(model, boss, "Link_Side_Back", "min")
    model.ClearSelection2(True)

    r = p.hole_d / 2.0
    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    model.SketchManager.CreateCircleByRadius(0.0, 0.0, 0.0, r)
    _exit_sketch(model)
    cut1 = _cut_extrude(model, p.bar_thickness, p.bar_thickness, ["Sketch2", "草图2"])
    _safe_set_feature_name(cut1, "Link_In")
    _name_cylindrical_face_from_feature(model, cut1, "Link_In")

    model.ClearSelection2(True)
    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    model.SketchManager.CreateCircleByRadius(p.L, 0.0, 0.0, r)
    _exit_sketch(model)
    cut2 = _cut_extrude(model, p.bar_thickness, p.bar_thickness, ["Sketch3", "草图3"])
    _safe_set_feature_name(cut2, "Link_Out")
    _name_cylindrical_face_from_feature(model, cut2, "Link_Out")

    return _save_silent(sw_app, model, out_path)


def _create_slider_part(sw_app, part_template: str, out_path: Path, p: MechanismParams) -> Path:
    model = _new_document(sw_app, part_template, 0)
    model.ClearSelection2(True)

    _select_plane(model, "Front Plane")
    _enter_sketch(model)

    half = p.slider_size / 2.0
    model.SketchManager.CreateCenterRectangle(0.0, 0.0, 0.0, half, half, 0.0)
    _exit_sketch(model)

    _boss_extrude(model, p.slider_thickness, merge=True)
    try:
        boss = model.FeatureByName("Boss-Extrude1")
    except Exception:
        boss = None
    if boss is not None:
        _name_planar_face_from_feature_by_z(model, boss, "Slider_Bottom", "min")
    model.ClearSelection2(True)

    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    r = p.hole_d / 2.0
    model.SketchManager.CreateCircleByRadius(0.0, 0.0, 0.0, r)
    _exit_sketch(model)
    cut1 = _cut_extrude(model, p.slider_thickness, p.slider_thickness, ["Sketch2", "草图2"])
    _safe_set_feature_name(cut1, "Slider_Joint")
    _name_cylindrical_face_from_feature(model, cut1, "Slider_Joint")

    return _save_silent(sw_app, model, out_path)


def _create_base_part(sw_app, part_template: str, out_path: Path, p: MechanismParams) -> Path:
    model = _new_document(sw_app, part_template, 0)
    model.ClearSelection2(True)

    base_len = p.base_len
    base_w = max(p.slider_size * 2.0, 0.06)
    base_t = p.bar_thickness
    slot_w = p.slot_width
    slot_len = min(base_len * 0.75, base_len - 0.04)
    slot_cx = base_len * 0.5 + p.pivot_x * 0.0
    pin_r = p.pin_d / 2.0

    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    model.SketchManager.CreateCornerRectangle(0.0, -base_w / 2.0, 0.0, base_len, base_w / 2.0, 0.0)
    _exit_sketch(model)

    _boss_extrude(model, base_t, merge=True)
    try:
        boss = model.FeatureByName("Boss-Extrude1")
    except Exception:
        boss = None
    if boss is not None:
        _name_planar_face_from_feature_by_z(model, boss, "Slide_Way", "max")
        _name_planar_face_from_feature_by_z(model, boss, "Base_Front", "max")

    model.ClearSelection2(True)
    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    model.SketchManager.CreateCenterRectangle(slot_cx, 0.0, 0.0, slot_cx + slot_len / 2.0, slot_w / 2.0, 0.0)
    _exit_sketch(model)
    _cut_extrude(model, base_t, base_t, ["Sketch2", "草图2"])

    model.ClearSelection2(True)
    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    r = p.hole_d / 2.0
    model.SketchManager.CreateCircleByRadius(p.pivot_x, 0.0, 0.0, r)
    _exit_sketch(model)
    cut1 = _cut_extrude(model, base_t, base_t, ["Sketch3", "草图3"])
    _safe_set_feature_name(cut1, "Fixed_Hole")
    _name_cylindrical_face_from_feature(model, cut1, "Fixed_Hole")

    model.ClearSelection2(True)
    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    model.SketchManager.CreateCircleByRadius(p.pivot_x, 0.0, 0.0, pin_r)
    _exit_sketch(model)
    _boss_extrude(model, p.pin_len, merge=True)

    return _save_silent(sw_app, model, out_path)


def _create_pin_part(sw_app, part_template: str, out_path: Path, p: MechanismParams) -> Path:
    model = _new_document(sw_app, part_template, 0)
    model.ClearSelection2(True)

    _select_plane(model, "Front Plane")
    _enter_sketch(model)
    model.SketchManager.CreateCircleByRadius(0.0, 0.0, 0.0, p.pin_d / 2.0)
    _exit_sketch(model)

    _boss_extrude(model, p.pin_len, merge=True)
    return _save_silent(sw_app, model, out_path)


def _asm_select_by_id2(model, name: str, type_str: str, append: bool) -> bool:
    try:
        return bool(model.Extension.SelectByID2(str(name), str(type_str), 0.0, 0.0, 0.0, bool(append), 0, None, 0))
    except Exception as e:
        sys.stderr.write(f"警告：装配体 SelectByID2 失败（{type_str}='{name}'）：{e}\n")
        return False


def _add_mate_safe(asm, mate_type: int, mate_align: int) -> None:
    try:
        mate = asm.AddMate5(mate_type, mate_align, False, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, False, False, 0.0, 0)
        if mate is None:
            sys.stderr.write("警告：AddMate5 返回 None\n")
    except Exception as e:
        sys.stderr.write(f"警告：AddMate5 异常：{e}\n")


def _set_component_color(comp, r: float, g: float, b: float) -> None:
    arr = [float(r), float(g), float(b), 1.0, 1.0, 0.0, 0.0, 0.0, 0.0]
    try:
        comp.MaterialPropertyValues = arr
        return
    except Exception:
        pass
    try:
        comp.MaterialPropertyValues2 = arr
    except Exception:
        pass


def _create_assembly(
    sw_app,
    asm_template: str,
    out_path: Path,
    base_path: Path,
    crank_path: Path,
    rod_path: Path,
    slider_path: Path,
    pin_b_path: Path,
    pin_c_path: Path,
    p: MechanismParams,
) -> Path:
    model = _new_document(sw_app, asm_template, 0)
    asm = model
    model.ClearSelection2(True)

    try:
        pivot_x = float(p.pivot_x)
        crank_x = pivot_x
        rod_x = pivot_x + float(p.R)
        slider_x = pivot_x + float(p.R + p.L)

        base_comp = asm.AddComponent5(str(base_path), 0, "", False, "", 0.0, 0.0, 0.0)
        crank_comp = asm.AddComponent5(str(crank_path), 0, "", False, "", float(crank_x), 0.0, 0.0)
        rod_comp = asm.AddComponent5(str(rod_path), 0, "", False, "", float(rod_x), 0.0, 0.0)
        slider_comp = asm.AddComponent5(str(slider_path), 0, "", False, "", float(slider_x), 0.0, 0.0)

        pin_b_comp = asm.AddComponent5(str(pin_b_path), 0, "", False, "", float(rod_x), 0.0, 0.0)
        pin_c_comp = asm.AddComponent5(str(pin_c_path), 0, "", False, "", float(slider_x), 0.0, 0.0)

        _set_component_color(base_comp, 0.7, 0.7, 0.7)
        _set_component_color(crank_comp, 0.9, 0.2, 0.2)
        _set_component_color(rod_comp, 0.2, 0.8, 0.2)
        _set_component_color(slider_comp, 0.2, 0.4, 0.9)
        _set_component_color(pin_b_comp, 0.9, 0.8, 0.2)
        _set_component_color(pin_c_comp, 0.9, 0.5, 0.1)
    except Exception as e:
        raise SolidWorksAutomationError(f"装配体插入零部件或设置位移失败：{e}")

    try:
        model.ForceRebuild3(True)
    except Exception:
        pass

    try:
        model.ShowNamedView2("*Isometric", 7)
        model.ViewZoomtofit2()
    except Exception:
        try:
            model.ViewIso()
            model.ViewZoomtofit2()
        except Exception:
            pass

    return _save_silent(sw_app, model, out_path)


def main() -> int:
    p = MechanismParams()
    out_dir = Path(__file__).resolve().parent / "generated_mechanism"
    out_dir.mkdir(parents=True, exist_ok=True)
    for ext in ("*.SLDPRT", "*.SLDASM"):
        for fp in out_dir.glob(ext):
            try:
                os.remove(fp)
            except Exception:
                pass

    base_path = out_dir / "Base.SLDPRT"
    crank_path = out_dir / "Crank.SLDPRT"
    rod_path = out_dir / "ConnectingRod.SLDPRT"
    slider_path = out_dir / "Slider.SLDPRT"
    pin_b_path = out_dir / "Pin_B.SLDPRT"
    pin_c_path = out_dir / "Pin_C.SLDPRT"
    asm_path = out_dir / "Assembly.SLDASM"

    sw_app = None
    try:
        sw_app = _connect_solidworks(visible=True)
        part_template = _find_template(sw_app, "part")
        asm_template = _find_template(sw_app, "assembly")

        base_path = _create_base_part(sw_app, part_template, base_path, p)
        crank_path = _create_crank_part(sw_app, part_template, crank_path, p)
        rod_path = _create_rod_part(sw_app, part_template, rod_path, p)
        slider_path = _create_slider_part(sw_app, part_template, slider_path, p)
        pin_b_path = _create_pin_part(sw_app, part_template, pin_b_path, p)
        pin_c_path = _create_pin_part(sw_app, part_template, pin_c_path, p)
        asm_path = _create_assembly(
            sw_app,
            asm_template,
            asm_path,
            base_path,
            crank_path,
            rod_path,
            slider_path,
            pin_b_path,
            pin_c_path,
            p,
        )

        _send_sw_msg(sw_app, f"曲柄滑块机构已生成：{asm_path}")
        return 0
    except SolidWorksAutomationError as e:
        if sw_app is not None:
            _send_sw_msg(sw_app, f"生成停止：{e}")
        sys.stderr.write(f"\n[提示] 自动化操作终止: {e}\n")
        return 1
    except Exception as e:
        detail = traceback.format_exc()
        if sw_app is not None:
            _send_sw_msg(sw_app, f"生成失败：{e}\n\n{detail}")
        sys.stderr.write(f"\n[错误] 发生未知异常：\n{detail}\n")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

