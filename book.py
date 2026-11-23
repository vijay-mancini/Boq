# -*- coding: utf-8 -*-
"""
Stage 4 - Revit Data Extraction Module (READ-ONLY)
Runs inside Revit (pyRevit / IronPython 2.7). Expects globals: doc, DB.
Writes a JSON snapshot; performs NO model modification (no Transaction is opened).

Scope rule: substructure = everything below the cut level ffl_m read from output/run_config.json.
All geometry is converted to metres / m2 / m3 through DB.UnitUtils (not hard-coded
factors); a unit self-check is stored in the snapshot meta block.
Extraction only records facts. Classification, measurement rules and NRM2 rounding
belong to later modules.
"""
import os
import json
import time
import clr  # noqa
import System

# ----------------------------------------------------------------------------
# CONFIG  (project independent; per-run values come from output/run_config.json)
# ----------------------------------------------------------------------------
PROJECT_DIR = globals().get("PROJECT_DIR", "D:/C/RevitApi/BOQ/book_ref_detail")
OUT_PATH = os.path.join(PROJECT_DIR, "output", "extraction.json")
_cfg = {}
try:
    with open(os.path.join(PROJECT_DIR, "output", "run_config.json")) as _f:
        _cfg = json.load(_f)
except Exception:
    pass
FFL_M = float(_cfg.get("ffl_m", 0.0))                 # substructure = everything below this level (metres)
TOL_M = 0.005                                          # boundary tolerance, metres
DATUM_SUSPECT_Z_M = FFL_M - float(_cfg.get("datum_suspect_below_m", 6.0))   # elements starting lower than this are flagged

CATEGORIES = [
    ("Walls", DB.BuiltInCategory.OST_Walls),
    ("Floors", DB.BuiltInCategory.OST_Floors),
    ("Structural Columns", DB.BuiltInCategory.OST_StructuralColumns),
    ("Structural Framing", DB.BuiltInCategory.OST_StructuralFraming),
    ("Structural Foundations", DB.BuiltInCategory.OST_StructuralFoundation),
    ("Hardscape", DB.BuiltInCategory.OST_Hardscape),
    ("Generic Models", DB.BuiltInCategory.OST_GenericModel),
    ("Roofs", DB.BuiltInCategory.OST_Roofs),
    ("Stairs", DB.BuiltInCategory.OST_Stairs),
    ("Site", DB.BuiltInCategory.OST_Site),
]
SCALE = 1000000.0  # none (kept for clarity)

# ----------------------------------------------------------------------------
# UNIT HELPERS (internal feet -> metric via UnitUtils)
# ----------------------------------------------------------------------------

def elem_name(e):
    """Element name that works on every IronPython/Revit build (Name property, else Element.Name getter)."""
    try:
        return e.Name
    except Exception:
        pass
    try:
        return DB.Element.Name.GetValue(e)
    except Exception:
        return ""

def to_m(v):
    return DB.UnitUtils.ConvertFromInternalUnits(v, DB.UnitTypeId.Meters)

def to_m2(v):
    return DB.UnitUtils.ConvertFromInternalUnits(v, DB.UnitTypeId.SquareMeters)

def to_m3(v):
    return DB.UnitUtils.ConvertFromInternalUnits(v, DB.UnitTypeId.CubicMeters)

FT_OF_M = 1.0 / to_m(1.0)   # internal feet per metre

def eid(x):
    try:
        return int(x.Value)
    except Exception:
        return int(x.IntegerValue)

def rnd(v, n=6):
    return None if v is None else round(v, n)

def clean_str(s):
    """Make a .NET/IronPython string JSON-safe: non-ASCII chars become \\uXXXX text."""
    out = []
    for ch in s:
        o = ord(ch)
        out.append(ch if o < 128 else "\\u%04x" % o)
    return "".join(out)

def scrub(o):
    if isinstance(o, dict):
        return dict((scrub(k), scrub(v)) for k, v in o.items())
    if isinstance(o, (list, tuple)):
        return [scrub(x) for x in o]
    if isinstance(o, basestring):
        return clean_str(o)
    return o

# ----------------------------------------------------------------------------
# GENERIC READERS
# ----------------------------------------------------------------------------
def type_name(el):
    try:
        t = doc.GetElement(el.GetTypeId())
        if t is None:
            return "", ""
        p = t.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME) or t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        n = p.AsString() if p else ""
        f = t.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        fam = f.AsString() if f else ""
        return (fam or ""), (n or "")
    except Exception:
        return "", ""

def param_str(p):
    try:
        st = p.StorageType
        if st == DB.StorageType.String:
            return p.AsString()
        if st == DB.StorageType.ElementId:
            return "id:%s" % eid(p.AsElementId())
        s = p.AsValueString()
        if s is not None:
            return s
        if st == DB.StorageType.Integer:
            return str(p.AsInteger())
        if st == DB.StorageType.Double:
            return str(p.AsDouble())
    except Exception:
        pass
    return None

def all_params(el):
    out = {}
    try:
        for p in el.Parameters:
            try:
                n = p.Definition.Name
                v = param_str(p)
                if v not in (None, ""):
                    out[n] = v
            except Exception:
                pass
    except Exception:
        pass
    return out

def bip_len(el, bip):
    p = el.get_Parameter(bip)
    return to_m(p.AsDouble()) if (p is not None and p.HasValue) else None

def bip_area(el, bip):
    p = el.get_Parameter(bip)
    return to_m2(p.AsDouble()) if (p is not None and p.HasValue) else None

def bip_vol(el, bip):
    p = el.get_Parameter(bip)
    return to_m3(p.AsDouble()) if (p is not None and p.HasValue) else None

def level_info(el):
    lid = None
    for bip in (DB.BuiltInParameter.LEVEL_PARAM, DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM, DB.BuiltInParameter.WALL_BASE_CONSTRAINT):
        p = el.get_Parameter(bip)
        if p is not None and p.HasValue and p.StorageType == DB.StorageType.ElementId:
            i = p.AsElementId()
            if eid(i) > 0:
                lid = i
                break
    if lid is None:
        try:
            if eid(el.LevelId) > 0:
                lid = el.LevelId
        except Exception:
            pass
    if lid is None:
        return None, None
    lv = doc.GetElement(lid)
    try:
        name = elem_name(lv)
    except Exception:
        name = None
    return name, rnd(to_m(lv.Elevation), 4)

def group_info(el):
    try:
        gid = el.GroupId
        if eid(gid) > 0:
            g = doc.GetElement(gid)
            gt = doc.GetElement(g.GetTypeId())
            p = gt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            return eid(gid), (p.AsString() if p else None)
    except Exception:
        pass
    return None, None

def material_info(el):
    out = {}
    try:
        for mid in el.GetMaterialIds(False):
            m = doc.GetElement(mid)
            if m is None:
                continue
            try:
                v = to_m3(el.GetMaterialVolume(mid))
            except Exception:
                v = None
            out[elem_name(m)] = rnd(v)
    except Exception:
        pass
    return out

def phase_names(el):
    r = {}
    try:
        p = el.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
        if p is not None and p.HasValue:
            r["created"] = param_str(p)
        p = el.get_Parameter(DB.BuiltInParameter.PHASE_DEMOLISHED)
        if p is not None and p.HasValue:
            r["demolished"] = param_str(p)
    except Exception:
        pass
    return r

# ----------------------------------------------------------------------------
# GEOMETRY
# ----------------------------------------------------------------------------
def get_solids(el):
    opts = DB.Options()
    opts.ComputeReferences = False
    opts.IncludeNonVisibleObjects = False
    opts.DetailLevel = DB.ViewDetailLevel.Fine
    solids = []
    ge = el.get_Geometry(opts)
    if ge is None:
        return solids

    def walk(g):
        for o in g:
            if isinstance(o, DB.Solid):
                if o.Volume > 1e-9 and o.Faces.Size > 0:
                    solids.append(o)
            elif isinstance(o, DB.GeometryInstance):
                sub = o.GetInstanceGeometry()
                if sub is not None:
                    walk(sub)
    walk(ge)
    return solids

def make_box(bb, z_top_ft):
    pad = 1.0
    x0, x1 = bb.Min.X - pad, bb.Max.X + pad
    y0, y1 = bb.Min.Y - pad, bb.Max.Y + pad
    z0 = bb.Min.Z - 1.0
    p = [DB.XYZ(x0, y0, z0), DB.XYZ(x1, y0, z0), DB.XYZ(x1, y1, z0), DB.XYZ(x0, y1, z0)]
    loop = DB.CurveLoop()
    for i in range(4):
        loop.Append(DB.Line.CreateBound(p[i], p[(i + 1) % 4]))
    from System.Collections.Generic import List
    loops = List[DB.CurveLoop]()
    loops.Add(loop)
    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(loops, DB.XYZ.BasisZ, z_top_ft - z0)

def face_metrics(solid):
    """Areas (internal ft2) by face orientation for a solid: bottom, top, side, other."""
    bottom = top = side = other = 0.0
    for f in solid.Faces:
        a = f.Area
        if isinstance(f, DB.PlanarFace):
            nz = f.FaceNormal.Z
            if nz < -0.99:
                bottom += a
            elif nz > 0.99:
                top += a
            elif abs(nz) < 0.01:
                side += a
            else:
                other += a
        else:
            other += a
    return bottom, top, side, other

def wall_normal(el):
    """Unit normal (XYZ) of a straight wall in plan, or None."""
    try:
        loc = el.Location
        if isinstance(loc, DB.LocationCurve) and isinstance(loc.Curve, DB.Line):
            d = loc.Curve.Direction
            n = DB.XYZ(-d.Y, d.X, 0.0)
            return n.Normalize()
    except Exception:
        pass
    return None

def face_pos_neg(solid, wn):
    """Areas (ft2) of planar vertical faces parallel to the wall faces, split by side."""
    pos = neg = 0.0
    for f in solid.Faces:
        if isinstance(f, DB.PlanarFace):
            fn = f.FaceNormal
            if abs(fn.Z) < 0.01:
                dp = fn.X * wn.X + fn.Y * wn.Y
                if dp > 0.99:
                    pos += f.Area
                elif dp < -0.99:
                    neg += f.Area
    return pos, neg

def below_metrics(el, bb, status, issues):
    """Returns dict of totals for full solid and portion at/below FFL."""
    ffl_ft = FFL_M * FT_OF_M
    solids = get_solids(el)
    res = {"solid_count": len(solids)}
    if not solids:
        res.update({"vol_total_m3": None, "vol_below_ffl_m3": None})
        return res
    tot = 0.0
    parts = {"vol": 0.0, "bottom": 0.0, "top": 0.0, "side": 0.0, "other": 0.0}
    zmin = 1e12
    zmax = -1e12
    comps = {}
    wn = wall_normal(el)
    wpos = wneg = 0.0
    for s in solids:
        tot += s.Volume
        if status == "below":
            sub = s
        elif status == "straddle":
            try:
                box = make_box(bb, ffl_ft)
                sub = DB.BooleanOperationsUtils.ExecuteBooleanOperation(s, box, DB.BooleanOperationsType.Intersect)
            except Exception as ex:
                issues.append("boolean_cut_failed:%s" % str(ex)[:80])
                continue
            if sub is None or sub.Volume <= 1e-12:
                continue
        else:
            continue
        parts["vol"] += sub.Volume
        if wn is not None:
            _p, _n = face_pos_neg(sub, wn)
            wpos += _p
            wneg += _n
        b, t, sd, o = face_metrics(sub)
        parts["bottom"] += b
        parts["top"] += t
        parts["side"] += sd
        parts["other"] += o
        szmin, szmax = 1e12, -1e12
        try:
            # exact z-range from edge end points (bounding-box corners can be loose)
            for ed in sub.Edges:
                c = ed.AsCurve()
                for k in (0, 1):
                    z = c.GetEndPoint(k).Z
                    zmin = min(zmin, z)
                    zmax = max(zmax, z)
                    szmin = min(szmin, z)
                    szmax = max(szmax, z)
        except Exception:
            pass
        # per-solid component record (material + z-range), aggregated
        try:
            mid = s.Faces.get_Item(0).MaterialElementId
            mname = elem_name(doc.GetElement(mid)) if eid(mid) > 0 else "(none)"
        except Exception:
            mname = "(unknown)"
        key = (mname, round(to_m(szmin), 3) if szmin < 1e11 else None, round(to_m(szmax), 3) if szmax > -1e11 else None)
        comp = comps.setdefault(key, {"n": 0, "vol": 0.0, "bottom": 0.0, "top": 0.0, "side": 0.0})
        comp["n"] += 1
        comp["vol"] += sub.Volume
        comp["bottom"] += b
        comp["top"] += t
        comp["side"] += sd
    res["vol_total_m3"] = rnd(to_m3(tot))
    res["vol_below_ffl_m3"] = rnd(to_m3(parts["vol"]))
    res["below_bottom_area_m2"] = rnd(to_m2(parts["bottom"]))
    res["below_top_area_m2"] = rnd(to_m2(parts["top"]))
    res["below_side_area_m2"] = rnd(to_m2(parts["side"]))
    res["below_other_area_m2"] = rnd(to_m2(parts["other"]))
    # components (material + z-range) of the below-FFL part; tiny ones (< 1 litre) are lumped
    clist = []
    minor = 0.0
    for key in sorted(comps, key=lambda k: (k[1] if k[1] is not None else 0, k[2] if k[2] is not None else 0, k[0])):
        c = comps[key]
        v3 = to_m3(c["vol"])
        if v3 < 0.001:
            minor += v3
            continue
        clist.append({"material": key[0], "z_min_m": key[1], "z_max_m": key[2], "solids": c["n"],
                      "vol_m3": rnd(v3), "bottom_area_m2": rnd(to_m2(c["bottom"])),
                      "top_area_m2": rnd(to_m2(c["top"])), "side_area_m2": rnd(to_m2(c["side"]))})
    if wn is not None or isinstance(el, DB.Wall):
        res["wall_face_area_pos_m2"] = rnd(to_m2(wpos)) if wn is not None else None
        res["wall_face_area_neg_m2"] = rnd(to_m2(wneg)) if wn is not None else None
        res["wall_is_straight"] = wn is not None
    res["components_below_ffl"] = clist
    res["minor_components_m3"] = rnd(minor)
    if zmin < 1e11:
        res["below_z_min_m"] = rnd(to_m(zmin), 4)
        res["below_z_max_m"] = rnd(to_m(zmax), 4)
    return res

def location_info(el):
    d = {}
    try:
        loc = el.Location
        if isinstance(loc, DB.LocationCurve):
            c = loc.Curve
            a, b = c.GetEndPoint(0), c.GetEndPoint(1)
            d["curve_length_m"] = rnd(to_m(c.Length))
            d["start_m"] = [rnd(to_m(a.X), 4), rnd(to_m(a.Y), 4), rnd(to_m(a.Z), 4)]
            d["end_m"] = [rnd(to_m(b.X), 4), rnd(to_m(b.Y), 4), rnd(to_m(b.Z), 4)]
            d["curve_type"] = c.GetType().Name
        elif isinstance(loc, DB.LocationPoint):
            p = loc.Point
            d["point_m"] = [rnd(to_m(p.X), 4), rnd(to_m(p.Y), 4), rnd(to_m(p.Z), 4)]
            try:
                d["rotation_rad"] = rnd(loc.Rotation, 6)
            except Exception:
                pass
    except Exception:
        pass
    return d

def dims(el):
    d = {}
    B = DB.BuiltInParameter
    for key, bip, kind in (
        ("length_m", B.CURVE_ELEM_LENGTH, "len"),
        ("wall_width_m", B.WALL_ATTR_WIDTH_PARAM, "len"),
        ("unconnected_height_m", B.WALL_USER_HEIGHT_PARAM, "len"),
        ("base_offset_m", B.WALL_BASE_OFFSET, "len"),
        ("top_offset_m", B.WALL_TOP_OFFSET, "len"),
        ("floor_thickness_m", B.FLOOR_ATTR_THICKNESS_PARAM, "len"),
        ("floor_offset_m", B.FLOOR_HEIGHTABOVELEVEL_PARAM, "len"),
        ("area_computed_m2", B.HOST_AREA_COMPUTED, "area"),
        ("host_volume_m3", B.HOST_VOLUME_COMPUTED, "vol"),
    ):
        try:
            if kind == "len":
                v = bip_len(el, bip)
            elif kind == "area":
                v = bip_area(el, bip)
            else:
                v = bip_vol(el, bip)
        except Exception:
            v = None
        if v is not None:
            d[key] = rnd(v)
    # Wall type width via type element, if instance param absent
    try:
        if "wall_width_m" not in d and isinstance(el, DB.Wall):
            d["wall_width_m"] = rnd(to_m(el.Width))
    except Exception:
        pass
    return d

# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def main():
    t0 = time.time()
    # unit self-check: 1 ft must equal 0.3048 m exactly
    unit_check = {
        "one_foot_in_m": to_m(1.0),
        "ok": abs(to_m(1.0) - 0.3048) < 1e-9,
        "one_ft3_in_m3": to_m3(1.0),
        "one_ft2_in_m2": to_m2(1.0),
        "ffl_internal_ft": FFL_M * FT_OF_M,
    }
    lev = {}
    for l in DB.FilteredElementCollector(doc).OfClass(DB.Level):
        lev[elem_name(l)] = rnd(to_m(l.Elevation), 4)

    elements = []
    skipped = {"above_ffl": 0, "no_bbox": 0, "excluded_dup": 0}
    validation = []
    seen_ids = set()
    ffl_lo = FFL_M - TOL_M
    ffl_hi = FFL_M + TOL_M

    for cname, bic in CATEGORIES:
        col = DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        for el in col:
            i = eid(el.Id)
            if i in seen_ids:
                skipped["excluded_dup"] += 1
                validation.append({"element_id": i, "severity": "ERROR", "code": "DUPLICATE_ELEMENT", "msg": "Element returned by more than one category collector"})
                continue
            seen_ids.add(i)
            bb = el.get_BoundingBox(None)
            if bb is None:
                skipped["no_bbox"] += 1
                validation.append({"element_id": i, "severity": "WARN", "code": "NO_BBOX", "category": cname,
                                   "msg": "No bounding box; cannot be tested against FFL"})
                continue
            zmin = to_m(bb.Min.Z)
            zmax = to_m(bb.Max.Z)
            if zmin >= ffl_lo:
                skipped["above_ffl"] += 1
                continue
            status = "below" if zmax <= ffl_hi else "straddle"
            fam, tname = type_name(el)
            issues = []
            geo = below_metrics(el, bb, status, issues)
            lvname, lvelev = level_info(el)
            gid, gname = group_info(el)
            rec = {
                "revit_class": el.GetType().Name,
                "element_id": i,
                "unique_id": el.UniqueId,
                "category": cname,
                "family": fam,
                "type": tname,
                "level": lvname,
                "level_elevation_m": lvelev,
                "workset": None,
                "mark": (param_str(el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)) if el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK) else None),
                "comments": (param_str(el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)) if el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS) else None),
                "group_id": gid,
                "group_name": gname,
                "phase": phase_names(el),
                "materials_m3": material_info(el),
                "bbox_m": {"min": [rnd(to_m(bb.Min.X), 4), rnd(to_m(bb.Min.Y), 4), rnd(zmin, 4)],
                           "max": [rnd(to_m(bb.Max.X), 4), rnd(to_m(bb.Max.Y), 4), rnd(zmax, 4)]},
                "ffl_status": status,
                "dimensions": dims(el),
                "location": location_info(el),
                "geometry": geo,
                "instance_params": all_params(el),
                "type_params": all_params(doc.GetElement(el.GetTypeId())) if eid(el.GetTypeId()) > 0 else {},
                "flags": [],
            }
            if doc.IsWorkshared:
                try:
                    rec["workset"] = doc.GetWorksetTable().GetWorkset(el.WorksetId).Name
                except Exception:
                    pass
            # ---- per-element validation ----
            if zmin < DATUM_SUSPECT_Z_M:
                rec["flags"].append("DATUM_SUSPECT")
                validation.append({"element_id": i, "severity": "WARN", "code": "DATUM_SUSPECT", "category": cname, "type": tname,
                                   "msg": "bbox z min %.2f m is below %.1f m (possible wrong datum)" % (zmin, DATUM_SUSPECT_Z_M)})
            if geo.get("solid_count", 0) == 0:
                rec["flags"].append("NO_SOLID")
                validation.append({"element_id": i, "severity": "ERROR", "code": "NO_SOLID", "category": cname, "type": tname,
                                   "msg": "No solid geometry; quantity cannot be measured"})
            else:
                hv = rec["dimensions"].get("host_volume_m3")
                vt = geo.get("vol_total_m3")
                if hv is not None and vt is not None and hv > 0:
                    diff = abs(hv - vt) / hv
                    rec["volume_check_rel_diff"] = rnd(diff, 5)
                    if diff > 0.01:
                        rec["flags"].append("VOLUME_MISMATCH")
                        validation.append({"element_id": i, "severity": "WARN", "code": "VOLUME_MISMATCH", "category": cname, "type": tname,
                                           "msg": "Geometry volume %.4f vs Revit host volume %.4f m3 (%.1f%%)" % (vt, hv, diff * 100)})
                if geo.get("vol_below_ffl_m3") is not None and geo["vol_below_ffl_m3"] < 1e-6:
                    rec["flags"].append("ZERO_VOLUME_BELOW_FFL")
            bz = geo.get("below_z_max_m")
            if bz is not None and bz > FFL_M + 2 * TOL_M:
                issues.append("CUT_CHECK_FAILED: below-FFL part reaches z=%.3f m" % bz)
                rec["flags"].append("CUT_CHECK_FAILED")
            for s in issues:
                validation.append({"element_id": i, "severity": "ERROR", "code": "GEOMETRY_ERROR", "category": cname, "type": tname, "msg": s})
            rec["flags"] = rec["flags"] + list(issues)
            elements.append(rec)

    # Toposolids: ground references (not substructure quantities)
    topo = []
    for el in (DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Toposolid).WhereElementIsNotElementType() if hasattr(DB.BuiltInCategory, "OST_Toposolid") else []):
        bb = el.get_BoundingBox(None)
        fam, tname = type_name(el)
        topo.append({
            "element_id": eid(el.Id), "unique_id": el.UniqueId, "family": fam, "type": tname,
            "level": level_info(el)[0],
            "host_volume_m3": rnd(bip_vol(el, DB.BuiltInParameter.HOST_VOLUME_COMPUTED)),
            "bbox_m": ({"min": [rnd(to_m(bb.Min.X), 3), rnd(to_m(bb.Min.Y), 3), rnd(to_m(bb.Min.Z), 3)],
                        "max": [rnd(to_m(bb.Max.X), 3), rnd(to_m(bb.Max.Y), 3), rnd(to_m(bb.Max.Z), 3)]} if bb else None),
        })

    summary = {}
    for r in elements:
        k = "%s | %s | %s" % (r["category"], r["type"], r["ffl_status"])
        s = summary.setdefault(k, {"count": 0, "vol_total_m3": 0.0, "vol_below_ffl_m3": 0.0})
        s["count"] += 1
        s["vol_total_m3"] += (r["geometry"].get("vol_total_m3") or 0.0)
        s["vol_below_ffl_m3"] += (r["geometry"].get("vol_below_ffl_m3") or 0.0)
    for s in summary.values():
        s["vol_total_m3"] = round(s["vol_total_m3"], 4)
        s["vol_below_ffl_m3"] = round(s["vol_below_ffl_m3"], 4)

    snapshot = {
        "meta": {
            "document": doc.Title,
            "revit_version": doc.Application.VersionNumber,
            "extracted_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "read_only": True,
            "ffl_m": FFL_M, "tolerance_m": TOL_M,
            "units": "metres, m2, m3 (converted from Revit internal feet via UnitUtils)",
            "unit_check": unit_check,
            "levels_m": lev,
            "element_count": len(elements),
            "skipped": skipped,
            "runtime_s": round(time.time() - t0, 2),
        },
        "elements": elements,
        "toposolids": topo,
        "validation": validation,
        "summary_by_type": summary,
    }
    d = os.path.dirname(OUT_PATH)
    if not os.path.isdir(d):
        os.makedirs(d)
    snapshot = scrub(snapshot)
    with open(OUT_PATH, "w") as f:
        f.write(json.dumps(snapshot, indent=1, sort_keys=True))
    return snapshot

_snap = main()
print("elements:", _snap["meta"]["element_count"], "skipped:", _snap["meta"]["skipped"],
      "validation items:", len(_snap["validation"]), "runtime_s:", _snap["meta"]["runtime_s"])
print("unit_check:", _snap["meta"]["unit_check"])
