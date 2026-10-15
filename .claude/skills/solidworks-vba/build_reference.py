# Builds reference/*.md for the skill from the locally installed SolidWorks API help.
# Run (Windows python):  python build_reference.py <mshc_all_dir> <cw_dir> <out_dir>
#   mshc_all_dir = unpacked APIHelpViewer.mshc (api_doc.ps1 unpacks it to %LOCALAPPDATA%\swapi_help\all)
#   cw_dir       = cworksapi.chm extracted with 7-Zip
import html
import os
import re
import sys

mshc, cw, out = sys.argv[1:4]
NSW = "SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks."
NCW = "SolidWorks.Interop.cosworks~SolidWorks.Interop.cosworks."
os.makedirs(out, exist_ok=True)


def clean(raw):
    raw = re.sub(r"(?is)<script.*?</script>", "", raw)
    raw = re.sub(r"(?is)<style.*?</style>", "", raw)
    raw = re.sub(r"(?i)<br\s*/?>|</p>|</tr>|</div>|</li>|</pre>", "\n", raw)
    raw = re.sub(r"<[^>]+>", " ", raw)
    raw = html.unescape(raw).replace("\xa0", " ").replace("﻿", "")
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in raw.split("\n")]
    return [ln for ln in lines if ln]


SKIP = re.compile(r"^(Language Filter|Collapse All|Send Feedback|SOLIDWORKS API Help$)")
DROP = ("Visual Basic (Declaration)", "Visual Basic (Usage)", "C#", "C++/CLI", "Visual Basic")
SECT = re.compile(r"^(Parameters|Return Value|Property Value|Remarks|Example|Examples|Requirements|See Also)$")


def page(path):
    """Condensed text of one member page, or None."""
    if not os.path.exists(path):
        return None
    lines = clean(open(path, encoding="utf-8", errors="ignore").read())
    lines = [ln for ln in lines if not SKIP.match(ln) and ln not in DROP]
    res = []
    i = 0
    while i < len(lines) and lines[i] != "Syntax":
        res.append(lines[i])
        i += 1
    i += 1
    decl = []
    while i < len(lines) and not lines[i].startswith("Dim instance"):
        decl.append(lines[i])
        i += 1
    res.append("```vb")
    res.extend(decl)
    res.append("```")
    while i < len(lines) and not SECT.match(lines[i]):
        i += 1
    while i < len(lines) and lines[i] not in ("Requirements", "See Also"):
        res.append(lines[i])
        i += 1
    return res


def cwpage(path):
    """Condensed text of one Simulation help page (older chm layout)."""
    if not os.path.exists(path):
        return None
    lines = clean(open(path, encoding="utf-8", errors="ignore").read())
    h = next((i for i, ln in enumerate(lines) if ln.startswith(": ") and "(I" in ln), None)
    if h is None:
        return None
    title = lines[h][2:].strip()
    j = next((i for i in range(h, len(lines)) if ".NET Syntax" in lines[i]), None)
    if j is None:
        return [title]
    summ = lines[j - 1] if j - 1 > h else ""
    txt = "\n".join(lines[j:])
    m = re.search(r"Visual Basic \(Declaration\)\s*(.*?)\s*Visual Basic \(Usage\)", txt, re.S)
    decl = m.group(1).strip() if m else ""
    decl = re.sub(r"\s+", " ", re.sub(r"\s*_\s*\n\s*", " ", decl))
    res = [title, summ, "```vb", decl, "```"]
    end = r"(?=\nVisual Basic for Applications|\nExample|\nRemarks|\nSee Also|\nAvailability|$)"
    m = re.search(r"\n(Parameters\b.*?)" + end, txt, re.S)
    if m:
        res.append(re.sub(r"\s+", " ", m.group(1)))
    m = re.search(r"\n((?:Return Value|Property Value)\b.*?)" + end, txt, re.S)
    if m:
        res.append(re.sub(r"\s+", " ", m.group(1)))
    m = re.search(r"Visual Basic for Applications \(VBA\) Syntax\s*(.*?)(?=\nExample|\nRemarks|\nSee Also|\nAvailability|$)", txt, re.S)
    if m:
        res.append("VBA: " + re.sub(r"\s+", " ", m.group(1)))
    m = re.search(r"\nRemarks\s*(.*?)(?=\nSee Also|\nAvailability|$)", txt, re.S)
    if m:
        res.append("Remarks: " + re.sub(r"\s+", " ", m.group(1)))
    ex = re.findall(r"([A-Z][A-Za-z\- ]+) \(VBA\)", txt)
    if ex:
        res.append("Examples: " + "; ".join(dict.fromkeys(e.strip() for e in ex)))
    return res


def members(base, ns, iface):
    pre = ns + iface + "~"
    names = sorted(f[len(pre):-5] for f in os.listdir(base) if f.startswith(pre) and f.endswith(".html"))
    return [n for n in names if not n.startswith("IGet") and not n.startswith("ISet")]


def doc_for(base, ns, iface, mems=None, cwmode=False):
    o = []
    lst = mems if mems is not None else members(base, ns, iface)
    for m in lst:
        fn = os.path.join(base, ns + iface + "~" + m + ".html")
        p = cwpage(fn) if cwmode else page(fn)
        if p is None:
            continue
        o.append("### %s.%s" % (iface, m))
        o.extend(p[1:] if cwmode else (p[1:] if len(p) > 1 and p[0].startswith(m) else p))
        o.append("")
    return o


def write(name, lines):
    with open(os.path.join(out, name), "w", encoding="ascii", errors="replace", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")


# ---------------------------------------------------------------- core
WL = ["ISldWorks", "IModelDoc2", "IModelDocExtension", "IAssemblyDoc", "IPartDoc", "IComponent2", "IBody2", "IFace2", "IEdge",
      "IFeature", "IFeatureManager", "ISketchManager", "ISketch", "IConfiguration", "IConfigurationManager", "IMathUtility",
      "IMathTransform", "IMathPoint", "ICustomPropertyManager", "IInterferenceDetectionMgr", "IInterference", "IEntity",
      "ISelectionMgr", "IRefPlane", "IExtrudeFeatureData2", "ICurve", "IVertex", "ISurface", "IMassProperty", "ITessellation",
      "ISketchSegment", "ISketchLine", "IModelView"]
CORE = """ClearSelection2 SetSuppression2 GetModelDoc2 SketchManager Save3 Extension SetPropertyByName2 Select2 Select4 GetCorrespondingEntity
GetBox EditRebuild3 Name2 GetTitle GetPathName GetNextFeature GetEntityName GetBodyBox GetBodies3 GetBodies2 FirstFeature GetDefinition Transform2
SaveAs IsSuppressed InsertSketch GetTypeName2 GetComponents GetArea AddToDB ArrayData ShowConfiguration2 ResolveAllLightWeightComponents
GetSuppression2 GetFaces ForceRebuild3 FeatureManager ConfigurationManager CloseDoc AccessSelections CustomPropertyManager
AddComponent5 ActiveSketch ActiveDoc ActiveConfiguration Activate Rebuild ModifyDefinition ModelToSketchTransform MultiplyTransform
CreateTransform CreatePoint CreatePlaneAtOffset3 GetMathUtility SetEntityName DeleteEntityName GetEdges GetCurve GetStartVertex GetChildren
GetOpenDocumentByName NewDocument OpenDoc6 ActivateDoc3 LoadAddIn GetAddInObject GetUserPreferenceStringValue GetImportFileData LoadFile4
FixComponent UnfixComponent GetInterferences TreatCoincidenceAsInterference TreatSubAssembliesAsComponents IncludeMultibodyPartInterferences
MakeInterferingPartsTransparent InterferenceDetectionManager Done GetMassProperties CreateMassProperty GetLastRebuildError AddConfiguration2
ActivateConfiguration DeleteSelection2 CreateLine CreateCornerRectangle FeatureExtrusion2 FeatureCut4 SetDepth GetDepth GetComponentAt
GetProcessID GetExecutablePath ViewZoomtofit2 EnumFeatures Get6 Add3 Normal Distance GetEntityByName ComponentCount Visible
GetTessTriangles GetTessTriangleCount Tessellate GetVertexCount GetVertexPoint NeedFaceFacetMap NeedVertexParams NeedVertexNormal GetTessellation
ReverseDirection IsHidden IsFixed GetSelectedObject6 GetSelectedObjectCount2 GetType GetPoint ModelType Hidden ComponentName""".split()

lines = ["# Core API reference (SolidWorks 2023, from the local help)", "",
         "Generated by build_reference.py from the unpacked sldworksapi help, grouped by interface.",
         "Only the calls this project uses or is likely to use. Look up anything else with api_doc.ps1.", ""]
found = set()
for iface in WL:
    names = members(mshc, NSW, iface)
    found.update(names)
    got = [m for m in CORE if m in names]
    if not got:
        continue
    lines.append("## " + iface)
    lines.append("")
    lines.extend(doc_for(mshc, NSW, iface, got))
write("core.md", lines)
print("core members not found in the whitelisted interfaces:", [m for m in CORE if m not in found])

# ---------------------------------------------------------------- simulation
CWI = ["ICosmosWorks", "ICWModelDoc", "ICWStudyManager", "ICWStudy", "ICWThermalStudyOptions", "ICWLoadsAndRestraintsManager",
       "ICWLoadsAndRestraints", "ICWHeatPower", "ICWHeatFlux", "ICWConvection", "ICWRadiation", "ICWTemperature", "ICWContactManager",
       "ICWContactSet", "ICWContactComponent", "ICWSolidManager", "ICWSolidBody", "ICWSolidComponent", "ICWMaterial", "ICWMesh",
       "ICWMeshControl", "ICWResults", "ICWPlot", "ICWResultsProbeManager", "ICWMassPropertiesManager",
       "ICWRunSpecStudiesRunMeshOptions", "ICWRunStudiesResults"]
lines = ["# Simulation (CosmosWorks) API reference, SolidWorks 2023", "",
         "Generated from cworksapi.chm. Thermal-relevant interfaces only. Entry: ICosmosWorks from the add-in",
         "(swApp.GetAddInObject(\"SldWorks.Simulation\")), then .ActiveDoc -> ICWModelDoc -> .StudyManager.", ""]
for iface in CWI:
    lines.append("## " + iface)
    lines.append("")
    lines.extend(doc_for(cw, NCW, iface, None, True))
write("simulation.md", lines)

EX = ["Apply_Thermostat-controlled_Heat_Power_for_Transient_Thermal_Study", "Get_Thermal_Values_at_Points", "Apply_Material_to_Bodies",
      "Apply_Mesh_Control", "Add_Component_Contacts", "Run_Studies_in_Batch_Mode", "Get_Mesh_Elements_and_Nodes",
      "Create_Plots_for_Static_Study", "Create_Results_Equation_Plot", "Duplicate_Study", "Copy_Mesh_and_Gen_Report",
      "Create_Curvature-based_Mesh"]
lines = ["# Simulation example code (VB6 from cworksapi.chm)", "",
         "Static-study examples are included where they show the call chain for materials, mesh, run and plots.",
         "The thermal calls (AddHeatPower, AddConvection, AddRadiation) are in simulation.md.", ""]
for e in EX:
    for suffix in ("_Example_VB.htm", "_Example_VBNET.htm"):
        p = os.path.join(cw, e + suffix)
        if not os.path.exists(p):
            continue
        raw = open(p, encoding="utf-8", errors="ignore").read()
        m = re.search(r"(?is)<pre[^>]*>(.*?)</pre>", raw)
        body = html.unescape(re.sub(r"<[^>]+>", "", m.group(1))) if m else "\n".join(clean(raw))
        lines.append("## %s (%s)" % (e.replace("_", " "), "VB6" if suffix == "_Example_VB.htm" else "VB.NET"))
        lines.append("```vb")
        lines.extend(body.replace("\xa0", " ").split("\n"))
        lines.append("```")
        lines.append("")
        break
write("simulation_examples.md", lines)
print("done")
