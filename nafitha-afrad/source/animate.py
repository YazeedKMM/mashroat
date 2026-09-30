"""Add a fade transition to every slide and an automatic, progressive fade-in
build for shapes named anim<step>_<n> by build.js."""
import re, sys, zipfile, shutil, os

src, dst = sys.argv[1], sys.argv[2]
EFFECT_MS = 450
GAP_MS = 250

def timing_xml(steps):
    ids = iter(range(3, 100000))
    t = 0
    step_pars = []
    for step in sorted(steps):
        spids = steps[step]
        effects = []
        for j, spid in enumerate(spids):
            node = 'afterEffect' if j == 0 else 'withEffect'
            effects.append(
                f'<p:par><p:cTn id="{next(ids)}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                f'<p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{next(ids)}" dur="{EFFECT_MS}"/>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
                f'</p:childTnLst></p:cTn></p:par>'
            )
        step_pars.append(
            f'<p:par><p:cTn id="{next(ids)}" fill="hold"><p:stCondLst><p:cond delay="{t}"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(effects)}</p:childTnLst></p:cTn></p:par>'
        )
        t += EFFECT_MS + GAP_MS
    outer = (
        f'<p:par><p:cTn id="{next(ids)}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
        f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst>'
        f'<p:childTnLst>{"".join(step_pars)}</p:childTnLst></p:cTn></p:par>'
    )
    bld = ''.join(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>' for s in sorted({x for v in steps.values() for x in v if x in SP_IDS}, key=int))
    return (
        '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        + outer +
        '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
        '</p:childTnLst></p:cTn></p:par></p:tnLst>'
        + (f'<p:bldLst>{bld}</p:bldLst>' if bld else '') +
        '</p:timing>'
    )

TRANSITION = '<p:transition spd="med"><p:fade/></p:transition>'

zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst + '.tmp', 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if re.fullmatch(r'ppt/slides/slide\d+\.xml', item.filename):
        xml = data.decode('utf-8')
        steps = {}
        SP_IDS = set()
        # shapes with text/fill (p:sp) get a bldP entry; pictures do not
        for m in re.finditer(r'<p:sp>\s*<p:nvSpPr>\s*<p:cNvPr id="(\d+)"', xml):
            SP_IDS.add(m.group(1))
        for m in re.finditer(r'<p:cNvPr id="(\d+)" name="anim(\d+)_\d+[^"]*"', xml):
            steps.setdefault(int(m.group(2)), []).append(m.group(1))
        extra = TRANSITION + (timing_xml(steps) if steps else '')
        if '</p:clrMapOvr>' in xml:
            xml = xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + extra, 1)
        else:
            xml = xml.replace('</p:cSld>', '</p:cSld>' + extra, 1)
        data = xml.encode('utf-8')
        print(item.filename, {k: len(v) for k, v in sorted(steps.items())})
    zout.writestr(item, data)
zout.close()
shutil.move(dst + '.tmp', dst)
