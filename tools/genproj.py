"""Generate ParamTable.vcxproj from ParamTable.vcxproj.in.

The include list is every folder under the Mastercam SDK's interfaces\\, read off
disk rather than typed. A missing one is a header "not found" that names nothing
useful, and the SDK's headers include each other by bare filename.

The SDK targets the v143 toolset. Visual Studio 2026's own v145 fails on it - the
BCG header rejects _MSC_VER >= 1950, and UniqueArray_CH.h puts `const` on a
constructor, a hard error C7731 in every language mode - so build.bat selects v143.

Usage:  python tools\\genproj.py [path-to-sdk]
"""
import io
import os
import sys

BS = chr(92)
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

sdk = sys.argv[1] if len(sys.argv) > 1 else (
    'C:' + BS + 'Program Files' + BS + 'Mastercam 2026' + BS + 'sdk')

inter = sorted(d for d in os.listdir(os.path.join(sdk, 'interfaces'))
               if os.path.isdir(os.path.join(sdk, 'interfaces', d)))

extra = []
for rel in ('NGP' + BS + 'Core' + BS + 'h',
            'NGP' + BS + 'Core' + BS + 'gaf' + BS + 'base' + BS + 'misc' + BS + 'h',
            '3rdParty' + BS + 'BCGCBPro' + BS + 'src'):
    ok = os.path.isdir(os.path.join(sdk, rel))
    print(('OK   ' if ok else 'MISS ') + rel)
    if ok:
        extra.append(rel)

inc = ['$(SDKDir)']
inc += ['$(SDKDir)interfaces']                # the SDK includes as "BuildTools/targetver_CH.h"
inc += ['$(SDKDir)' + r for r in extra]
inc += ['$(SDKDir)interfaces' + BS + d for d in inter]

# Second-level folders. Most of the SDK's headers find each other through their own
# directory, but the CHookAPI ones include their neighbours by bare name from a sibling
# folder (cdbIArcGeometry_CH.h wants cmaPoint3D_CH.h, which is in CHookAPI\\Math), and
# the sister project needed the same set. Every subfolder of CHookAPI, plus the job
# setup service folders.
second = []
chook = os.path.join(sdk, 'interfaces', 'CHookAPI')
for d in sorted(os.listdir(chook)):
    if os.path.isdir(os.path.join(chook, d)):
        second.append('CHookAPI' + BS + d)
for rel in ('JobSetup' + BS + 'Services', 'JobSetupCore' + BS + 'Data',
            'JobSetupCore' + BS + 'Services'):
    if os.path.isdir(os.path.join(sdk, 'interfaces', rel)):
        second.append(rel)
inc += ['$(SDKDir)interfaces' + BS + r for r in second]
print('%d second-level folders' % len(second))

tmpl = io.open(os.path.join(ROOT, 'ParamTable.vcxproj.in'), encoding='utf-8').read()
proj = tmpl.replace('__INCLUDES__', ';'.join(inc))
io.open(os.path.join(ROOT, 'ParamTable.vcxproj'), 'w', encoding='utf-8',
        newline='').write(proj)
print('%d interface folders; wrote ParamTable.vcxproj (%d bytes)' % (len(inter), len(proj)))
