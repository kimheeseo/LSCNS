"""Blender 4.x: 유기적인 동물 메시를 만들고 단일 HTML에 내장합니다.
실행: blender -b --python tools/build_marine_models.py -- /absolute/path/index.html
외부 모델·텍스처·플러그인을 사용하지 않습니다.
"""
import bpy, bmesh, math, json, sys, re
from pathlib import Path
bpy.ops.object.select_all(action='SELECT')
bpy.ops.object.delete(use_global=False)

def finish(name, vertices, faces, subdivisions=1):
    data=bpy.data.meshes.new(name)
    data.from_pydata(vertices, [], faces)
    data.update()
    bm=bmesh.new();bm.from_mesh(data)
    bmesh.ops.recalc_face_normals(bm, faces=list(bm.faces))
    bm.to_mesh(data);bm.free()
    ob=bpy.data.objects.new(name,data);bpy.context.collection.objects.link(ob)
    bpy.context.view_layer.objects.active=ob;ob.select_set(True)
    if subdivisions:
        mod=ob.modifiers.new('유기 곡면','SUBSURF');mod.levels=subdivisions
        bpy.ops.object.modifier_apply(modifier=mod.name)
    for poly in ob.data.polygons:poly.use_smooth=True
    ob.data.calc_loop_triangles()
    out={'p':[], 'n':[], 'uv':[], 'i':[]}
    for v in ob.data.vertices:
        x,y,z=v.co;out['p'] += [round(x,5),round(y,5),round(z,5)]
        out['n'] += [round(k,5) for k in v.normal]
        out['uv'] += [round((x+1.2)/2.4,5),round((z+.85)/1.7,5)]
    for tri in ob.data.loop_triangles:out['i'] += list(tri.vertices)
    ob.select_set(False)
    return out

def loft(name, profile, segments=24):
    vertices=[];faces=[]
    for x,ry,rz,cy in profile:
        for j in range(segments):
            t=2*math.pi*j/segments
            vertices.append((x,cy+ry*math.cos(t),rz*math.sin(t)))
    for k in range(len(profile)-1):
        for j in range(segments):
            a=k*segments+j;b=k*segments+(j+1)%segments
            faces.append((a,b,b+segments,a+segments))
    faces.append(tuple(reversed(range(segments))))
    faces.append(tuple((len(profile)-1)*segments+j for j in range(segments)))
    return finish(name,vertices,faces)

def fin(name, outline, thickness):
    # 렌즈 단면으로 두께를 만들어 평면 삼각형 인상을 제거합니다.
    n=len(outline);cx=sum(p[0] for p in outline)/n;cy=sum(p[1] for p in outline)/n
    vertices=[(x,y,0) for x,y in outline]+[(cx,cy,thickness),(cx,cy,-thickness)]
    faces=[]
    for j in range(n):faces.extend([(j,(j+1)%n,n),((j+1)%n,j,n+1)])
    return finish(name,vertices,faces,2)

models={
 'sharkBody':loft('상어 몸체',[
 (-2.05,.07,.07,0),(-1.65,.13,.15,-.015),(-1.15,.32,.31,0),
 (-.6,.49,.45,.015),(0,.55,.5,.035),(.65,.49,.43,.025),
 (1.15,.35,.32,0),(1.65,.2,.23,-.045),(1.95,.08,.13,-.075),(2.04,.015,.025,-.08)]),
 'dorsalFin':fin('상어 등지느러미',[(-.85,.32),(-.52,.63),(-.22,1.35),(.05,1.24),(.18,.7),(.62,.38)],.09),
 'pectoralFin':fin('상어 가슴지느러미',[(-.68,0),(-1.15,-1.4),(-.79,-1.32),(.48,-.06),(.35,.18)],.075),
 'sharkTail':fin('상어 꼬리지느러미',[(.2,0),(-.25,.28),(-.98,1.05),(-1.1,1.1),(-.79,.24),(-.55,0),(-1.02,-.72),(-.92,-.8),(-.18,-.17)],.07),
 'turtleFlipper':fin('거북이 물갈퀴',[(-.34,.05),(-.42,-.3),(-.13,-1.42),(.08,-1.58),(.28,-.98),(.42,-.35),(.32,.07)],.07)
}
# 편평한 배면과 둥근 등면을 가진 등껍질입니다.
verts=[];faces=[];rows=12;cols=32
for k in range(rows+1):
    p=.001+(math.pi-.002)*k/rows
    for j in range(cols):
        a=j*math.pi*2/cols;y=.45*math.cos(p)
        if y<0:y*=.4
        verts.append((1.08*math.sin(p)*math.cos(a),y,.79*math.sin(p)*math.sin(a)))
for k in range(rows):
    for j in range(cols):
        a=k*cols+j;b=k*cols+(j+1)%cols;faces.append((a,a+cols,b+cols,b))
faces.append(tuple(reversed(range(cols))))
faces.append(tuple(rows*cols+j for j in range(cols)))
models['turtleShell']=finish('거북이 등껍질',verts,faces)
# JSON은 미리 삼각분할한 위치·법선·UV·인덱스를 포함합니다.
payload=json.dumps(models,separators=(',',':'),ensure_ascii=True)
args=sys.argv[sys.argv.index('--')+1:] if '--' in sys.argv else []
if args:
    target=Path(args[0]);source=target.read_text()
    begin='/* BLENDER_MARINE_BEGIN */';end='/* BLENDER_MARINE_END */'
    block=begin+'\nconst marineModels='+payload+';\n'+end
    if begin in source:source=re.sub(re.escape(begin)+r'[\s\S]*?'+re.escape(end),lambda m:block,source)
    else:source=source.replace('function createShark(){',block+'\nfunction createShark(){',1)
    target.write_text(source)
print('MARINE_MODELS',len(payload),'bytes',sum(len(m['p'])//3 for m in models.values()),'vertices')
