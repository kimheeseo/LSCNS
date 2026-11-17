"""Parametric planar gmsh OCC geometry, physical domains, circular PML and mesh diagnostics.

Research prototype. Thin-wall resolution target is a meshing request, not a convergence certificate.
All geometry lengths in micrometres.
"""
from pathlib import Path
import json
import math
import numpy as np

try:
    from .schema import validate_design
except ImportError:
    from schema import validate_design


def hex_centers(rings, pitch, omit=0):
    result = []
    for q in range(-rings, rings + 1):
        for r in range(-rings, rings + 1):
            ring = max(abs(q),abs(r),abs(q+r))
            if ring <= rings and ring > omit:
                result.append((pitch * (q + r/2), pitch * math.sqrt(3)*r/2))
    return result


def mcf_centers(n, pitch, kind):
    if kind == "ring":
        return [(pitch*math.cos(2*math.pi*i/n),pitch*math.sin(2*math.pi*i/n)) for i in range(n)]
    centers = [(0.,0.)]
    for r in range(1,4):
        ring = [(pitch*(q+s/2),pitch*math.sqrt(3)*s/2)
                for q in range(-r,r+1) for s in range(-r,r+1)
                if max(abs(q),abs(s),abs(q+s))==r]
        centers.extend(sorted(ring,key=lambda x:math.atan2(x[1],x[0])))
        if len(centers) >= n: break
    return centers[:n]


def mesh_quality(points, tris):
    xy = points[tris]
    ab = np.linalg.norm(xy[:,1]-xy[:,0],axis=1)
    bc = np.linalg.norm(xy[:,2]-xy[:,1],axis=1)
    ca = np.linalg.norm(xy[:,0]-xy[:,2],axis=1)
    area = 0.5*np.abs(np.cross(xy[:,1]-xy[:,0],xy[:,2]-xy[:,0]))
    if len(area)==0 or not np.all(np.isfinite(area)) or np.min(area)<=0:
        raise RuntimeError("invalid mesh: nonpositive/NaN triangular area")
    edges=np.column_stack((ab,bc,ca))
    hmax=edges.max(axis=1)
    aspect=hmax*hmax/(2*area)  # longest edge / shortest altitude, >= 2/sqrt(3)
    return {"triangles":int(len(tris)),"vertices":int(len(points)),
            "edge_min_um":float(edges.min()),"edge_max_um":float(edges.max()),
            "area_min_um2":float(area.min()),"area_max_um2":float(area.max()),
            "aspect_ratio_max":float(aspect.max()),"aspect_ratio_p99":float(np.percentile(aspect,99))}


def build_mesh(design, output_dir=None):
    """Return dict arrays p: (N,2), t:(M,3), material_ids and geometry/mesh metrics.

    Gmsh is optional for config validation but required here.
    """
    validate_design(design)
    import gmsh
    d=design;p=d["parameters"];fam=d["family"];kind=d["variant"]
    R=float(d["domain_radius_um"]); b=d["boundary"]
    lam_um=d["wavelength_nm"]*1e-3
    gmsh.initialize()
    try:
        gmsh.option.setNumber("General.Terminal",0)
        gmsh.model.add("LSCNS_fiber_v1")
        occ=gmsh.model.occ
        objects=[]
        notes=[]
        walls=[]
        fine_curves=[]
        def disk(x,y,rx,ry=None):
            ry=rx if ry is None else ry
            if math.hypot(x,y)+max(rx,ry)>=R*.99:
                raise ValueError("geometry inclusion outside physical domain")
            return occ.addDisk(x,y,0,float(rx),float(ry))
        def ring(x,y,r,t):
            if t<=0 or t>=r: raise ValueError("invalid ring thickness")
            outer=disk(x,y,r)
            inner=occ.addDisk(x,y,0,r-t,r-t)
            cut,_=occ.cut([(2,outer)],[(2,inner)],removeObject=True,removeTool=True)
            if len(cut)!=1: raise RuntimeError("ring OCC cut did not yield one surface")
            walls.append(float(t))
            return cut[0][1]
        def add(surface,material):
            objects.append(((2,surface),material))
        base=occ.addDisk(0,0,0,R,R)
        add(base,"background")
        if b["type"]=="pml":
            outer=occ.addDisk(0,0,0,R+b["thickness_um"],R+b["thickness_um"])
            inner=occ.addDisk(0,0,0,R,R)
            pieces,_=occ.cut([(2,outer)],[(2,inner)],removeObject=True,removeTool=True)
            add(pieces[0][1],"pml")
        if fam=="index_guiding":
            a=p["core_radius_um"]
            add(disk(0,0,a),"core")
            if kind in ("trench_assisted","w_profile"):
                start=p.get("trench_inner_um",a*1.3)
                end=p.get("trench_outer_um",a*1.8)
                if not a < start < end < R: raise ValueError("invalid trench radii")
                add(ring(0,0,end,end-start),"trench")
            if kind in ("graded","multilayer"):
                notes.append("Graded / multilayer approximate piecewise radial rings; coefficients are constant within each ring")
                count=int(p.get("layer_count",4) if kind=="multilayer" else 8)
                for i in range(count):
                    r=a*(1-i/(count+1))
                    add(disk(0,0,r),"core" if i%2==0 else "silica")
        elif fam in ("solid_pcf","pbgf"):
            pitch=p["pitch_um"];ratio=p["d_over_pitch"]
            omit=(2 if p.get("defect_cells",1)==19 else 1 if p.get("defect_cells",1)==7 else 0)
            for x,y in hex_centers(int(p["rings"]),pitch,omit=omit) if kind!="square" else [
                (i*pitch,j*pitch) for i in range(-int(p["rings"]),int(p["rings"])+1)
                for j in range(-int(p["rings"]),int(p["rings"])+1) if (i or j)]:
                rr=pitch*ratio/2
                if math.hypot(x,y)+rr<R*.94:
                    add(disk(x,y,rr,rr*p.get("ellipse_ratio",1)),"air")
            if fam=="pbgf":
                defect=(1 if omit==1 else 2)*pitch*.87
                add(disk(0,0,defect),"air")
                if p["core_surround_um"]>0:
                    add(ring(0,0,defect+p["core_surround_um"],p["core_surround_um"]),"silica")
                notes.append("PBGF finite cross-section only; Bloch bandgap map NOT computed")
            else:
                notes.append("Endlessly single-mode guide is conceptual only; finite rings do not prove single-mode")
        elif fam in ("anti_resonant","kagome"):
            N=int(p.get("tube_count",6) if fam=="anti_resonant" else 8)
            core=p["core_radius_um"]
            tube=float(p.get("tube_radius_um",p.get("pitch_um",3)*1.8))
            wall=p["wall_thickness_um"]
            distance=core+tube+float(p.get("gap_um",.6))
            for i in range(N):
                ang=2*math.pi*i/N
                x,y=distance*math.cos(ang),distance*math.sin(ang)
                add(ring(x,y,tube,wall),"silica")
                if kind in ("nanf","dnanf"):
                    r2=tube*p.get("nested_radius_ratio",.48)
                    add(ring(x,y,r2,wall*.7),"silica")
                    if kind=="dnanf":
                        add(ring(x,y,r2*.48,wall*.5),"silica")
            notes.append("Simplified circular-tube layout, not exact Kagome / conjoined / hypocycloid / negative curvature CAD")
        elif fam=="bragg":
            a=p["core_radius_um"]
            add(disk(0,0,a),"air")
            for i in range(int(p["layer_count"])):
                r=a+(i+1)*p["layer_pitch_um"]
                add(ring(0,0,r,p["layer_pitch_um"]),"high" if i%2==0 else "low")
        elif fam=="mcf":
            centers=mcf_centers(int(p["core_count"]),p["pitch_um"],kind)
            for x,y in centers:
                if kind=="trench_assisted":
                    add(ring(x,y,p["core_radius_um"]+p["trench_width_um"],p["trench_width_um"]),"trench")
                add(disk(x,y,p["core_radius_um"]),"core")
        elif fam=="polarization_maintaining":
            a=p["core_radius_um"]
            add(disk(0,0,a,a*p.get("ellipse_ratio",1) if kind=="elliptic_core" else a),"core")
            if kind!="elliptic_core":
                for x in (-p["rod_offset_um"],p["rod_offset_um"]):
                    add(disk(x,0,p["rod_radius_um"],p["rod_radius_um"]*.6 if kind=="bow_tie" else None),"silica")
            notes.append("Stress rods are geometry only: stress-optic anisotropy and birefringence NOT solved")
        occ.synchronize()
        input_shapes=[x[0] for x in objects]
        _,mapping=occ.fragment([input_shapes[0]],input_shapes[1:],removeObject=True,removeTool=True)
        occ.synchronize()
        labels={}
        for source,(shape,material) in enumerate(objects):
            for dim,tag in mapping[source]:
                if dim==2: labels[tag]=material
        for dim,tag in gmsh.model.getEntities(2):
            if tag not in labels: raise RuntimeError(f"unlabelled domain surface {tag}")
        materials=sorted(set(labels.values()))
        groups={}
        for j,mat in enumerate(materials):
            tags=[tag for tag,material in labels.items() if material==mat]
            idx=gmsh.model.addPhysicalGroup(2,tags,100+j)
            gmsh.model.setPhysicalName(2,idx,mat)
            groups[mat]=idx
        external=gmsh.model.getBoundary(gmsh.model.getEntities(2),combined=True,oriented=False)
        if external:
            idx=gmsh.model.addPhysicalGroup(1,sorted({tag for dim,tag in external if dim==1}),20)
            gmsh.model.setPhysicalName(1,idx,"OUTER_"+b["type"].upper())
        gmsh.option.setNumber("Mesh.ElementOrder",1)
        gmsh.option.setNumber("Mesh.MeshSizeFromCurvature",0)
        gmsh.option.setNumber("Mesh.MeshSizeExtendFromBoundary",0)
        gmsh.option.setNumber("Mesh.MeshSizeFromPoints",0)
        nmax=max(1.00027,1.45)
        h_active=min(d["mesh"]["bulk_max_um"],lam_um/(nmax*d["mesh"]["elements_per_wavelength"]))
        wall_size=min(walls)/d["mesh"]["through_wall_elements"] if walls else h_active
        h_fine=min(h_active,wall_size)
        h_bulk=float(d["mesh"]["bulk_max_um"])
        gmsh.option.setNumber("Mesh.MeshSizeMin",max(h_fine*.75,.001))
        gmsh.option.setNumber("Mesh.MeshSizeMax",h_bulk)
        # Use localized boundary refinement; do not force sub-micrometre global mesh.
        boundary_curves=sorted({tag for dim,tag in gmsh.model.getBoundary(gmsh.model.getEntities(2),combined=False,oriented=False) if dim==1})
        if boundary_curves:
            distance=gmsh.model.mesh.field.add("Distance")
            gmsh.model.mesh.field.setNumbers(distance,"CurvesList",boundary_curves)
            gmsh.model.mesh.field.setNumber(distance,"Sampling",int(d["mesh"]["curvature_segments"]))
            threshold=gmsh.model.mesh.field.add("Threshold")
            gmsh.model.mesh.field.setNumber(threshold,"IField",distance)
            gmsh.model.mesh.field.setNumber(threshold,"SizeMin",h_fine)
            gmsh.model.mesh.field.setNumber(threshold,"SizeMax",h_bulk)
            gmsh.model.mesh.field.setNumber(threshold,"DistMin",min(walls)*.5 if walls else h_active)
            gmsh.model.mesh.field.setNumber(threshold,"DistMax",max(2*lam_um,3*h_bulk))
            gmsh.model.mesh.field.setAsBackgroundMesh(threshold)
        gmsh.model.mesh.generate(2)
        node_tags,xyz,_=gmsh.model.mesh.getNodes()
        xy=np.asarray(xyz,dtype=float).reshape(-1,3)[:,:2]
        node_to_idx={int(t):i for i,t in enumerate(node_tags)}
        triangles=[];ids=[]
        for dim,tag in gmsh.model.getEntities(2):
            types,_,element_nodes=gmsh.model.mesh.getElements(2,tag)
            for etype,nodes in zip(types,element_nodes):
                if int(etype)!=2: raise RuntimeError("only linear 3-node triangles are accepted")
                tri=np.asarray(nodes,dtype=np.int64).reshape(-1,3)
                triangles.extend([[node_to_idx[int(t)] for t in row] for row in tri])
                ids.extend([labels[tag]]*len(tri))
        tri=np.asarray(triangles,dtype=np.int32)
        if len(set(map(tuple,sorted(map(tuple,np.sort(tri,axis=1)))))) != len(tri):
            raise RuntimeError("duplicate mesh elements")
        properties=mesh_quality(xy,tri)
        properties.update({"materials":{m:int(np.count_nonzero(np.array(ids)==m)) for m in materials},
                           "domain_labels":{str(k):v for k,v in labels.items()},
                           "physical_groups":groups,"boundary_type":b["type"],
                           "wall_min_um":min(walls) if walls else None,
                           "target_elements_per_wall":d["mesh"]["through_wall_elements"],
                           "target_fine_edge_um":h_fine,"target_bulk_edge_um":h_bulk,
                           "notes":notes,"validation_status":"mesh metrics only; convergence not established"})
        if output_dir:
            dst=Path(output_dir);dst.mkdir(parents=True,exist_ok=True)
            np.savez_compressed(dst/"mesh.npz",points_um=xy,triangles=tri,material_ids=np.asarray(ids,dtype="U40"))
            (dst/"mesh_report.json").write_text(json.dumps(properties,indent=2,ensure_ascii=False,allow_nan=False),encoding="utf-8")
            gmsh.write(str(dst/"geometry.msh"))
        return {"points_um":xy,"triangles":tri,"material_ids":np.asarray(ids,dtype="U40"),"quality":properties}
    finally:
        gmsh.finalize()
