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
    from .design_checks import design_diagnostics
    from .mesh_checks import material_size_targets, measured_resolution, measure_walls
except ImportError:
    from schema import validate_design
    from design_checks import design_diagnostics
    from mesh_checks import material_size_targets, measured_resolution, measure_walls


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
    area = 0.5*np.abs((xy[:,1,0]-xy[:,0,0])*(xy[:,2,1]-xy[:,0,1])-(xy[:,1,1]-xy[:,0,1])*(xy[:,2,0]-xy[:,0,0]))
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
        wall_shapes=[]
        def disk(x,y,rx,ry=None):
            ry=rx if ry is None else ry
            if math.hypot(x,y)+max(rx,ry)>=R*.99:
                raise ValueError("geometry inclusion outside physical domain")
            # OCC requires the first radius to be the major axis. Preserve requested
            # x/y semiaxes by rotating a swapped ellipse instead of changing shape.
            surface=occ.addDisk(x,y,0,float(max(rx,ry)),float(min(rx,ry)))
            if ry>rx: occ.rotate([(2,surface)],x,y,0,0,0,1,math.pi/2)
            return surface
        def ring(x,y,r,t):
            if t<=0 or t>=r: raise ValueError("invalid ring thickness")
            outer=disk(x,y,r)
            inner=occ.addDisk(x,y,0,r-t,r-t)
            cut,_=occ.cut([(2,outer)],[(2,inner)],removeObject=True,removeTool=True)
            if len(cut)!=1: raise RuntimeError("ring OCC cut did not yield one surface")
            walls.append(float(t))
            wall_shapes.append({"centre_um":[float(x),float(y)],"outer_radius_um":float(r),
                                "thickness_um":float(t),"material":None,"surface":cut[0][1]})
            return cut[0][1]
        def add(surface,material):
            objects.append(((2,surface),material))
            for wall in wall_shapes:
                if wall["surface"]==surface: wall["material"]=material
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
                    add(disk(x,0,p["rod_radius_um"],p["rod_radius_um"]*.6 if kind=="bow_tie" else None),"stress_rod")
            notes.append("STEP6b: stress_rod is a distinct mechanical domain; isotropic photoelasticity requires sourced E, nu, alpha, p11, p12 inputs. Rod shapes remain simplified; core-area B is NOT modal neff splitting.")
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
        boundary_groups={}
        if external:
            boundary_groups["OUTER_"+b["type"].upper()]={"tag":20,"curves":sorted({tag for dim,tag in external if dim==1})}
        curve_groups={}
        for _,tag in gmsh.model.getEntities(1):
            adjacent,_=gmsh.model.getAdjacencies(1,tag)
            pair=sorted({labels[int(t)] for t in adjacent if int(t) in labels})
            if len(pair)>1:
                name="PML_INTERFACE" if "pml" in pair else "INTERFACE_"+"__".join(pair)
                curve_groups.setdefault(name,[]).append(tag)
        for j,(name,curves) in enumerate(sorted(curve_groups.items())):
            gid=gmsh.model.addPhysicalGroup(1,curves,30+j)
            gmsh.model.setPhysicalName(1,gid,name)
            boundary_groups[name]={"tag":gid,"curves":curves}
        targets=material_size_targets(d,materials)
        h_active=min(x["requested_edge_um"] for x in targets.values())
        policy=d["mesh"].get("resolution_policy","boundary_refined")
        wall_size=min(walls)/(d["mesh"]["through_wall_elements"]+1) if walls else h_active
        h_fine=min(h_active,wall_size)
        h_bulk=min(d["mesh"]["bulk_max_um"],.8*h_active) if policy=="wavelength" else float(d["mesh"]["bulk_max_um"])
        gmsh.option.setNumber("Mesh.ElementOrder",1)
        gmsh.option.setNumber("Mesh.MeshSizeFromCurvature",int(d["mesh"]["curvature_segments"]))
        gmsh.option.setNumber("Mesh.MinimumCircleNodes",int(d["mesh"]["curvature_segments"]))
        gmsh.option.setNumber("Mesh.MeshSizeExtendFromBoundary",0)
        gmsh.option.setNumber("Mesh.MeshSizeFromPoints",0)
        gmsh.option.setNumber("Mesh.MeshSizeMin",0)
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
        material_ids=np.asarray(ids,dtype="U40")
        boundary_edges={}
        for name,group in boundary_groups.items():
            edges=[]
            for curve in group["curves"]:
                types,_,nodes=gmsh.model.mesh.getElements(1,curve)
                for etype,nn in zip(types,nodes):
                    if int(etype)!=1: raise RuntimeError("only linear boundary edges supported")
                    edges.extend([[node_to_idx[int(t)] for t in row] for row in np.asarray(nn).reshape(-1,2)])
            boundary_edges[name]=np.asarray(edges,dtype=np.int32)
        descriptors=[{k:v for k,v in w.items() if k!="surface"} for w in wall_shapes]
        wall_report=measure_walls(xy,tri,material_ids,descriptors,d["mesh"]["through_wall_elements"],
                                  int(d["mesh"].get("wall_probe_angles",24)))
        resolution=measured_resolution(xy,tri,material_ids,targets)
        e1,e2=xy[tri[:,1]]-xy[tri[:,0]],xy[tri[:,2]]-xy[tri[:,0]]
        areas=.5*np.abs(e1[:,0]*e2[:,1]-e1[:,1]*e2[:,0])
        material_areas={mat:float(areas[material_ids==mat].sum()) for mat in materials}
        cad_areas={mat:sum(gmsh.model.occ.getMass(2,tag) for tag,label in labels.items() if label==mat) for mat in materials}
        properties=mesh_quality(xy,tri)
        notes.extend(design_diagnostics(d)["approximations"])
        if policy=="boundary_refined": notes.append("Compatibility policy: bulk wavelength resolution is measured, NOT enforced; select wavelength for whole-domain limit")
        if wall_report["status"]=="FAIL_SAMPLED": notes.append("FAIL: sampled circular wall resolution; refine mesh before optics")
        properties.update({"engine_version":"0.8.4","resolution_policy":policy,
                           "materials":{m:int(np.count_nonzero(material_ids==m)) for m in materials},
                           "material_area_um2":material_areas,"cad_material_area_um2":cad_areas,
                           "domain_labels":{str(k):v for k,v in labels.items()},
                           "physical_groups":groups,"boundary_groups":boundary_groups,"boundary_type":b["type"],
                           "wall_min_um":min(walls) if walls else None,
                           "target_elements_per_wall":d["mesh"]["through_wall_elements"],
                           "target_fine_edge_um":h_fine,"target_bulk_edge_um":h_bulk,
                           "wavelength_size_safety_factor":.8 if policy=="wavelength" else None,
                           "curvature_elements_per_2pi":int(d["mesh"]["curvature_segments"]),
                           "wall_resolution":wall_report,"wavelength_resolution":resolution,
                           "design_diagnostics":design_diagnostics(d),
                           "notes":notes,"validation_status":"geometry/mesh diagnostics only; optical convergence not established"})
        if output_dir:
            dst=Path(output_dir);dst.mkdir(parents=True,exist_ok=True)
            np.savez_compressed(dst/"mesh.npz",points_um=xy,triangles=tri,material_ids=material_ids,
                                **{"boundary_"+k:v for k,v in boundary_edges.items()})
            (dst/"mesh_report.json").write_text(json.dumps(properties,indent=2,ensure_ascii=False,allow_nan=False),encoding="utf-8")
            gmsh.write(str(dst/"geometry.msh"))
            selected=np.linspace(0,len(tri)-1,min(6000,len(tri)),dtype=int)
            preview={"schema":"LSCNS_mesh_preview_v084","engine_version":"0.8.4","input":d,
                     "quality":properties,"triangles":[{"vertices_um":xy[tri[i]].tolist(),"material":str(material_ids[i])}
                                                        for i in selected],
                     "boundary_edges":{k:xy[v].tolist() for k,v in boundary_edges.items()},
                     "sampling":{"shown_triangles":len(selected),"total_triangles":len(tri),
                                 "basis":"deterministic triangle subset; all labelled boundary edges exported"}}
            (dst/"mesh_preview.json").write_text(json.dumps(preview,separators=(",",":"),allow_nan=False)+"\n")
        strict=d["mesh"].get("strict_checks",policy=="wavelength")
        if strict and (wall_report["status"]=="FAIL_SAMPLED" or resolution["status"]=="FAIL_POLICY"):
            raise RuntimeError("mesh resolution checks failed; inspect mesh_report.json and refine")
        return {"points_um":xy,"triangles":tri,"material_ids":material_ids,
                "boundary_edges":boundary_edges,"quality":properties}
    finally:
        gmsh.finalize()

