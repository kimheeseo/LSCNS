"""Labeled mesh / quality / convergence facade."""
def generate(design,output_dir=None):
 from geometry import build_mesh
 return build_mesh(design,output_dir)
def quality(points,triangles):
 from geometry import mesh_quality
 return mesh_quality(points,triangles)
def study(*args,**kwargs):
 from modal_study import run_cases
 return run_cases(*args,**kwargs)
