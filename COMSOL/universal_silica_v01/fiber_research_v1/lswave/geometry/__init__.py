"""Parametric cross-section facade; existing geometry implementation preserved."""
def build(design,output_dir=None):
 from geometry import build_mesh
 return build_mesh(design,output_dir)
