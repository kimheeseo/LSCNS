"""Postprocess facade; computation and browser presentation stay separate."""
from .formulas import loss,area,gamma,basis,REGISTRY
def metrics(result,context=None):
 from step4_v086 import postprocess
 return postprocess(result,context)
def track_sweep(result):
 from step4_v086 import track_sweep as existing
 return existing(result)
