"""Benchmark API facade. Results and archived assets remain at ../../benchmarks."""
def validate(step3_report,out,run_photonics=True):
 from validate_step7_v089 import build_report
 return build_report(step3_report,out,run_photonics)
