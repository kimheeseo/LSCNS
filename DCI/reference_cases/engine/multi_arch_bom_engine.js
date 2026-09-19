/*
 * Multi-Architecture Data Center BOM Engine
 *
 * Shared calculation core for all public reference cases.
 * IMPORTANT: this file contains no reference/expected answers.
 * Every output is derived only from design_input.json.
 */

function product(xs) {
  return (xs || []).reduce((a, b) => a * Number(b), 1);
}

function ceilDiv(a, b) {
  if (!Number.isFinite(a) || !Number.isFinite(b) || b <= 0) return null;
  return Math.ceil(a / b);
}

function requirePositive(name, value) {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) throw new Error(name + " must be > 0");
  return n;
}

function deriveCommon(input) {
  const perHost = requirePositive("accelerator.per_host", input.accelerator && input.accelerator.per_host);
  const targetHostsInput = input.target_hosts != null ? requirePositive("target_hosts", input.target_hosts) : null;
  const targetAcceleratorsInput = input.target_accelerators != null ? requirePositive("target_accelerators", input.target_accelerators) : null;
  if (targetHostsInput == null && targetAcceleratorsInput == null) throw new Error("target_hosts or target_accelerators is required");
  const targetAccelerators = targetAcceleratorsInput != null ? targetAcceleratorsInput : targetHostsInput * perHost;

  let acceleratorsPerBlock = null;
  if (input.building_block && Array.isArray(input.building_block.dimensions)) {
    acceleratorsPerBlock = product(input.building_block.dimensions);
  } else if (input.building_block && input.building_block.accelerators_per_block) {
    acceleratorsPerBlock = requirePositive(
      "building_block.accelerators_per_block",
      input.building_block.accelerators_per_block
    );
  }

  const blocksPerRack = input.rack && input.rack.blocks_per_rack != null
    ? Number(input.rack.blocks_per_rack) : null;
  const acceleratorsPerRack = input.rack && input.rack.accelerators_per_rack
    ? requirePositive("rack.accelerators_per_rack", input.rack.accelerators_per_rack)
    : (acceleratorsPerBlock && blocksPerRack != null)
      ? acceleratorsPerBlock * blocksPerRack
      : null;

  const hostCount = targetHostsInput != null ? targetHostsInput : targetAccelerators / perHost;
  const buildingBlockCount = acceleratorsPerBlock ? ceilDiv(targetAccelerators, acceleratorsPerBlock) : null;
  const hostsPerBlock = acceleratorsPerBlock ? acceleratorsPerBlock / perHost : null;
  const rackCount = acceleratorsPerRack ? ceilDiv(targetAccelerators, acceleratorsPerRack) : null;
  const hostsPerRack = acceleratorsPerRack ? acceleratorsPerRack / perHost : null;

  const dcnPerChip = input.accelerator && input.accelerator.dcn_bandwidth_per_chip_gbps != null
    ? Number(input.accelerator.dcn_bandwidth_per_chip_gbps) : null;
  const dcnPerHost = dcnPerChip != null ? dcnPerChip * perHost : null;

  const nicsPerHost = input.network && input.network.nics_per_host != null
    ? Number(input.network.nics_per_host) : null;
  const nicSpeedGbps = input.network && input.network.nic_speed_gbps != null
    ? Number(input.network.nic_speed_gbps) : null;
  const totalNicCount = nicsPerHost != null ? hostCount * nicsPerHost : null;
  const nicAggregatePerHostGbps = nicsPerHost != null && nicSpeedGbps != null
    ? nicsPerHost * nicSpeedGbps : null;
  const dcnAggregatePodTbps = dcnPerChip != null
    ? targetAccelerators * dcnPerChip / 1000 : (nicAggregatePerHostGbps != null ? hostCount * nicAggregatePerHostGbps / 1000 : null);

  const iciPortsPerChip = input.accelerator && input.accelerator.ici_ports_per_chip != null
    ? Number(input.accelerator.ici_ports_per_chip) : null;
  const iciPortsTotal = iciPortsPerChip != null ? targetAccelerators * iciPortsPerChip : null;

  const peakBf16TflopsPerChip = input.accelerator && input.accelerator.peak_bf16_tflops_per_chip != null
    ? Number(input.accelerator.peak_bf16_tflops_per_chip) : null;
  const peakBf16PflopsPerPod = peakBf16TflopsPerChip != null
    ? peakBf16TflopsPerChip * targetAccelerators / 1000 : null;

  const peakFp8TflopsPerChip = input.accelerator && input.accelerator.peak_fp8_tflops_per_chip != null
    ? Number(input.accelerator.peak_fp8_tflops_per_chip) : null;
  const peakFp8PflopsPerPod = peakFp8TflopsPerChip != null
    ? peakFp8TflopsPerChip * targetAccelerators / 1000 : null;

  const hbmCapacityGibPerChip = input.accelerator && input.accelerator.hbm_capacity_gib_per_chip != null
    ? Number(input.accelerator.hbm_capacity_gib_per_chip) : null;
  const hbmCapacityGibPerPod = hbmCapacityGibPerChip != null
    ? hbmCapacityGibPerChip * targetAccelerators : null;

  const hbmBandwidthGBpsPerChip = input.accelerator && input.accelerator.hbm_bandwidth_GBps_per_chip != null
    ? Number(input.accelerator.hbm_bandwidth_GBps_per_chip) : null;
  const hbmBandwidthTBpsPerPod = hbmBandwidthGBpsPerChip != null
    ? hbmBandwidthGBpsPerChip * targetAccelerators / 1000 : null;

  const tensorCoresPerChip = input.accelerator && input.accelerator.tensor_cores_per_chip != null
    ? Number(input.accelerator.tensor_cores_per_chip) : null;
  const tensorCoresTotal = tensorCoresPerChip != null ? tensorCoresPerChip * targetAccelerators : null;

  const sparseCoresPerChip = input.accelerator && input.accelerator.sparse_cores_per_chip != null
    ? Number(input.accelerator.sparse_cores_per_chip) : null;
  const sparseCoresTotal = sparseCoresPerChip != null ? sparseCoresPerChip * targetAccelerators : null;

  let maxSlice = {};
  if (input.validation_slice && Array.isArray(input.validation_slice.dimensions)) {
    const sliceAccelerators = product(input.validation_slice.dimensions);
    maxSlice = {
      max_slice_accelerator_count: sliceAccelerators,
      max_slice_host_count: sliceAccelerators / perHost,
      max_slice_block_count: acceleratorsPerBlock ? sliceAccelerators / acceleratorsPerBlock : null
    };
  }

  return {
    accelerator_count: targetAccelerators,
    accelerators_per_host: perHost,
    cpu_host_count: hostCount,
    accelerators_per_block: acceleratorsPerBlock,
    building_block_count: buildingBlockCount,
    hosts_per_block: hostsPerBlock,
    blocks_per_rack: blocksPerRack,
    accelerators_per_rack: acceleratorsPerRack,
    compute_rack_count: rackCount,
    hosts_per_rack: hostsPerRack,
    dcn_bandwidth_per_chip_gbps: dcnPerChip,
    dcn_bandwidth_per_host_gbps: dcnPerHost,
    nics_per_host: nicsPerHost,
    nic_speed_gbps: nicSpeedGbps,
    total_nic_count: totalNicCount,
    nic_aggregate_bandwidth_per_host_gbps: nicAggregatePerHostGbps,
    dcn_bandwidth_per_pod_tbps: dcnAggregatePodTbps,
    ici_ports_per_chip: iciPortsPerChip,
    ici_ports_total: iciPortsTotal,
    peak_bf16_tflops_per_chip: peakBf16TflopsPerChip,
    peak_bf16_pflops_per_pod: peakBf16PflopsPerPod,
    peak_fp8_tflops_per_chip: peakFp8TflopsPerChip,
    peak_fp8_pflops_per_pod: peakFp8PflopsPerPod,
    hbm_capacity_gib_per_chip: hbmCapacityGibPerChip,
    hbm_capacity_gib_per_pod: hbmCapacityGibPerPod,
    hbm_bandwidth_GBps_per_chip: hbmBandwidthGBpsPerChip,
    hbm_bandwidth_TBps_per_pod: hbmBandwidthTBpsPerPod,
    tensor_cores_per_chip: tensorCoresPerChip,
    tensor_cores_total: tensorCoresTotal,
    sparse_cores_per_chip: sparseCoresPerChip,
    sparse_cores_total: sparseCoresTotal,
    total_gpu_memory_gb_per_host: input.accelerator && input.accelerator.memory_gb_per_accelerator != null
      ? Number(input.accelerator.memory_gb_per_accelerator) * perHost : null,
    ...maxSlice
  };
}

function deriveMachineProfile(profile) {
  const accelerators = Number(profile.accelerators_per_host || 0);
  const gpuNics = Number(profile.gpu_nics_per_host || 0);
  const serviceNics = Number(profile.service_nics_per_host || 0);
  const gpuMem = profile.memory_gb_per_accelerator != null ? Number(profile.memory_gb_per_accelerator) : null;
  const gpuNetwork = profile.gpu_network_bandwidth_total_gbps != null ? Number(profile.gpu_network_bandwidth_total_gbps) : null;
  const serviceNetwork = profile.service_network_bandwidth_total_gbps != null ? Number(profile.service_network_bandwidth_total_gbps) : null;
  return {
    accelerator_count_per_host: accelerators || null,
    physical_nic_count_per_host: gpuNics + serviceNics || null,
    gpu_nic_count_per_host: gpuNics || null,
    service_nic_count_per_host: serviceNics || null,
    total_gpu_memory_gb_per_host: gpuMem != null ? gpuMem * accelerators : null,
    gpu_nic_to_accelerator_ratio: accelerators > 0 ? gpuNics / accelerators : null,
    data_network_count_per_host: profile.data_networks_per_host != null ? Number(profile.data_networks_per_host) : (gpuNics || null),
    max_network_bandwidth_gbps: gpuNetwork != null && serviceNetwork != null ? gpuNetwork + serviceNetwork :
      (profile.max_network_bandwidth_gbps != null ? Number(profile.max_network_bandwidth_gbps) : null)
  };
}

function deriveStorage(input) {
  const comps = input.storage && Array.isArray(input.storage.components_pb) ? input.storage.components_pb : null;
  return comps ? { storage_total_pb: comps.reduce((a,b)=>a+Number(b),0) } : {};
}

function deriveRackPower(input) {
  const p=input.rack_power || null;
  if(!p) return {};
  const shelf= p.bbu_shelf_kw != null ? Number(p.bbu_shelf_kw) : null;
  const rack= p.rack_power_kw != null ? Number(p.rack_power_kw) : null;
  const durationMin = p.backup_duration_minutes != null ? Number(p.backup_duration_minutes) : null;
  return {
    bbu_shelves_required_for_rack: shelf && rack ? Math.ceil(rack/shelf) : null,
    bbu_pair_power_kw: shelf != null ? shelf*2 : null,
    backup_duration_seconds: durationMin != null ? durationMin*60 : null
  };
}

function deriveOpticalTorus(input, common) {
  const t = input.topology || {};
  const result = { topology_type: "optical_torus" };

  const faces = t.faces != null ? Number(t.faces) : null;
  const linksPerFace = t.links_per_face != null ? Number(t.links_per_face) : null;
  if (faces != null && linksPerFace != null) {
    const opticalLinksPerBlock = faces * linksPerFace;
    const opticalLinksPerRack = opticalLinksPerBlock * common.blocks_per_rack;
    const rackOcsLinkEndpoints = common.compute_rack_count * opticalLinksPerRack;
    result.optical_links_per_block = opticalLinksPerBlock;
    result.optical_links_per_rack = opticalLinksPerRack;
    result.rack_to_ocs_link_instances = rackOcsLinkEndpoints;

    const ocs = t.ocs || null;
    if (ocs && ocs.total_ports != null) {
      const totalPorts = requirePositive("topology.ocs.total_ports", ocs.total_ports);
      const sparePorts = Number(ocs.spare_ports || 0);
      const workingPorts = totalPorts - sparePorts;
      if (workingPorts <= 0) throw new Error("OCS working ports must be > 0");

      const ocsCount = rackOcsLinkEndpoints / workingPorts;
      const opposingShare = t.opposing_faces_share_ocs !== false;
      result.ocs_count = ocsCount;
      result.ocs_per_block = opposingShare ? opticalLinksPerBlock / 2 : opticalLinksPerBlock;
      result.ocs_total_ports_each = totalPorts;
      result.ocs_working_ports_each = workingPorts;
      result.ocs_spare_ports_each = sparePorts;
      result.ocs_working_ports_total = ocsCount * workingPorts;
      result.ocs_spare_ports_total = ocsCount * sparePorts;
    }
  }
  return result;
}

function deriveClos(input, common) {
  const t = input.topology || {};
  const tiers = Number(t.tiers || 2);
  const rails = Number(t.rails || 1);
  const dualTor = Boolean(t.dual_tor);
  const dualPlane = Boolean(t.dual_plane);

  const serverLinks = Number((input.network && input.network.links_per_host) || rails || 1);
  const torDownlinks = Number((input.network && input.network.tor_downlinks) || 0);
  const torUplinks = Number((input.network && input.network.tor_uplinks) || 0);

  let torCount = null;
  if (torDownlinks > 0) {
    const endpointLinks = common.cpu_host_count * serverLinks * (dualTor ? 2 : 1);
    torCount = ceilDiv(endpointLinks, torDownlinks);
  }

  return {
    topology_type: "clos",
    clos_tiers: tiers,
    rail_count: rails,
    dual_tor: dualTor,
    dual_plane: dualPlane,
    server_network_links: common.cpu_host_count * serverLinks * (dualTor ? 2 : 1),
    tor_count: torCount,
    tor_downlinks_each: torDownlinks || null,
    tor_uplinks_each: torUplinks || null
  };
}

function deriveArchitecture(input) {
  if (Array.isArray(input.machine_profiles)) {
    const out={};
    input.machine_profiles.forEach(p=>{
      const d=deriveMachineProfile(p);
      Object.entries(d).forEach(([k,v])=>{ out[p.id+"__"+k]=v; });
    });
    return out;
  }
  const common = deriveCommon(input);
  Object.assign(common, deriveStorage(input), deriveRackPower(input));
  if(input.machine_profile){
    Object.assign(common, deriveMachineProfile(input.machine_profile));
  }
  const topologyType = input.topology && input.topology.type;

  let topology = {};
  if (topologyType === "optical_torus" || topologyType === "torus3d" || topologyType === "3d_torus" || topologyType === "torus2d" || topologyType === "2d_torus") {
    topology = deriveOpticalTorus(input, common);
    topology.topology_type = topologyType;
  } else if (topologyType === "clos" || topologyType === "leaf_spine") {
    topology = deriveClos(input, common);
  } else if (topologyType) {
    topology = { topology_type: topologyType, unsupported_topology: true };
  }

  return Object.assign({}, common, topology);
}

function errorPct(calculated, reference) {
  if (calculated == null || reference == null) return null;
  if (Number(reference) === 0) return Number(calculated) === 0 ? 0 : 100;
  return Math.abs(Number(calculated) - Number(reference)) / Math.abs(Number(reference)) * 100;
}

function compareReference(calculated, referenceMetrics, thresholdPct) {
  const threshold = Number(thresholdPct == null ? 10 : thresholdPct);
  const metrics = {};
  const errors = [];
  let comparable = 0;
  let verifiable = 0;

  Object.entries(referenceMetrics || {}).forEach(([key, meta]) => {
    if (meta && meta.verifiable === false) return;
    verifiable += 1;
    const reference = meta && Object.prototype.hasOwnProperty.call(meta, "value") ? meta.value : meta;
    const calc = Object.prototype.hasOwnProperty.call(calculated, key) ? calculated[key] : null;
    const err = errorPct(calc, reference);
    const status = err == null ? "NOT_SUPPORTED" : (err <= threshold ? "PASS" : "FAIL");
    if (err != null) {
      comparable += 1;
      errors.push(err);
    }
    metrics[key] = {
      reference,
      calculated: calc,
      error_pct: err,
      status
    };
  });

  const mape = errors.length ? errors.reduce((a,b)=>a+b,0) / errors.length : null;
  const maxError = errors.length ? Math.max(...errors) : null;
  const coverage = verifiable ? comparable / verifiable * 100 : 0;
  const failed = Object.values(metrics).some(m => m.status === "FAIL");

  return {
    status: comparable === 0 ? "NOT_SUPPORTED" : failed ? "FAIL" : coverage < 100 ? "PARTIAL" : "PASS",
    pass_threshold_error_pct: threshold,
    coverage_pct: coverage,
    mape_pct: mape,
    max_error_pct: maxError,
    metrics
  };
}

const api = { deriveArchitecture, compareReference, errorPct };
if (typeof module !== "undefined" && module.exports) module.exports = api;
if (typeof window !== "undefined") window.MultiArchBomEngine = api;
