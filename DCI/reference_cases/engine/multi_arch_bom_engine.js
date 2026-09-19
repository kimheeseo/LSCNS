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
  const bandwidthPerHostGbps = input.network && input.network.bandwidth_per_host_gbps != null
    ? Number(input.network.bandwidth_per_host_gbps) : null;
  const clusterNetworkEndpointTbps = bandwidthPerHostGbps != null ? hostCount * bandwidthPerHostGbps / 1000 : null;

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
    network_bandwidth_per_host_gbps: bandwidthPerHostGbps,
    cluster_network_endpoint_tbps: clusterNetworkEndpointTbps,
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


function deriveNetworkPlan(input) {
  const n=input.network_plan||{};
  const hosts=Number(n.hosts||0);
  const linksPerHost=Number(n.links_per_host||0);
  const redundancy=Number(n.link_redundancy_factor||1);
  const logicalDown=Number(n.logical_downlinks_per_tor||0);
  const physicalDown=Number(n.physical_downlinks_per_tor||0);
  const uplinks=Number(n.uplinks_per_tor||0);
  const backupDown=Number(n.backup_downlinks_per_tor||0);
  const serverLinks=hosts*linksPerHost*redundancy;
  const torCount=logicalDown>0?Math.ceil(serverLinks/logicalDown):null;
  return {
    host_count:hosts,
    server_network_links:serverLinks,
    tor_count:torCount,
    tor_logical_downlinks_total:torCount!=null?torCount*logicalDown:null,
    tor_physical_downlinks_total:torCount!=null&&physicalDown?torCount*physicalDown:null,
    tor_uplinks_total:torCount!=null&&uplinks?torCount*uplinks:null,
    tor_backup_downlinks_total:torCount!=null&&backupDown?torCount*backupDown:null
  };
}

function deriveFabricDesign(input) {
  const f=input.fabric_design||{};
  const active=Number(f.active_nodes||0);
  const slots=Number(f.design_nodes||active);
  const links=Number(f.links_per_node||0);
  const ufm=Number(f.ufm_links||0);
  const leafDown=Number(f.leaf_down_ports||0);
  const leafUp=Number(f.leaf_up_ports||0);
  const spineDown=Number(f.spine_down_ports||f.spine_ports||0);
  const spineUp=Number(f.spine_up_ports||0);
  const corePorts=Number(f.core_ports||0);
  const endpointCables=active*links+ufm;
  const designEndpointPorts=slots*links;
  const leafCount=leafDown?Math.ceil(designEndpointPorts/leafDown):null;
  const leafSpineLinks=leafCount!=null&&leafUp?leafCount*leafUp:null;
  const spineCount=leafSpineLinks!=null&&spineDown?Math.ceil(leafSpineLinks/spineDown):null;
  const spineCoreLinks=spineCount!=null&&spineUp?spineCount*spineUp:null;
  const coreCount=spineCoreLinks!=null&&corePorts?Math.ceil(spineCoreLinks/corePorts):null;
  return {
    node_leaf_cable_count:endpointCables,
    leaf_switch_count:leafCount,
    leaf_spine_cable_count:leafSpineLinks,
    spine_switch_count:spineCount,
    spine_core_cable_count:spineCoreLinks,
    core_switch_count:coreCount
  };
}

function deriveRackSystem(input) {
  const r=input.rack_system||{};
  const racks=Number(r.racks||1);
  const trays=Number(r.compute_trays_per_rack||0);
  const gpuPerTray=Number(r.gpus_per_tray||0);
  const cpuPerTray=Number(r.cpus_per_tray||0);
  const switchTrays=Number(r.nvlink_switch_trays_per_rack||0);
  const switchesPerTray=Number(r.nvswitches_per_switch_tray||0);
  const tor=Number(r.tor_switches_per_rack||0);
  const powerShelves=Number(r.power_shelves_per_rack||0);
  const psusPerShelf=Number(r.psus_per_power_shelf||0);
  const psuKw=Number(r.psu_kw||0);
  const nvmePerTray=Number(r.data_nvme_per_tray||0);
  const nvmeTb=Number(r.data_nvme_tb_each||0);
  const bootPerTray=Number(r.boot_nvme_per_tray||0);
  const bootTb=Number(r.boot_nvme_tb_each||0);
  return {
    compute_tray_count:racks*trays,
    accelerator_count:racks*trays*gpuPerTray,
    cpu_count:racks*trays*cpuPerTray,
    nvlink_switch_tray_count:racks*switchTrays,
    nvswitch_count:racks*switchTrays*switchesPerTray,
    tor_switch_count:racks*tor,
    power_shelf_count:racks*powerShelves,
    psu_count:racks*powerShelves*psusPerShelf,
    installed_psu_capacity_kw:racks*powerShelves*psusPerShelf*psuKw,
    data_nvme_count:racks*trays*nvmePerTray,
    data_nvme_capacity_tb:racks*trays*nvmePerTray*nvmeTb,
    boot_nvme_count:racks*trays*bootPerTray,
    boot_nvme_capacity_tb:racks*trays*bootPerTray*bootTb
  };
}

function deriveHpcSystem(input) {
  const h=input.hpc_system||{};
  const nodes=Number(h.nodes||0);
  const acc=Number(h.accelerators_per_node||0);
  const visible=Number(h.visible_gpus_per_node||acc);
  const nics=Number(h.nics_per_node||0);
  const nicGbps=Number(h.nic_speed_gbps||0);
  const racks=Number(h.racks||0);
  return {
    node_count:nodes,
    accelerator_count:nodes*acc,
    visible_gpu_count:nodes*visible,
    nic_count:nodes*nics,
    injection_bandwidth_per_node_gbps:nics*nicGbps,
    aggregate_endpoint_injection_pbps:nodes*nics*nicGbps/1e6,
    aggregate_endpoint_injection_PBps:nodes*nics*nicGbps/8/1e6,
    nodes_per_rack:racks?nodes/racks:null
  };
}

function deriveScaleUnit(input) {
  const s=input.scale_unit||{};
  const racks=Number(s.racks||1);
  const trays=Number(s.trays_per_rack||0);
  const gpuPerTray=Number(s.gpus_per_tray||0);
  const computeNics=Number(s.compute_nics_per_tray||0);
  const computeNicGbps=Number(s.compute_nic_speed_gbps||0);
  const convergedLinks=Number(s.converged_links_per_tray||0);
  const convergedGbps=Number(s.converged_link_speed_gbps||0);
  const groups=Number(s.fabric_group_count||0);
  const leafPerGroup=Number(s.leaf_switches_per_group||0);
  const spinePerGroup=Number(s.spine_switches_per_group||0);
  return {
    tray_count:racks*trays,
    accelerator_count:racks*trays*gpuPerTray,
    compute_nic_count:racks*trays*computeNics,
    compute_bandwidth_tbps:racks*trays*computeNics*computeNicGbps/1000,
    converged_link_count:racks*trays*convergedLinks,
    converged_bandwidth_tbps:racks*trays*convergedLinks*convergedGbps/1000,
    fabric_group_count:groups||null,
    leaf_switch_count:groups&&leafPerGroup?groups*leafPerGroup:null,
    spine_switch_count:groups&&spinePerGroup?groups*spinePerGroup:null
  };
}

function deriveClusterScale(input) {
  const s=input.cluster_scale||{};
  const aps=Number(s.accelerators_per_system||0);
  const targetAcc=Number(s.target_accelerators||0);
  const systems=Number(s.systems||0) || (targetAcc&&aps?Math.ceil(targetAcc/aps):0);
  const gpuMem=Number(s.gpu_memory_gb_per_accelerator||0);
  const nvmeCount=Number(s.nvme_devices_per_system||0);
  const nvmeTb=Number(s.nvme_tb_each||0);
  const expansion=Number(s.expansion_factor||0);
  const nicsPerAcc=Number(s.network_nics_per_accelerator||0);
  const networkGbps=Number(s.network_bandwidth_gbps_per_system||0);
  return {
    system_count:systems,
    compute_core_count:systems*Number(s.compute_cores_per_system||0),
    cpu_count:systems*Number(s.cpus_per_system||0),
    memory_tb:systems*Number(s.memory_tb_per_system||0),
    fabric_bandwidth_tbps:systems*Number(s.fabric_bandwidth_tbps_per_system||0),
    accelerator_count:systems*aps,
    node_count:systems||null,
    gpu_memory_gb:systems*aps*gpuMem,
    nvme_count:systems*nvmeCount,
    nvme_capacity_tb:systems*nvmeCount*nvmeTb,
    network_nic_count:systems*aps*nicsPerAcc,
    network_bandwidth_gbps:systems*networkGbps,
    expanded_accelerator_count:expansion?systems*aps*expansion:null
  };
}

function deriveClosUnit(input) {
  const u=input.clos_unit||{};
  const tor=Number(u.tor_count||0), spine=Number(u.spine_count||0);
  const linksPerPair=Number(u.links_per_tor_spine_pair||0);
  const speed=Number(u.link_speed_gbps||0);
  const links=tor*spine*linksPerPair;
  return {
    tor_count:tor,spine_count:spine,
    tor_spine_link_count:links,
    tor_spine_aggregate_tbps:links*speed/1000
  };
}

function deriveArchitecture(input) {
  if (input.network_plan) return deriveNetworkPlan(input);
  if (input.clos_unit) return deriveClosUnit(input);
  if (input.fabric_design) return deriveFabricDesign(input);
  if (input.rack_system) return deriveRackSystem(input);
  if (input.hpc_system) return deriveHpcSystem(input);
  if (input.scale_unit) return deriveScaleUnit(input);
  if (input.cluster_scale) return deriveClusterScale(input);
  if (input.rack_power && !input.accelerator && !Array.isArray(input.machine_profiles)) {
    return deriveRackPower(input);
  }
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
