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
  const targetAccelerators = requirePositive("target_accelerators", input.target_accelerators);
  const perHost = requirePositive("accelerator.per_host", input.accelerator && input.accelerator.per_host);

  let acceleratorsPerBlock = null;
  if (input.building_block && Array.isArray(input.building_block.dimensions)) {
    acceleratorsPerBlock = product(input.building_block.dimensions);
  } else if (input.building_block && input.building_block.accelerators_per_block) {
    acceleratorsPerBlock = requirePositive(
      "building_block.accelerators_per_block",
      input.building_block.accelerators_per_block
    );
  }

  const blocksPerRack = Number((input.rack && input.rack.blocks_per_rack) || 1);
  const acceleratorsPerRack = input.rack && input.rack.accelerators_per_rack
    ? requirePositive("rack.accelerators_per_rack", input.rack.accelerators_per_rack)
    : acceleratorsPerBlock
      ? acceleratorsPerBlock * blocksPerRack
      : null;

  const hostCount = targetAccelerators / perHost;
  const rackCount = acceleratorsPerRack ? ceilDiv(targetAccelerators, acceleratorsPerRack) : null;
  const hostsPerRack = acceleratorsPerRack ? acceleratorsPerRack / perHost : null;

  return {
    accelerator_count: targetAccelerators,
    accelerators_per_host: perHost,
    cpu_host_count: hostCount,
    accelerators_per_block: acceleratorsPerBlock,
    blocks_per_rack: blocksPerRack,
    accelerators_per_rack: acceleratorsPerRack,
    compute_rack_count: rackCount,
    hosts_per_rack: hostsPerRack
  };
}

function deriveOpticalTorus(input, common) {
  const t = input.topology || {};
  const ocs = t.ocs || {};

  const faces = requirePositive("topology.faces", t.faces);
  const linksPerFace = requirePositive("topology.links_per_face", t.links_per_face);
  const totalPorts = requirePositive("topology.ocs.total_ports", ocs.total_ports);
  const sparePorts = Number(ocs.spare_ports || 0);
  const workingPorts = totalPorts - sparePorts;
  if (workingPorts <= 0) throw new Error("OCS working ports must be > 0");

  const opticalLinksPerBlock = faces * linksPerFace;
  const opticalLinksPerRack = opticalLinksPerBlock * common.blocks_per_rack;
  const rackOcsLinkEndpoints = common.compute_rack_count * opticalLinksPerRack;
  const ocsCount = rackOcsLinkEndpoints / workingPorts;
  const opposingShare = t.opposing_faces_share_ocs !== false;
  const ocsPerBlock = opposingShare ? opticalLinksPerBlock / 2 : opticalLinksPerBlock;

  return {
    topology_type: "optical_torus",
    optical_links_per_block: opticalLinksPerBlock,
    optical_links_per_rack: opticalLinksPerRack,
    rack_to_ocs_link_instances: rackOcsLinkEndpoints,
    ocs_count: ocsCount,
    ocs_per_block: ocsPerBlock,
    ocs_total_ports_each: totalPorts,
    ocs_working_ports_each: workingPorts,
    ocs_spare_ports_each: sparePorts,
    ocs_working_ports_total: ocsCount * workingPorts,
    ocs_spare_ports_total: ocsCount * sparePorts
  };
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
  const common = deriveCommon(input);
  const topologyType = input.topology && input.topology.type;

  let topology = {};
  if (topologyType === "optical_torus") {
    topology = deriveOpticalTorus(input, common);
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
