(() => {
  'use strict';
  window.DCBOMEngine = {
    SYSTEMS: {
      h200:{name:'NVIDIA DGX H200',gpu:'H200',gpus:8,ru:8,power:10.2,links:8,linkSpeed:400,cordCount:6},
      b200:{name:'NVIDIA DGX B200',gpu:'B200',gpus:8,ru:10,power:14.3,links:8,linkSpeed:400,cordCount:6},
      b300:{name:'NVIDIA DGX B300',gpu:'B300',gpus:8,ru:10,power:15,links:8,linkSpeed:800,cordCount:12}
    }
  };
  window.DC_PRIVATE_API = 'https://datacenter-api-function-production.up.railway.app';
})();
