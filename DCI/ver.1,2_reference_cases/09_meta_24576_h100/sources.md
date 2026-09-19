# Sources

- Meta Engineering, Building Meta's GenAI Infrastructure:
  https://engineering.fb.com/2024/03/12/data-center-engineering/building-metas-genai-infrastructure/
  - two 24,576-H100 clusters
  - RoCE: Arista 7800 + Wedge400 + Minipack2
  - second cluster: NVIDIA Quantum-2 InfiniBand
  - 400 Gbps endpoints
  - Grand Teton compute platform

- OCP Grand Teton AMD-based CPU Tray Specification:
  https://www.opencompute.org/documents/grand-teton-amd-based-cpu-tray-specification-v1-0-pdf
  - accelerator tray shows GPU 0 through GPU 7
