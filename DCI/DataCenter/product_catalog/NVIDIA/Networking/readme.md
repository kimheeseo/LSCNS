# NVIDIA Networking

NVIDIA 공식 네트워킹 페이지를 기준으로 데이터센터 BOM/제품 조회용 메타데이터를 정리합니다.

## 제품군

- InfiniBand: Quantum-X800, Quantum-2, Skyway, MetroX-3 XC
- Ethernet switches: Spectrum-6 SN6000, Spectrum-4 SN5000, Spectrum-3 SN4000, Spectrum-2 SN3000, Spectrum SN2000
- SuperNIC / NIC: ConnectX-9, ConnectX-8, BlueField-3 SuperNIC, ConnectX-7, ConnectX-6 Dx, ConnectX-6 Lx
- DPU / Storage: BlueField-4 DPU, BlueField-4 STX, BlueField-3 DPU
- Platform: Spectrum-X Ethernet
- Interconnect / optical transceivers: LinkX 400G OSFP, 400G QSFP-DD, 200G QSFP56, 100G QSFP28

공식 링크와 공개 사양만 catalog.json에 기록하며, 실제 SKU·폼팩터·거리·호스트 호환성은 제품 링크에서 최종 확인합니다.

## Additional NVIDIA resources (2026-10-09)

The following ten user-provided preview links were checked against official NVIDIA resource titles. Prefer the canonical URLs below in product cards; the original preview URLs are also retained as `requestedUrl` in the research catalog.

1. **NVIDIA Quantum-X800 InfiniBand Platform · Solution Overview** (Solution Overview, Quantum-X800 InfiniBand)
   - Official: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/solution-overview-gt
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/solution-overview-gt?pflpid=8026&&lb-mode=preview
2. **NVIDIA BlueField-3 DPU · Datasheet** (Datasheet, BlueField-3 DPU)
   - Official: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/datasheet-nvidia-bluefield
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/datasheet-nvidia-bluefield?pflpid=8026&&lb-mode=preview
3. **NVIDIA ConnectX-9 SuperNIC · Datasheet** (Datasheet, ConnectX-9 SuperNIC)
   - Official: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/connectx-9-supernic-datasheet
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/connectx-9-supernic-datasheet?pflpid=8026&&lb-mode=preview
4. **NVIDIA Spectrum-X Ethernet · Datasheet** (Datasheet, Spectrum-X Ethernet)
   - Official: https://resources.nvidia.com/en-us-networking-ai/networking-ethernet-1
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/networking-ethernet-1?pflpid=8026&&lb-mode=preview
5. **NVIDIA BlueField-4 DPU · Datasheet** (Datasheet, BlueField-4 DPU)
   - Official: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/bluefield-4-dpu-datasheet
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/bluefield-4-dpu-datasheet?pflpid=8026&&lb-mode=preview
6. **NVIDIA Spectrum SN6000 Ethernet Switch · Datasheet** (Datasheet, Spectrum-6 SN6000)
   - Official: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/ethernet-datasheet-spectrum-sn6000-switch
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/ethernet-datasheet-spectrum-sn6000-switch?pflpid=8026&&lb-mode=preview
7. **How AI Factories Can Shift Left to Accelerate Infrastructure Readiness** (Technical Article, AI Factory Design)
   - Official: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/how-ai-factories-can-shift-left-to-accelerate-1179169
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/how-ai-factories-can-shift-left-to-accelerate-1179169?pflpid=8026&&lb-mode=preview
8. **NVIDIA DSX Air Boosts Time to Token With Accelerated Simulation** (Technical Article, DSX Air Simulation)
   - Official: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/nvidia-dsx-air-boosts-time-to-token-with-1179168
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/nvidia-dsx-air-boosts-time-to-token-with-1179168?pflpid=8026&&lb-mode=preview
9. **NVIDIA BlueField-4 Powers New Scale-In Network Infrastructure for Agentic AI Factories** (Technical Article, BlueField-4 DPU)
   - Official: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/nvidia-bluefield-4-powers-new-scale-in-network-infrastructure-for-agentic-ai-factories
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/nvidia-bluefield-4-powers-new-scale-in-network-infrastructure-for-agentic-ai-factories?pflpid=8026&&lb-mode=preview
10. **Giga-Scale AI and the Ethernet Evolution: How Spectrum-X Ethernet Rewrites the Rules** (Technical Article, Spectrum-X Ethernet)
   - Official: https://developer.nvidia.com/blog/giga-scale-ai-ethernet-evolution-spectrum-x-ethernet-rewrites-rules/
   - User preview: https://resources.nvidia.com/en-us-accelerated-networking-resource-library-ms/en-us-accelerated-networking-resource-library/giga-scale-ai-ethernet-evolution-spectrum-x-ethernet-rewrites-rules?pflpid=8026&&lb-mode=preview

All ten records have `vendorReferenceOnly: true` and are browsable but excluded from automatic BOM/SKU matching. Documents without verified product-specific detailed specifications are deliberately not assigned new numerical spec claims.
