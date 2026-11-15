# Cloudflare Pages 배포
기존 GitHub Pages와 BOM 도구는 유지하며, Cloudflare 전용 빌드 결과의 첫 화면만 3D 투어로 연결합니다.

Cloudflare에서 Workers & Pages → Pages → Git 저장소 연결을 선택하고 다음을 설정합니다.

| 설정 | 값 |
| --- | --- |
| 저장소 | kimheeseo/LSCNS |
| Production branch | main |
| Framework preset | None |
| Root directory | DCI/DataCenter |
| Build command | node build-cloudflare.mjs |
| Build output directory | .cloudflare-dist |

프로젝트 이름은 사용 가능한 이름을 선택합니다. 무료 pages.dev 주소는 Cloudflare가 배포 후 발급합니다.
GitHub 연동을 완료하면 main 변경 시 자동 배포됩니다. 실제 배포 완료 여부는 Cloudflare 대시보드에서 확인합니다.

빌드: 정적 파일을 별도 디렉터리에 복사하고 첫 화면을 3D 투어로 연결합니다. 테스트·서버 코드와 숨김 파일은 제외하며 25 MiB 초과 파일은 빌드를 중단합니다.
로그인, GitHub 저장소 접근 연결은 Cloudflare 계정 소유자가 완료해야 합니다.
