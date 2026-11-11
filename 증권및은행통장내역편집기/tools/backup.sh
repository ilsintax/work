#!/usr/bin/env bash
# 변경 전 부분 백업 (공통 도구 Backup-Project.ps1 이 SynologyDrive 폴더에서 실행되지 않아 쓰는 대체 방식)
# 사용: bash tools/backup.sh <작업명> <작업ID> <파일1> [파일2 ...]   (프로젝트 루트에서)
# 결과: version/YYYYMMDD_HHMMSS_작업명/{files/, manifest.json, RESTORE.md, VERIFIED(전부 일치할 때만)}
set -euo pipefail
[ $# -ge 3 ] || { echo "사용: bash tools/backup.sh <작업명> <작업ID> <파일...>"; exit 2; }
name="$1"; task="$2"; shift 2
dest="version/$(date +%Y%m%d_%H%M%S)_${name}"
[ -e "$dest" ] && dest="${dest}_$RANDOM"
mkdir -p "$dest/files"
total=0; ok=0; entries=""
for f in "$@"; do
  [ -f "$f" ] || { echo "없는 파일: $f"; exit 1; }
  mkdir -p "$dest/files/$(dirname "$f")"
  cp -p "$f" "$dest/files/$f"
  s=$(sha256sum "$f" | cut -d' ' -f1); c=$(sha256sum "$dest/files/$f" | cut -d' ' -f1)
  total=$((total + 1)); v=false; [ "$s" = "$c" ] && { ok=$((ok + 1)); v=true; }
  entries="${entries}${entries:+,
}    {\"path\": \"$f\", \"size\": $(stat -c %s "$f"), \"sha256\": \"$s\", \"verified\": $v}"
done
cat > "$dest/manifest.json" <<EOF
{
  "taskId": "$task",
  "createdAt": "$(date '+%Y-%m-%dT%H:%M:%S%z')",
  "sourceRoot": "$(pwd -W 2>/dev/null || pwd)",
  "type": "partial",
  "reason": "$name (tools/backup.sh: cp + sha256sum 대체 방식)",
  "files": [
$entries
  ],
  "excluded": "목록에 없는 파일(변경 대상 아님), node_modules, tests/output"
}
EOF
printf '# 복원 방법\n\n1. 현재 상태를 먼저 별도 백업합니다.\n2. files/ 아래 파일을 같은 상대경로로 프로젝트 루트에 복사합니다.\n3. manifest.json 의 sha256 과 비교합니다.\n' > "$dest/RESTORE.md"
if [ "$ok" = "$total" ]; then echo "verified $(date '+%FT%T%z') $ok/$total" > "$dest/VERIFIED"; echo "백업 완료: $dest ($ok/$total)"; else echo "백업 검증 실패: $dest ($ok/$total)"; exit 1; fi
