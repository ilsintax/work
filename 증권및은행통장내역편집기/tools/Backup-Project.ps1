param(
 [Parameter(Mandatory=$true)][string]$ProjectRoot,
 [Parameter(Mandatory=$true)][ValidatePattern('^[\p{L}\p{N}_-]+$')][string]$TaskName,
 [Parameter(Mandatory=$true)][string]$TaskId,
 [string[]]$ExcludeDirectoryNames=@()
)
$ErrorActionPreference='Stop'
$root=(Get-Item -LiteralPath $ProjectRoot -Force).FullName
if(-not (Test-Path -LiteralPath $root -PathType Container)){throw 'ProjectRoot must be a directory'}
if((Get-Item -LiteralPath $root -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked root is unsupported'}
$excluded=@('version','.git')+$ExcludeDirectoryNames
foreach($n in $excluded){if([string]::IsNullOrWhiteSpace($n) -or $n -match '[/\\]' -or $n -in @('.','..')){throw 'Exclusions must be directory names'}}
$script:omitted=[Collections.Generic.List[string]]::new()
function Inventory([string]$dir){
 foreach($item in Get-ChildItem -LiteralPath $dir -Force){
  if($item.PSIsContainer -and $item.Name -in $excluded){$script:omitted.Add([IO.Path]::GetRelativePath($root,$item.FullName));continue}
  if($item.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "Link/reparse point requires separate handling: $($item.FullName)"}
  if($item.PSIsContainer){Inventory $item.FullName}else{$item}
 }
}
$initial=@(Inventory $root | Sort-Object FullName)
$version=Join-Path $root 'version'
if(Test-Path -LiteralPath $version){if((Get-Item -LiteralPath $version -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked version directory is unsupported'}}
$stamp=[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTimeOffset]::UtcNow,'Korea Standard Time').ToString('yyyyMMdd_HHmmss')
$dest=Join-Path $version ($stamp+'_'+$TaskName+'_'+[guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path (Join-Path $dest 'files') -Force | Out-Null
$records=@(foreach($f in $initial){
 $relative=[IO.Path]::GetRelativePath($root,$f.FullName)
 $hash=(Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash
 $target=Join-Path (Join-Path $dest 'files') $relative
 New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
 Copy-Item -LiteralPath $f.FullName -Destination $target
 if((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $hash){throw "Copy mismatch: $relative"}
 [pscustomobject]@{path=$relative;size=(Get-Item -LiteralPath $target).Length;sha256=$hash}
})
$final=@(Inventory $root | Sort-Object FullName)
if(($initial.FullName -join "`n") -cne ($final.FullName -join "`n")){throw 'Source file list changed during backup'}
foreach($record in $records){if((Get-FileHash -LiteralPath (Join-Path $root $record.path) -Algorithm SHA256).Hash -ne $record.sha256){throw "Source changed: $($record.path)"}}
$manifest=[ordered]@{schemaVersion=1;taskId=$TaskId;taskName=$TaskName;createdAt=[DateTimeOffset]::UtcNow.ToOffset([TimeSpan]::FromHours(9)).ToString('o');sourceRoot=$root;scope='all regular files except explicit exclusions; empty directories and filesystem metadata not preserved';excludeDirectoryNames=$excluded;excludedPaths=@($script:omitted|Sort-Object -Unique);verified=$true;files=$records;restore='See RESTORE.md; recheck hashes before restoring';limitations=@('Not an atomic live-data snapshot','External files are not included','No independent remote backup')}
$manifest|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $dest 'manifest.json') -Encoding utf8
@'
# 복원
1. 현재 작업본을 별도로 백업합니다. 원본 경로가 새 컴퓨터에서 맞는지 확인합니다.
2. manifest.json의 모든 path에 대해 files/ 사본의 크기와 SHA256을 다시 확인합니다.
3. files/를 빈 임시 폴더에 복사하고 해시를 비교합니다. 제외 항목·외부 파일·빈 폴더·설치 의존성을 별도로 준비합니다.
4. 해당 프로젝트의 설치·기본 실행 검증을 수행합니다. 실행하지 못했다면 파일 일치만 확인했다고 기록합니다.
5. 확인된 파일을 원래 상대경로로 반영합니다. 작업 이후 신규 파일과 이동 파일은 HISTORY를 참고해 별도로 처리합니다. 임의 일괄 삭제하지 않습니다.
6. 복원 결과와 검증을 HISTORY에 추가하고 HANDOFF를 갱신합니다.
'@ | Set-Content -LiteralPath (Join-Path $dest 'RESTORE.md') -Encoding utf8
'Verified SHA256 and source file list at backup time.'|Set-Content -LiteralPath (Join-Path $dest 'VERIFIED') -Encoding utf8
$dest

