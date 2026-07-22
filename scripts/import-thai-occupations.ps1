param(
  [string]$OutputPath = (Join-Path $PSScriptRoot '..\thai-occupations-data.js')
)

$sourcePage = 'https://www.doe.go.th/prd/lmia/custom/param/site/131/cat/40/sub/0/pull/detail/view/detail/object_id/2790?token=%3Ftoken%3D'
$html = (Invoke-WebRequest -UseBasicParsing $sourcePage).Content
$listMatch = [regex]::Match($html, '<ol>(.*?)</ol>', 'Singleline,IgnoreCase')
$groupNames = @([regex]::Matches($listMatch.Groups[1].Value, '<li[^>]*>(.*?)</li>', 'Singleline,IgnoreCase') | Select-Object -First 9 | ForEach-Object {
  $text = [System.Net.WebUtility]::HtmlDecode(([regex]::Replace($_.Groups[1].Value, '<[^>]+>', ''))).Trim()
  ($text -replace '^.*?:\s*', '').Trim()
})
if ($groupNames.Count -ne 9) { throw 'Expected 9 occupation groups from the source page' }

$tableMatch = [regex]::Match($html, '<table\s+id="table31252"[^>]*>(.*?)</table>', 'Singleline,IgnoreCase')
if (-not $tableMatch.Success) { throw 'Occupation table table31252 was not found' }

$occupations = [System.Collections.Generic.List[object]]::new()
$seen = [System.Collections.Generic.HashSet[string]]::new()
$rows = [regex]::Matches($tableMatch.Groups[1].Value, '<tr[^>]*>(.*?)</tr>', 'Singleline,IgnoreCase')
foreach ($row in $rows | Select-Object -Skip 1) {
  $cells = [regex]::Matches($row.Groups[1].Value, '<td[^>]*>(.*?)</td>', 'Singleline,IgnoreCase')
  for ($column = 0; $column -lt [Math]::Min($cells.Count, $groupNames.Count); $column++) {
    $links = [regex]::Matches($cells[$column].Groups[1].Value, '<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', 'Singleline,IgnoreCase')
    foreach ($link in $links) {
      $name = [System.Net.WebUtility]::HtmlDecode(([regex]::Replace($link.Groups[2].Value, '<[^>]+>', ''))).Trim()
      if (-not $name -or -not $seen.Add($name)) { continue }
      $occupations.Add([ordered]@{
        name = $name
        group = $column + 1
        source = [System.Net.WebUtility]::HtmlDecode($link.Groups[1].Value)
      })
    }
  }
}

if ($occupations.Count -lt 400) { throw "Only $($occupations.Count) occupations were found; expected at least 400" }

$groupsJson = $groupNames | ConvertTo-Json -Compress
$occupationsJson = $occupations | ConvertTo-Json -Depth 3 -Compress
$generated = @"
// Generated from the Department of Employment occupation classification. Do not edit by hand.
const thaiOccupationSource = '$sourcePage';
const thaiOccupationGroups = $groupsJson;
const thaiOccupations = $occupationsJson;

function normalizeThaiOccupationText(text) {
  return String(text || '').toLocaleLowerCase('th-TH').replace(/\s+/g, '').replace(/[^\p{L}\p{N}]+/gu, '');
}

function findThaiOccupations(query, limit = 20) {
  const normalizedQuery = normalizeThaiOccupationText(query);
  if (!normalizedQuery) return [];
  return thaiOccupations
    .map((occupation, index) => {
      const name = normalizeThaiOccupationText(occupation.name);
      const score = name === normalizedQuery ? 0
        : name.startsWith(normalizedQuery) ? 1
        : name.includes(normalizedQuery) || normalizedQuery.includes(name) ? 2
        : 99;
      return { occupation, score, index };
    })
    .filter(item => item.score < 99)
    .sort((a, b) => a.score - b.score || a.index - b.index)
    .slice(0, limit)
    .map(item => item.occupation);
}
"@

[System.IO.File]::WriteAllText((Resolve-Path (Split-Path $OutputPath)).Path + '\' + (Split-Path $OutputPath -Leaf), $generated, [System.Text.UTF8Encoding]::new($false))
Write-Output "Imported $($occupations.Count) occupations from the Department of Employment"
