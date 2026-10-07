
#requires -Version 7.0

<#
.SYNOPSIS
    Generates interactive CPRS Git activity reports.

.DESCRIPTION
    Analyzes local Git history across CPRS repositories and generates
    a searchable HTML report, CSV export, or both.

    Reports include commit details, authors, branches containing each
    commit, changed files, and merge information.

    A comprehensive summary is displayed directly in PowerShell.

    The script does not fetch, pull, push, or modify Git repositories.

.PARAMETER Repositories
    One or more repository directories.

    Default repositories:
      C:\Development-CPRS\Cprs
      C:\Development-CPRS\cprs-batch
      C:\Development-CPRS\cprs-sasprogs

.PARAMETER Days
    Number of days of commit history to retrieve, based on committer
    date. Valid range: 1–3650. Default: 30.

.PARAMETER Author
    Filter commits by author name or email address.
    Case-insensitive partial matching.

.PARAMETER SortBy
    Sort commits by Date, Author, Repository, or Branch.
    Default: Date.

.PARAMETER Direction
    Ascending or Descending.
    Default: Descending.

.PARAMETER Format
    HTML, CSV, or Both.
    Default: HTML.

.PARAMETER OutputDirectory
    Destination for report files.
    Default: $HOME\Documents\GitReports

.PARAMETER AllBranches
    Include commits reachable from all locally available Git refs.

    Without this option, only commits reachable from HEAD are included.

.PARAMETER OpenReport
    Automatically open the generated HTML report.

.PARAMETER PassThru
    Return a structured PowerShell result object in addition to the
    formatted console summary.

    Useful when calling the script from other PowerShell scripts.

.EXAMPLE
    gitreport

    Generate a 30-day HTML report.

.EXAMPLE
    gitreport -Help

    Display complete documentation through the profile wrapper.

.EXAMPLE
    gitreport -Days 7 -AllBranches -OpenReport

    Open a report covering the last seven days across all branches.

.EXAMPLE
    gitreport -Days 90 -AllBranches -Format Both -OpenReport

    Generate and open HTML and CSV reports covering 90 days.

.EXAMPLE
    gitreport -Author "Leon" -Days 30 -OpenReport

    Report commits authored by Leon.

.EXAMPLE
    gitreport -SortBy Repository -Direction Ascending -OpenReport

    Sort activity alphabetically by repository.

.EXAMPLE
    gitreport -SortBy Author -OpenReport

    Sort activity by developer.

.EXAMPLE
    gitreport -Repositories "C:\Development-CPRS\Cprs" -Format CSV

    Export activity for one repository.

.EXAMPLE
    gitreport -Days 30 -Format Both -PassThru

    Display the summary and return a structured result object.

.NOTES
    Requires PowerShell 7 or later and Git in PATH.

    Reporting uses locally available Git history. Remote-tracking
    branches reflect the last successful fetch.

    Branches shown are references currently containing the commit,
    not necessarily the branch where it originated.

    Merge file changes are compared against the first parent.

    Git does not provide a reliable original push date.
#>

[CmdletBinding()]
param(
    [string[]]$Repositories = @(
        'C:\Development-CPRS\Cprs'
        'C:\Development-CPRS\cprs-batch'
        'C:\Development-CPRS\cprs-sasprogs'
    ),

    [ValidateRange(1, 3650)]
    [int]$Days = 30,

    [string]$Author,

    [ValidateSet('Date', 'Author', 'Repository', 'Branch')]
    [string]$SortBy = 'Date',

    [ValidateSet('Descending', 'Ascending')]
    [string]$Direction = 'Descending',

    [ValidateSet('HTML', 'CSV', 'Both')]
    [string]$Format = 'HTML',

    [string]$OutputDirectory = (
        Join-Path $HOME 'Documents\GitReports'
    ),

    [switch]$AllBranches,

    [switch]$OpenReport,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =============================================================================
# INITIALIZATION
# =============================================================================

$started = Get-Date

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw 'Git executable was not found in PATH.'
}

$null = New-Item -ItemType Directory -Path $OutputDirectory -Force

$records = [System.Collections.Generic.List[object]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()
$repositorySummary = [System.Collections.Generic.List[object]]::new()

$since = $started.AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:sszzz')

# =============================================================================
# GIT DATA COLLECTION
# =============================================================================

foreach ($repo in $Repositories) {

    $repoName = Split-Path -Path $repo -Leaf

    if (-not (Test-Path -LiteralPath $repo -PathType Container)) {

        $warnings.Add("Repository missing: $repo")

        $repositorySummary.Add(
            [pscustomobject]@{
                Name    = $repoName
                Path    = $repo
                Status  = 'Missing'
                Commits = 0
            }
        )

        continue
    }

    $root = & git -C $repo rev-parse --show-toplevel 2>$null

    if ($LASTEXITCODE -ne 0) {

        $warnings.Add("Not a Git repository: $repo")

        $repositorySummary.Add(
            [pscustomobject]@{
                Name    = $repoName
                Path    = $repo
                Status  = 'Invalid'
                Commits = 0
            }
        )

        continue
    }

    $name = Split-Path -Path $root -Leaf

    # Retrieve commit metadata.
    $gitLogArgs = @(
        'log'
        '--no-show-signature'
        "--since=$since"
        '--format=%H%x1f%an%x1f%ae%x1f%cn%x1f%aI%x1f%cI%x1f%s%x1f%P'
    )

    if ($AllBranches) {
        $gitLogArgs += '--all'
    }
    else {
        $gitLogArgs += 'HEAD'
    }

    $meta = & git -C $repo @gitLogArgs 2>$null

    if ($LASTEXITCODE -ne 0) {

        $warnings.Add("Git log failed: $repo")

        $repositorySummary.Add(
            [pscustomobject]@{
                Name    = $name
                Path    = $repo
                Status  = 'Git error'
                Commits = 0
            }
        )

        continue
    }

    $repoCommitCount = 0

    foreach ($line in $meta) {

        $parts = $line -split [char]31, 8

        if ($parts.Count -ne 8) {
            continue
        }

        $hash, $who, $email, $committer,
        $authored, $committed, $subject, $parents = $parts

        if (
            $Author -and
            "$who $email" -notmatch [regex]::Escape($Author)
        ) {
            continue
        }

        $parentHashes = @(
            $parents -split ' ' | Where-Object { $_ }
        )

        # =====================================================================
        # CHANGED FILES
        # =====================================================================

        $changes = [System.Collections.Generic.List[object]]::new()

        if ($parentHashes.Count -gt 1) {

            # Merge: compare the merged tree against its first parent.
            $raw = & git -C $repo diff `
                --name-status -M `
                $parentHashes[0] $hash 2>$null
        }
        else {

            # Root and normal commits.
            $raw = & git -C $repo diff-tree `
                --root `
                --no-commit-id `
                --name-status `
                -r -M $hash 2>$null
        }

        if ($LASTEXITCODE -ne 0) {

            $warnings.Add(
                "File listing failed for $hash in $name"
            )

            continue
        }

        foreach ($entry in $raw) {

            $fileParts = $entry -split "`t"

            if ($fileParts.Count -lt 2) {
                continue
            }

            $status = $fileParts[0]
            $file = $fileParts[-1]

            $previousPath = if ($fileParts.Count -gt 2) {
                $fileParts[1]
            }
            else {
                ''
            }

            $changes.Add(
                [pscustomobject]@{
                    Status       = $status
                    Path         = $file
                    PreviousPath = $previousPath
                }
            )
        }

        # =====================================================================
        # BRANCH REFERENCES
        # =====================================================================

        # Symbolic references such as origin/HEAD are excluded.
        $branchRefs = & git -C $repo for-each-ref `
            "--contains=$hash" `
            '--format=%(refname:short)%09%(symref)' `
            refs/heads refs/remotes 2>$null

        $branches = @(
            $branchRefs |
                ForEach-Object {

                    $refParts = $_ -split "`t", 2

                    if (
                        $refParts.Count -eq 2 -and
                        $refParts[0] -and
                        -not $refParts[1]
                    ) {
                        $refParts[0]
                    }
                } |
                Sort-Object -Unique
        )

        $branchText = if ($branches.Count -gt 0) {
            $branches -join ', '
        }
        else {
            '(detached/unreferenced)'
        }

        # =====================================================================
        # COMMIT RECORD
        # =====================================================================

        $records.Add(
            [pscustomobject]@{
                Repository = $name
                Branch     = $branchText
                Author     = $who
                AuthorEmail = $email
                Committer  = $committer

                Date = (
                    [datetimeoffset]::Parse($committed).
                        ToLocalTime().
                        ToString('yyyy-MM-dd HH:mm:ss zzz')
                )

                AuthorDate = $authored
                Hash       = $hash
                ShortHash  = $hash.Substring(0, 10)
                Message    = $subject

                Files      = @($changes.ToArray())
                FileCount  = $changes.Count
                Merge      = ($parentHashes.Count -gt 1)
            }
        )

        $repoCommitCount++
    }

    $repositorySummary.Add(
        [pscustomobject]@{
            Name    = $name
            Path    = $repo
            Status  = if ($repoCommitCount -gt 0) {
                'Included'
            }
            else {
                'No matching commits'
            }
            Commits = $repoCommitCount
        }
    )
}

# =============================================================================
# SORT RECORDS
# =============================================================================

$descending = $Direction -eq 'Descending'

$sorted = @(
    $records.ToArray() |
        Sort-Object -Property $SortBy, Date -Descending:$descending
)

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

$base = Join-Path `
    $OutputDirectory `
    "CPRS_Git_Activity_$stamp"

# =============================================================================
# CSV EXPORT
# =============================================================================

$csvPath = $null
$htmlPath = $null

if ($Format -in @('CSV', 'Both')) {

    $csvPath = "$base.csv"

    $flat = @(
        $sorted | ForEach-Object {

            $record = $_

            $fileText = (
                $record.Files | ForEach-Object {
                    "$($_.Status): $($_.Path)"
                }
            ) -join ' | '

            [pscustomobject]@{
                Repository  = $record.Repository
                Branch      = $record.Branch
                Author      = $record.Author
                AuthorEmail = $record.AuthorEmail
                Committer   = $record.Committer
                Date        = $record.Date
                Hash        = $record.Hash
                Message     = $record.Message
                FileCount   = $record.FileCount
                Merge       = $record.Merge
                Files       = $fileText
            }
        }
    )

    if ($flat.Count -gt 0) {

        $flat | Export-Csv `
            -LiteralPath $csvPath `
            -NoTypeInformation `
            -Encoding utf8BOM
    }
    else {

        'Repository,Branch,Author,AuthorEmail,Committer,Date,Hash,Message,FileCount,Merge,Files' |
            Set-Content -LiteralPath $csvPath -Encoding utf8BOM
    }
}

# =============================================================================
# HTML REPORT
# =============================================================================

if ($Format -in @('HTML', 'Both')) {

    $htmlPath = "$base.html"

    $json = ConvertTo-Json `
        -InputObject $sorted `
        -Depth 7 `
        -Compress

    # Prevent embedded data from terminating the HTML script element.
    $json = $json.
        Replace('<', '\u003c').
        Replace('>', '\u003e').
        Replace('&', '\u0026')

    $warningHtml = [System.Net.WebUtility]::HtmlEncode(
        ($warnings -join ' | ')
    )

    $repoHtml = [System.Net.WebUtility]::HtmlEncode(
        (
            $repositorySummary |
                ForEach-Object {
                    "$($_.Name): $($_.Commits) commits ($($_.Status))"
                }
        ) -join ' | '
    )

    $html = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CPRS Git Activity</title>
<style>
:root {
    font-family: Segoe UI,system-ui,sans-serif;
    color: #1c293b;
    background: #f4f6fa;
}
body {
    margin: auto;
    padding: 30px;
    max-width: 1700px;
}
h1 { margin: 0; }
header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 22px;
}
.sub { color: #637084; margin-top: 5px; }
.panel {
    background: white;
    border: 1px solid #e1e5ec;
    border-radius: 12px;
    padding: 18px;
    box-shadow: 0 3px 14px #162d4810;
}
.controls {
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
    margin-bottom: 16px;
}
input,select,button {
    font: inherit;
    padding: 9px 12px;
    border: 1px solid #cbd4df;
    border-radius: 7px;
    background: white;
}
input { min-width: 210px; }
button { cursor: pointer; }
table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}
th {
    text-align: left;
    color: #475569;
    background: #f4f6fa;
    cursor: pointer;
    white-space: nowrap;
}
th,td {
    padding: 12px;
    border-bottom: 1px solid #e7eaf0;
    vertical-align: top;
}
tr:hover { background: #f9fbff; }
.hash {
    font-family: Consolas,monospace;
    color: #325b92;
}
.pill {
    background: #eaf1fa;
    border-radius: 5px;
    padding: 3px 6px;
    display: inline-block;
    max-width: 240px;
    overflow-wrap: anywhere;
}
.files {
    max-width: 420px;
    max-height: 190px;
    overflow: auto;
}
.files div {
    padding: 3px 0;
    overflow-wrap: anywhere;
}
small { color: #68768a; }
.status {
    font-weight: 700;
    margin-right: 6px;
    color: #35699e;
}
.stats { margin-bottom: 15px; color: #4b5d75; }
.warn { color: #a34b22; margin: 10px 0; }
details summary { cursor: pointer; color: #285b99; }
@media(max-width:850px) {
    body { padding: 12px; }
    .panel { overflow: auto; }
}
</style>
</head>
<body>

<header>
    <div>
        <h1>CPRS Git Activity</h1>
        <div class="sub">
            Local repository commit history
            · generated <span id="generated"></span>
        </div>
    </div>
    <button onclick="downloadCSV()">Export filtered CSV</button>
</header>

<section class="panel">
    <div class="controls">
        <input id="q" placeholder="Search author, file, message, hash…">
        <select id="repo"><option value="">All repositories</option></select>
        <select id="author"><option value="">All authors</option></select>
        <select id="branch"><option value="">All branches</option></select>
        <input type="date" id="from" title="From date">
        <input type="date" id="to" title="Through date">
        <select id="sort">
            <option value="Date">Date</option>
            <option value="Author">Author</option>
            <option value="Repository">Repository</option>
            <option value="Branch">Branch</option>
        </select>
        <select id="dir">
            <option value="desc">Descending</option>
            <option value="asc">Ascending</option>
        </select>
        <button onclick="resetFilters()">Reset</button>
    </div>

    <div id="stats" class="stats"></div>
    <div id="repo-summary" class="stats"></div>
    <div id="warnings" class="warn"></div>

    <table>
        <thead>
            <tr>
                <th data-sort="Repository">Repository</th>
                <th data-sort="Branch">Branch(es)</th>
                <th data-sort="Author">Author</th>
                <th data-sort="Date">Commit date</th>
                <th>Commit</th>
                <th>Files changed</th>
            </tr>
        </thead>
        <tbody id="rows"></tbody>
    </table>
</section>

<script id="data" type="application/json">__DATA__</script>

<script>
const data = JSON.parse(document.getElementById('data').textContent);
const $ = id => document.getElementById(id);

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({
    '&':'&amp;',
    '<':'&lt;',
    '>':'&gt;',
    '"':'&quot;',
    "'":'&#39;'
}[c]));

let visible = [];

$('generated').textContent = new Date().toLocaleString();

for (const [id,key] of [
    ['repo','Repository'],
    ['author','Author'],
    ['branch','Branch']
]) {
    const vals = [...new Set(
        data.flatMap(r =>
            id === 'branch'
                ? r.Branch.split(', ')
                : [r[key]]
        )
    )].sort();

    for (const v of vals) {
        const o = document.createElement('option');
        o.value = v;
        o.textContent = v;
        $(id).appendChild(o);
    }
}

function filter() {
    const q = $('q').value.toLowerCase();
    const repo = $('repo').value;
    const author = $('author').value;
    const branch = $('branch').value;
    const from = $('from').value;
    const to = $('to').value;
    const key = $('sort').value;
    const sgn = $('dir').value === 'asc' ? 1 : -1;

    visible = data.filter(r =>
        (!repo || r.Repository === repo) &&
        (!author || r.Author === author) &&
        (!branch || r.Branch.split(', ').includes(branch)) &&
        (!from || r.Date.slice(0,10) >= from) &&
        (!to || r.Date.slice(0,10) <= to) &&
        (!q || [
            r.Repository,
            r.Branch,
            r.Author,
            r.AuthorEmail,
            r.Message,
            r.Hash,
            ...r.Files.map(f => f.Path)
        ].join(' ').toLowerCase().includes(q))
    ).sort((a,b) =>
        sgn * String(a[key]).localeCompare(String(b[key]))
    );

    $('stats').textContent =
        `${visible.length} commits · ` +
        `${new Set(visible.map(r => r.Repository)).size} repositories · ` +
        `${new Set(visible.map(r => r.Author)).size} authors`;

    $('rows').innerHTML = visible.map(r => `
        <tr>
            <td><b>${esc(r.Repository)}</b></td>
            <td><span class="pill">${esc(r.Branch)}</span></td>
            <td>${esc(r.Author)}<br><small>${esc(r.AuthorEmail)}</small></td>
            <td>${esc(r.Date)}</td>
            <td>
                <span class="hash">${esc(r.ShortHash)}</span>
                ${r.Merge ? ' · Merge' : ''}
                <br>${esc(r.Message)}
            </td>
            <td>
                <details>
                    <summary>${r.FileCount} file(s)</summary>
                    <div class="files">
                        ${r.Files.map(f => `
                            <div>
                                <span class="status">${esc(f.Status)}</span>
                                ${esc(
                                    f.PreviousPath
                                        ? f.PreviousPath + ' → ' + f.Path
                                        : f.Path
                                )}
                            </div>
                        `).join('')}
                    </div>
                </details>
            </td>
        </tr>
    `).join('') ||
    '<tr><td colspan="6">No matching commits.</td></tr>';
}

function resetFilters() {
    for (const id of [
        'q','repo','author','branch','from','to'
    ]) {
        $(id).value = '';
    }
    $('sort').value = 'Date';
    $('dir').value = 'desc';
    filter();
}

for (const id of [
    'q','repo','author','branch',
    'from','to','sort','dir'
]) {
    $(id).addEventListener('input', filter);
}

document.querySelectorAll('th[data-sort]').forEach(th => {
    th.onclick = () => {
        $('sort').value = th.dataset.sort;
        filter();
    };
});

function downloadCSV() {
    const cols = [
        'Repository','Branch','Author','Date',
        'Hash','Message','FileCount','Files'
    ];

    const quote = v =>
        '"' + String(v ?? '').replaceAll('"','""') + '"';

    const lines = [
        cols.join(','),
        ...visible.map(r =>
            cols.map(k => quote(
                k === 'Files'
                    ? r.Files.map(f =>
                        f.Status + ': ' + f.Path
                    ).join(' | ')
                    : r[k]
            )).join(',')
        )
    ];

    const blob = new Blob(
        ['\ufeff' + lines.join('\r\n')],
        { type: 'text/csv;charset=utf-8' }
    );

    const a = document.createElement('a');
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = 'CPRS_Git_Filtered.csv';
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

filter();
</script>
</body>
</html>
'@

    $html = $html.Replace('__DATA__', $json)

    $html = $html.Replace(
        '<div id="warnings" class="warn"></div>',
        "<div id=`"warnings`" class=`"warn`">$warningHtml</div>"
    )

    $html = $html.Replace(
        '<div id="repo-summary" class="stats"></div>',
        "<div id=`"repo-summary`" class=`"stats`">Configured repositories: $repoHtml</div>"
    )

    [System.IO.File]::WriteAllText(
        $htmlPath,
        $html,
        [System.Text.UTF8Encoding]::new($false)
    )

    if ($OpenReport) {
        Start-Process -FilePath $htmlPath
    }
}

# =============================================================================
# SUMMARY STATISTICS
# =============================================================================

$finished = Get-Date
$duration = $finished - $started

$uniqueRepositories = @(
    $sorted |
        ForEach-Object { $_.Repository } |
        Sort-Object -Unique
).Count

$contributors = @(
    $sorted |
        ForEach-Object { $_.AuthorEmail } |
        Sort-Object -Unique
).Count

$mergeCommits = @(
    $sorted | Where-Object { $_.Merge }
).Count

$fileChanges = (
    $sorted | Measure-Object -Property FileCount -Sum
).Sum

if ($null -eq $fileChanges) {
    $fileChanges = 0
}

# Count unique file paths independently of the change status.
$uniqueFilePaths = @(
    $sorted |
        ForEach-Object { $_.Files } |
        ForEach-Object { $_.Path } |
        Sort-Object -Unique
).Count

$result = [pscustomobject]@{
    Generated              = $finished
    Days                   = $Days
    Commits                = $sorted.Count
    Repositories           = $uniqueRepositories
    ConfiguredRepositories = $repositorySummary.Count
    Contributors           = $contributors
    MergeCommits           = $mergeCommits
    FileChanges            = $fileChanges
    UniqueFiles            = $uniqueFilePaths
    HTML                   = $htmlPath
    CSV                    = $csvPath
    RepositoryDetails      = @($repositorySummary.ToArray())
    Warnings               = @($warnings.ToArray())
    Duration               = $duration
}

# =============================================================================
# FORMATTED CONSOLE SUMMARY
# =============================================================================

$line = '=' * 74
$separator = '-' * 74

Write-Host ""
Write-Host $line -ForegroundColor Cyan
Write-Host "                       CPRS GIT ACTIVITY REPORT" -ForegroundColor Cyan
Write-Host $line -ForegroundColor Cyan

Write-Host ""
Write-Host ("Generated        : {0}" -f $finished.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Host ("Reporting Period : Last {0} days" -f $Days)
Write-Host ("Branch Scope     : {0}" -f $(if ($AllBranches) {
    'All branches'
} else {
    'Current HEAD'
}))
Write-Host ("Author Filter    : {0}" -f $(if ($Author) {
    $Author
} else {
    'All authors'
}))
Write-Host ("Sort             : {0} ({1})" -f $SortBy, $Direction)
Write-Host ("Output Format    : {0}" -f $Format)

Write-Host ""
Write-Host "SUMMARY" -ForegroundColor Yellow
Write-Host $separator
Write-Host ("Total Commits    : {0}" -f $result.Commits)
Write-Host ("Repositories     : {0} of {1}" -f $result.Repositories, $result.ConfiguredRepositories)
Write-Host ("Contributors     : {0}" -f $result.Contributors)
Write-Host ("Merge Commits    : {0}" -f $result.MergeCommits)
Write-Host ("File Changes     : {0}" -f $result.FileChanges)
Write-Host ("Unique Files     : {0}" -f $result.UniqueFiles)

Write-Host ""
Write-Host "REPOSITORY BREAKDOWN" -ForegroundColor Yellow
Write-Host $separator

Write-Host (
    "{0,-22} {1,8}   {2}" -f
    'Repository', 'Commits', 'Status'
)

Write-Host (
    "{0,-22} {1,8}   {2}" -f
    '----------', '-------', '------'
)

foreach ($repo in $repositorySummary) {

    Write-Host (
        "{0,-22} {1,8}   {2}" -f
        $repo.Name,
        $repo.Commits,
        $repo.Status
    )
}

Write-Host ""
Write-Host "CONTRIBUTORS" -ForegroundColor Yellow
Write-Host $separator

$authorSummary = @(
    $sorted |
        Group-Object AuthorEmail |
        ForEach-Object {
            [pscustomobject]@{
                Author  = $_.Group[0].Author
                Email   = $_.Name
                Commits = $_.Count
            }
        } |
        Sort-Object Commits -Descending
)

if ($authorSummary.Count -eq 0) {
    Write-Host "No matching contributors."
}
else {

    Write-Host (
        "{0,-24} {1,8}" -f 'Author', 'Commits'
    )

    Write-Host (
        "{0,-24} {1,8}" -f '------', '-------'
    )

    foreach ($item in $authorSummary) {
        Write-Host (
            "{0,-24} {1,8}" -f
            $item.Author,
            $item.Commits
        )
    }
}

Write-Host ""
Write-Host "REPORT FILES" -ForegroundColor Yellow
Write-Host $separator

if ($htmlPath) {
    Write-Host ("HTML : {0}" -f $htmlPath) -ForegroundColor Green
}

if ($csvPath) {
    Write-Host ("CSV  : {0}" -f $csvPath) -ForegroundColor Green
}

Write-Host ""
Write-Host "EXECUTION" -ForegroundColor Yellow
Write-Host $separator

Write-Host (
    "Elapsed Time : {0:N2} seconds" -f
    $duration.TotalSeconds
)

Write-Host ""
Write-Host "VALIDATION / WARNINGS" -ForegroundColor Yellow
Write-Host $separator

if ($warnings.Count -eq 0) {
    Write-Host "Warnings : None" -ForegroundColor Green
}
else {

    Write-Host (
        "Warnings : {0}" -f $warnings.Count
    ) -ForegroundColor Yellow

    foreach ($warning in $warnings) {
        Write-Host ("  - {0}" -f $warning) -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host $line -ForegroundColor Cyan

if ($warnings.Count -eq 0) {
    Write-Host "Report completed successfully." -ForegroundColor Green
}
else {
    Write-Host "Report completed with warnings." -ForegroundColor Yellow
}

Write-Host $line -ForegroundColor Cyan
Write-Host ""

# =============================================================================
# OPTIONAL STRUCTURED RESULT
# =============================================================================

if ($PassThru) {

    Write-Host ""
    Write-Host "STRUCTURED REPORT DETAILS" -ForegroundColor Cyan
    Write-Host ('-' * 74)

    $result |
        Select-Object -ExcludeProperty RepositoryDetails, Warnings |
        Format-List |
        Out-Host

    Write-Host "REPOSITORY DETAILS" -ForegroundColor Yellow
    Write-Host ('-' * 74)

    $result.RepositoryDetails |
        Format-Table Name, Path, Status, Commits -AutoSize |
        Out-Host

    Write-Host "WARNINGS" -ForegroundColor Yellow
    Write-Host ('-' * 74)

    if ($result.Warnings.Count -eq 0) {
        Write-Host "None" -ForegroundColor Green
    }
    else {
        $result.Warnings | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Yellow
        }
    }

    Write-Host ""
}