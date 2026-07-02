"""
Embed job_reporter.py content into deploy_reporter.ps1 as a PowerShell here-string.
Replaces the Invoke-WebRequest download step so no second GitHub download is needed.
This fixes TLS issues on old Windows PCs.
"""
import re

# Read both files
with open("job_reporter.py", encoding="utf-8") as f:
    py_content = f.read()

with open("deploy_reporter.ps1", encoding="utf-8") as f:
    ps_content = f.read()

# Verify no chars that would break PS single-quote here-string
assert "'@" not in py_content, "job_reporter.py contains '@' which would break PS here-string"
assert all(ord(c) <= 127 for c in py_content), "job_reporter.py has non-ASCII chars"

# Build the replacement block
embedded_block = (
    "# -- Step 3: Install directory + job_reporter.py (embedded - no download needed) --\n"
    "Write-Step \"Creating $INSTALL_DIR ...\"\n"
    "New-Item -ItemType Directory -Force -Path $INSTALL_DIR | Out-Null\n"
    "\n"
    "Write-Step \"Writing job_reporter.py (embedded)...\"\n"
    "$jobReporterContent = @'\n"
    + py_content
    + "\n'@\n"
    "[System.IO.File]::WriteAllText(\"$INSTALL_DIR\\job_reporter.py\", $jobReporterContent, [System.Text.Encoding]::ASCII)\n"
    "Write-Host \"job_reporter.py written.\" -ForegroundColor Green"
)

# Replace the old download block
old_block = (
    "# -- Step 3: Install directory + job_reporter.py ------------------------------\n"
    "Write-Step \"Creating $INSTALL_DIR ...\"\n"
    "New-Item -ItemType Directory -Force -Path $INSTALL_DIR | Out-Null\n"
    "\n"
    "Write-Step \"Downloading job_reporter.py from GitHub...\"\n"
    "Invoke-WebRequest -Uri $GITHUB_RAW -OutFile \"$INSTALL_DIR\\job_reporter.py\" -UseBasicParsing\n"
    "Write-Host \"job_reporter.py downloaded.\" -ForegroundColor Green"
)

if old_block not in ps_content:
    print("ERROR: old block not found in deploy_reporter.ps1 - check whitespace/line endings")
    # Show context for debugging
    idx = ps_content.find("job_reporter.py downloaded")
    print(repr(ps_content[max(0,idx-300):idx+50]))
else:
    new_content = ps_content.replace(old_block, embedded_block, 1)
    # Also remove the now-unused $GITHUB_RAW variable line
    new_content = re.sub(r'\$GITHUB_RAW\s+=\s+"[^"]+"\s*\n', '', new_content)
    with open("deploy_reporter.ps1", "w", encoding="utf-8", newline="\r\n") as f:
        f.write(new_content)
    print("Done. deploy_reporter.ps1 updated with embedded job_reporter.py")
