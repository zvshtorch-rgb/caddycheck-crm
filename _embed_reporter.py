"""
Update the embedded job_reporter.py here-string in deploy_reporter.ps1.
Handles the case where the file is already embedded (replaces existing here-string).
"""

with open("job_reporter.py", encoding="utf-8") as f:
    py_content = f.read()

assert "'@" not in py_content, "job_reporter.py contains '@' which breaks PS here-string"
assert all(ord(c) <= 127 for c in py_content), "job_reporter.py has non-ASCII chars"

with open("deploy_reporter.ps1", encoding="utf-8") as f:
    ps_content = f.read()

OPEN_MARKER = "$jobReporterContent = @'\n"
CLOSE_MARKER = "\n'@"

start = ps_content.find(OPEN_MARKER)
if start == -1:
    print("ERROR: open marker not found")
    exit(1)

content_start = start + len(OPEN_MARKER)
end = ps_content.find(CLOSE_MARKER, content_start)
if end == -1:
    print("ERROR: close marker not found")
    exit(1)

new_ps = (
    ps_content[:content_start]
    + py_content
    + ps_content[end:]
)

with open("deploy_reporter.ps1", "w", encoding="utf-8", newline="\r\n") as f:
    f.write(new_ps)

print("Done. job_reporter.py re-embedded into deploy_reporter.ps1")
