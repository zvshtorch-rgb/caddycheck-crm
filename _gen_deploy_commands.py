import os, sys
sys.path.insert(0, ".")
os.environ["SUPABASE_URL"] = "https://rdoxihpmghrvroddnkdi.supabase.co"
os.environ["SUPABASE_SERVICE_ROLE_KEY"] = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"
    ".eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJkb3hpaHBtZ2hydnJvZGRua2RpIiwicm9sZSI6"
    "InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3MzY0NTExNywiZXhwIjoyMDg5MjIxMTE3fQ"
    ".umgghE4z-ClVQ0KY8LQJhJtbG2tYlVh0fY0d9JnYXBA"
)

from services.supabase_service import _get_client

client = _get_client()
rows = client.table("projects").select("project_name").order("project_name").execute().data
names = sorted(set(r["project_name"].strip() for r in rows if r.get("project_name", "").strip()))

url = "https://raw.githubusercontent.com/zvshtorch-rgb/caddycheck-crm/main/deploy_reporter.ps1"

# Force TLS 1.2 so the download works on older Windows PCs (pre-2016)
tls_fix = "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"

lines = []
for name in names:
    cmd = (
        tls_fix + "; "
        + 'Invoke-WebRequest -Uri "' + url + '?nocache=$(Get-Random)"'
        + " -OutFile 'C:\\deploy.ps1' -UseBasicParsing;"
        + ' powershell -ExecutionPolicy Bypass -File \'C:\\deploy.ps1\' -ProjectName "' + name + '";'
        + " & 'C:\\CaddyCheck\\run_reporter.bat'"
    )
    lines.append("# " + name)
    lines.append(cmd)
    lines.append("")

output = "\n".join(lines)
with open("deploy_all_projects.txt", "w", encoding="utf-8") as f:
    f.write(output)

print(f"Written {len(names)} projects to deploy_all_projects.txt")
