import json
import sys
from datetime import datetime


def bool_line(enabled, label):
    return f"- {'Yes' if enabled else 'No'}: {label}"


def build_session_block(sessions, start_fresh):
    if start_fresh:
        return (
            "### Clash Session Strategy\n"
            "- Current cycle starts fresh: existing clash detection sessions are deleted from beginning.\n"
            "- Regeneration allowed: default session templates can be re-created with one action in UI.\n"
        )

    kept = [s for s in sessions if s.get("Keep")]
    if not kept:
        return (
            "### Clash Session Strategy\n"
            "- No sessions marked to keep.\n"
            "- Use `Generate Back Missing Sessions` in UI to restore default sessions.\n"
        )

    lines = ["### Clash Session Strategy", "- Keep sessions:"]
    for s in kept:
        lines.append(f"  - {s.get('Name', 'Unnamed')} ({s.get('DisciplinePair', 'Unassigned')})")
    return "\n".join(lines) + "\n"


def main():
    raw = sys.stdin.read()
    payload = json.loads(raw) if raw.strip() else {}

    project_name = payload.get("ProjectName") or "[Project Name]"
    package_method = payload.get("PackageMethod") or "[Not Selected]"

    lines = [
        f"## BIM Execution Plan Input Summary - {project_name}",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "### Project Information",
        f"- Project Number: {payload.get('ProjectNumber') or '[Not set]'}",
        f"- Project Name: {project_name}",
        f"- Project Address: {payload.get('ProjectAddress') or '[Not set]'}",
        f"- Project Owner/Client: {payload.get('Client') or '[Not set]'}",
        f"- Project Type: {payload.get('ProjectType') or '[Not set]'}",
        f"- Contract Type: {payload.get('ContractType') or '[Not set]'}",
        f"- Project Description: {payload.get('ProjectDescription') or '[Not set]'}",
        f"- BIM Lead: {payload.get('BimLead') or '[Not set]'}",
        f"- Coordination Meeting Cadence: {payload.get('CoordinationMeetingCadence') or '[Not set]'}",
        "",
        "### ACC Collaboration Method",
        f"- Primary Method: {package_method}",
        f"- Auto-Publish Cadence: {payload.get('AutoPublishCadence') or '[Not set]'}",
        f"- Package Sharing Timeline: {payload.get('SharingFrequency') or '[Not set]'}",
        f"- Package Naming Convention: {payload.get('PackageNamingConvention') or '[Not set]'}",
        "",
        "### Geo-Referencing",
        f"- Geocoordinate System: {payload.get('GeoCoordinateSystem') or '[Not set]'}",
        f"- Coordinates Acquired From Model: {payload.get('AcquireCoordinatesFromModel') or '[Not set]'}",
        "",
        "### Approved Software Versions",
        f"- Autodesk Revit: {payload.get('RevitVersion') or '[Not set]'}",
        f"- Autodesk AutoCAD: {payload.get('AutoCadVersion') or '[Not set]'}",
        f"- Autodesk Civil 3D: {payload.get('Civil3DVersion') or '[Not set]'}",
        f"- Autodesk Desktop Connector: {payload.get('DesktopConnectorVersion') or '[Not set]'}",
        f"- Bluebeam Revu: {payload.get('BluebeamVersion') or '[Not set]'}",
        "",
        build_session_block(payload.get("Sessions", []), payload.get("StartFresh", False)),
        "### Recommended Views for ACC Publishing",
        "- 3D_Coordination (local geometry only)",
        "- 3D_Coordination_Links (discipline links loaded)",
        "",
        "### Notes",
        "- This output is generated from your form and can be pasted into Notebook B / Notebook D / Appendix E sections.",
        "- Keep using discipline package naming patterns like '<Project>_Shared for <Purpose>'.",
    ]

    sys.stdout.write("\n".join(lines).strip() + "\n")


if __name__ == "__main__":
    main()
