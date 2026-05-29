"""
SharePoint Deploy Script for Tycom Repos
=========================================
Reads .tycom-deploy.yml from the calling repo, authenticates to Graph API,
and mirrors files to SharePoint. Supports add, update, and delete operations
with mirror and additive sync modes.

Called by the reusable workflow at .github/workflows/sharepoint-deploy.yml.

Environment variables required:
    GRAPH_TENANT_ID
    GRAPH_CLIENT_ID
    GRAPH_CLIENT_SECRET
"""

import fnmatch
import hashlib
import json
import os
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

import msal
import requests
import yaml


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SITE_HOSTNAME = "tycomdc.sharepoint.com"
SITE_PATH = "/sites/TycomInternal"

# Upload size limit for simple PUT (4 MB). Files larger than this would need
# a resumable upload session, but Tycom repos contain no files that large.
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024


# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------

def get_access_token() -> str:
    """Acquire an OAuth2 token using client credentials flow via MSAL."""
    tenant_id = os.environ.get("GRAPH_TENANT_ID")
    client_id = os.environ.get("GRAPH_CLIENT_ID")
    client_secret = os.environ.get("GRAPH_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        print("::error::Missing one or more Graph API environment variables "
              "(GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET)")
        sys.exit(1)

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "unknown"))
        print(f"::error::Token acquisition failed: {error}")
        sys.exit(1)

    return result["access_token"]


# ---------------------------------------------------------------------------
# Graph API helpers
# ---------------------------------------------------------------------------

class GraphClient:
    """Thin wrapper around Microsoft Graph API for SharePoint file operations."""

    def __init__(self, token: str):
        self.session = requests.Session()
        self.session.headers["Authorization"] = f"Bearer {token}"
        self._drive_id = None
        self._site_id = None

    @property
    def site_id(self) -> str:
        if self._site_id is None:
            url = f"{GRAPH_BASE}/sites/{SITE_HOSTNAME}:{SITE_PATH}"
            resp = self.session.get(url)
            resp.raise_for_status()
            self._site_id = resp.json()["id"]
        return self._site_id

    def get_drive_id(self, library_name: str) -> str:
        """Resolve a document library name to its drive ID."""
        url = f"{GRAPH_BASE}/sites/{self.site_id}/drives"
        resp = self.session.get(url)
        resp.raise_for_status()
        for drive in resp.json().get("value", []):
            if drive["name"] == library_name:
                return drive["id"]
        raise ValueError(
            f"Document library '{library_name}' not found on site. "
            f"Available: {[d['name'] for d in resp.json().get('value', [])]}"
        )

    def list_remote_files(self, drive_id: str, remote_path: str) -> dict:
        """
        Recursively list all files under remote_path.
        Returns {relative_path: {"id": item_id, "sha256": hash, "size": n}}.
        """
        files = {}
        self._walk_remote(drive_id, remote_path, "", files)
        return files

    def _walk_remote(self, drive_id: str, base_path: str, rel_prefix: str,
                     files: dict):
        """Recursively walk a SharePoint folder."""
        if base_path:
            url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{base_path}:/children"
        else:
            url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"

        while url:
            resp = self.session.get(url)
            if resp.status_code == 404:
                # Folder doesn't exist yet -- that's fine, no remote files
                return
            resp.raise_for_status()
            data = resp.json()

            for item in data.get("value", []):
                name = item["name"]
                rel = f"{rel_prefix}{name}" if not rel_prefix else f"{rel_prefix}/{name}"

                if "folder" in item:
                    child_path = f"{base_path}/{name}" if base_path else name
                    self._walk_remote(drive_id, child_path, rel, files)
                elif "file" in item:
                    sha256 = (item.get("file", {})
                              .get("hashes", {})
                              .get("sha256Hash", ""))
                    files[rel] = {
                        "id": item["id"],
                        "sha256": sha256.lower() if sha256 else "",
                        "size": item.get("size", 0),
                    }

            url = data.get("@odata.nextLink")

    def upload_file(self, drive_id: str, remote_path: str, local_path: Path):
        """Upload a file using simple PUT (< 4 MB)."""
        file_size = local_path.stat().st_size
        if file_size > SIMPLE_UPLOAD_LIMIT:
            return self._upload_large_file(drive_id, remote_path, local_path)

        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}:/content"
        with open(local_path, "rb") as f:
            data = f.read()

        resp = self.session.put(
            url,
            data=data,
            headers={"Content-Type": "application/octet-stream"},
        )
        resp.raise_for_status()
        return resp.json()

    def _upload_large_file(self, drive_id: str, remote_path: str,
                           local_path: Path):
        """Upload using resumable upload session for files > 4 MB."""
        url = (f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}"
               f":/createUploadSession")
        resp = self.session.post(url, json={
            "item": {"@microsoft.graph.conflictBehavior": "replace"}
        })
        resp.raise_for_status()
        upload_url = resp.json()["uploadUrl"]

        file_size = local_path.stat().st_size
        chunk_size = 10 * 1024 * 1024  # 10 MB chunks

        with open(local_path, "rb") as f:
            offset = 0
            while offset < file_size:
                chunk = f.read(chunk_size)
                end = offset + len(chunk) - 1
                headers = {
                    "Content-Range": f"bytes {offset}-{end}/{file_size}",
                    "Content-Length": str(len(chunk)),
                }
                resp = requests.put(upload_url, data=chunk, headers=headers)
                resp.raise_for_status()
                offset += len(chunk)

        return resp.json()

    def delete_item(self, drive_id: str, item_id: str):
        """Delete a file or folder by item ID."""
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
        resp = self.session.delete(url)
        # 204 No Content is success
        if resp.status_code not in (200, 204):
            resp.raise_for_status()


# ---------------------------------------------------------------------------
# Local file scanning
# ---------------------------------------------------------------------------

def compute_sha256(path: Path) -> str:
    """Compute SHA256 hash of a local file."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def scan_local_files(source_dir: Path, includes: list, excludes: list) -> dict:
    """
    Walk source_dir and return files matching include/exclude patterns.
    Returns {relative_path: {"local_path": Path, "sha256": hash, "size": n}}.
    """
    files = {}

    for root, dirs, filenames in os.walk(source_dir):
        # Skip .git directory
        if ".git" in dirs:
            dirs.remove(".git")
        if ".github" in dirs:
            dirs.remove(".github")

        for name in filenames:
            full = Path(root) / name
            rel = str(full.relative_to(source_dir)).replace("\\", "/")

            # Check excludes first (takes priority)
            if _matches_any(rel, excludes):
                continue

            # If includes are specified, file must match at least one
            if includes and not _matches_any(rel, includes):
                continue

            files[rel] = {
                "local_path": full,
                "sha256": compute_sha256(full),
                "size": full.stat().st_size,
            }

    return files


def _matches_any(path: str, patterns: list) -> bool:
    """Check if a path matches any of the glob patterns."""
    for pattern in patterns:
        if fnmatch.fnmatch(path, pattern):
            return True
        # Also check if any parent directory matches (for patterns like __pycache__/**)
        if pattern.endswith("/**") or pattern.endswith("/**/*"):
            dir_pattern = pattern.rstrip("/*")
            parts = path.split("/")
            for i in range(len(parts)):
                if fnmatch.fnmatch(parts[i], dir_pattern):
                    return True
        # Check directory-style patterns like ".git/"
        if pattern.endswith("/"):
            if path.startswith(pattern) or f"/{pattern}" in f"/{path}":
                return True
    return False


# ---------------------------------------------------------------------------
# Diff and sync
# ---------------------------------------------------------------------------

def compute_diff(local_files: dict, remote_files: dict, sync_mode: str):
    """
    Compare local and remote file sets.
    Returns (to_add, to_update, to_delete, unchanged).
    Each is a list of relative paths.
    """
    local_set = set(local_files.keys())
    remote_set = set(remote_files.keys())

    to_add = sorted(local_set - remote_set)

    to_update = []
    unchanged = []
    for rel in sorted(local_set & remote_set):
        local_hash = local_files[rel]["sha256"]
        remote_hash = remote_files[rel]["sha256"]
        if local_hash != remote_hash:
            to_update.append(rel)
        else:
            unchanged.append(rel)

    if sync_mode == "mirror":
        to_delete = sorted(remote_set - local_set)
    else:
        # additive mode: never delete remote files
        to_delete = []

    return to_add, to_update, to_delete, unchanged


def _upload_manifest(client: GraphClient, drive_id: str, remote_path: str,
                     repo_name: str, local_files: dict):
    """Upload a _manifest.json to the deploy root for staleness verification.

    Reads git metadata from GitHub Actions environment variables. If those are
    not available (e.g., running locally), falls back to sensible defaults.
    """
    manifest = {
        "repo": repo_name,
        "branch": os.environ.get("GITHUB_REF_NAME", "unknown"),
        "commit_sha": os.environ.get("GITHUB_SHA", "unknown"),
        "deployed_at": datetime.now(timezone.utc).isoformat(),
        "deployed_by": "github-actions",
        "run_id": os.environ.get("GITHUB_RUN_ID", ""),
        "file_count": len(local_files),
    }

    # Write to a temp file and upload
    manifest_bytes = json.dumps(manifest, indent=2).encode("utf-8")
    with tempfile.NamedTemporaryFile(
        suffix=".json", delete=False, mode="wb"
    ) as tmp:
        tmp.write(manifest_bytes)
        tmp_path = Path(tmp.name)

    try:
        dest = f"{remote_path}/_manifest.json" if remote_path else "_manifest.json"
        client.upload_file(drive_id, dest, tmp_path)
        print(f"    M _manifest.json (staleness manifest)")
    except Exception as e:
        # Manifest upload failure is non-fatal -- log and continue
        print(f"    ! Failed to upload _manifest.json: {e}")
    finally:
        tmp_path.unlink(missing_ok=True)


def sync_deployment(client: GraphClient, deployment: dict, repo_root: Path,
                    repo_name: str = "unknown"):
    """Execute a single deployment target from .tycom-deploy.yml."""
    target = deployment.get("target", "sharepoint")
    if target != "sharepoint":
        print(f"  Skipping non-SharePoint target: {target}")
        return None

    # Resolve library and paths
    site_library = deployment.get("site_library", "")
    if "/" in site_library:
        # Format: "Tycom Internal/Agents" means site "Tycom Internal",
        # library "Agents"
        _, library_name = site_library.rsplit("/", 1)
    else:
        library_name = site_library

    remote_path = deployment.get("remote_path", "")
    source_path = deployment.get("source_path", ".")
    includes = deployment.get("include", [])
    excludes = deployment.get("exclude", [])
    sync_mode = deployment.get("sync_mode", "mirror")

    source_dir = repo_root / source_path
    if not source_dir.exists():
        print(f"::error::Source directory does not exist: {source_dir}")
        return None

    print(f"  Library: {library_name}")
    print(f"  Remote path: {remote_path}")
    print(f"  Source: {source_dir}")
    print(f"  Sync mode: {sync_mode}")
    print(f"  Include patterns: {len(includes)}")
    print(f"  Exclude patterns: {len(excludes)}")

    # Resolve drive ID
    drive_id = client.get_drive_id(library_name)
    print(f"  Drive ID: {drive_id[:20]}...")

    # Scan local files
    print("  Scanning local files...")
    local_files = scan_local_files(source_dir, includes, excludes)
    print(f"  Local files: {len(local_files)}")

    # Scan remote files
    print("  Scanning remote files...")
    remote_files = client.list_remote_files(drive_id, remote_path)
    print(f"  Remote files: {len(remote_files)}")

    # Compute diff
    to_add, to_update, to_delete, unchanged = compute_diff(
        local_files, remote_files, sync_mode
    )

    print(f"  Diff: +{len(to_add)} add, ~{len(to_update)} update, "
          f"-{len(to_delete)} delete, ={len(unchanged)} unchanged")

    errors = []

    # Upload new files
    for rel in to_add:
        dest = f"{remote_path}/{rel}" if remote_path else rel
        try:
            client.upload_file(drive_id, dest, local_files[rel]["local_path"])
            print(f"    + {rel}")
        except Exception as e:
            msg = f"Failed to upload {rel}: {e}"
            print(f"    ! {msg}")
            errors.append(msg)

    # Update changed files
    for rel in to_update:
        dest = f"{remote_path}/{rel}" if remote_path else rel
        try:
            client.upload_file(drive_id, dest, local_files[rel]["local_path"])
            print(f"    ~ {rel}")
        except Exception as e:
            msg = f"Failed to update {rel}: {e}"
            print(f"    ! {msg}")
            errors.append(msg)

    # Delete removed files (mirror mode only)
    for rel in to_delete:
        item_id = remote_files[rel]["id"]
        try:
            client.delete_item(drive_id, item_id)
            print(f"    - {rel}")
        except Exception as e:
            msg = f"Failed to delete {rel}: {e}"
            print(f"    ! {msg}")
            errors.append(msg)

    # Upload deploy manifest for staleness verification
    _upload_manifest(client, drive_id, remote_path, repo_name, local_files)

    return {
        "added": len(to_add),
        "updated": len(to_update),
        "deleted": len(to_delete),
        "unchanged": len(unchanged),
        "errors": errors,
    }


# ---------------------------------------------------------------------------
# Job summary
# ---------------------------------------------------------------------------

def write_job_summary(results: list):
    """Write a markdown summary to $GITHUB_STEP_SUMMARY."""
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not summary_path:
        print("(GITHUB_STEP_SUMMARY not set -- printing summary to stdout)")
        summary_path = None

    lines = ["## SharePoint Deploy Summary", ""]

    total_added = 0
    total_updated = 0
    total_deleted = 0
    total_unchanged = 0
    total_errors = 0

    for i, r in enumerate(results):
        if r is None:
            continue
        lines.append(f"### Deployment {i + 1}")
        lines.append("")
        lines.append(f"| Metric | Count |")
        lines.append(f"|--------|-------|")
        lines.append(f"| Added | {r['added']} |")
        lines.append(f"| Updated | {r['updated']} |")
        lines.append(f"| Deleted | {r['deleted']} |")
        lines.append(f"| Unchanged | {r['unchanged']} |")
        lines.append(f"| Errors | {len(r['errors'])} |")
        lines.append("")

        if r["errors"]:
            lines.append("**Errors:**")
            for err in r["errors"]:
                lines.append(f"- {err}")
            lines.append("")

        total_added += r["added"]
        total_updated += r["updated"]
        total_deleted += r["deleted"]
        total_unchanged += r["unchanged"]
        total_errors += len(r["errors"])

    if len(results) > 1:
        lines.append("### Totals")
        lines.append("")
        lines.append(f"| Metric | Count |")
        lines.append(f"|--------|-------|")
        lines.append(f"| Added | {total_added} |")
        lines.append(f"| Updated | {total_updated} |")
        lines.append(f"| Deleted | {total_deleted} |")
        lines.append(f"| Unchanged | {total_unchanged} |")
        lines.append(f"| Errors | {total_errors} |")
        lines.append("")

    summary = "\n".join(lines)
    print(summary)

    if summary_path:
        with open(summary_path, "a") as f:
            f.write(summary)

    return total_errors


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("SharePoint Deploy")
    print("=" * 60)
    print()

    # Find the repo root (where .tycom-deploy.yml lives)
    # In GitHub Actions, the checkout is at GITHUB_WORKSPACE
    workspace = os.environ.get("GITHUB_WORKSPACE", ".")
    repo_root = Path(workspace).resolve()

    config_path = repo_root / ".tycom-deploy.yml"
    if not config_path.exists():
        print(f"::error::Config file not found: {config_path}")
        sys.exit(1)

    # Parse config
    with open(config_path) as f:
        config = yaml.safe_load(f)

    repo_name = config.get("repo", "unknown")
    deployments = config.get("deploy", [])

    print(f"Repo: {repo_name}")
    print(f"Deployments: {len(deployments)}")
    print()

    if not deployments:
        print("No deployments configured. Nothing to do.")
        sys.exit(0)

    # Authenticate
    print("Authenticating with Graph API...")
    token = get_access_token()
    client = GraphClient(token)
    print("Authenticated.")
    print()

    # Process each deployment
    results = []
    for i, deployment in enumerate(deployments):
        target = deployment.get("target", "sharepoint")
        print(f"--- Deployment {i + 1}: {target} ---")

        if target != "sharepoint":
            print(f"  Skipping (only SharePoint targets are handled by this script)")
            results.append(None)
            continue

        result = sync_deployment(client, deployment, repo_root, repo_name)
        results.append(result)
        print()

    # Write summary
    print("=" * 60)
    error_count = write_job_summary(results)

    if error_count > 0:
        print(f"\n::error::{error_count} error(s) during deployment")
        sys.exit(1)
    else:
        print("\nDeploy completed successfully.")


if __name__ == "__main__":
    main()
