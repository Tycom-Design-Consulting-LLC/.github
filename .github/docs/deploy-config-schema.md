# .tycom-deploy.yml Configuration Schema

This document defines the format for `.tycom-deploy.yml`, the per-repo configuration file that controls how files are deployed from Git to SharePoint (and eventually other targets).

## Location

Place `.tycom-deploy.yml` in the root of your repository.

## Top-level fields

```yaml
version: 1                  # Schema version (always 1 for now)
repo: pyt-tools              # Repository name (for logging)
agent: dev-tools             # Owning agent name (for logging)

deploy:                      # List of deployment targets
  - target: sharepoint
    # ... target-specific fields
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `version` | integer | Yes | Schema version. Currently `1`. |
| `repo` | string | Yes | Repository name, used in logging and summaries. |
| `agent` | string | No | Name of the agent that owns this repo. Informational only. |
| `deploy` | list | Yes | List of deployment target objects. Each one is processed independently. |

## Deployment target: `sharepoint`

Mirrors files from the repo to a SharePoint document library via Microsoft Graph API.

```yaml
deploy:
  - target: sharepoint
    site_library: "Tycom Internal/Agents"
    remote_path: "dev-tools-agent/pyt-tools"
    source_path: "."
    sync_mode: mirror
    include:
      - "**/*.pyt"
      - "**/*.py"
    exclude:
      - "__pycache__/**"
      - "*.pyc"
```

| Field | Type | Required | Default | Description |
|-------|------|----------|---------|-------------|
| `target` | string | Yes | -- | Must be `sharepoint`. |
| `site_library` | string | Yes | -- | Format: `"Site Name/Library Name"`. The script resolves the site from the known SharePoint hostname and looks up the named document library. |
| `remote_path` | string | Yes | -- | Folder path within the document library where files will be synced. Created automatically if it doesn't exist. |
| `source_path` | string | No | `"."` | Local directory (relative to repo root) to scan for files. |
| `sync_mode` | string | No | `"mirror"` | Either `mirror` or `additive`. See below. |
| `include` | list of strings | No | `[]` (all files) | Glob patterns for files to include. If empty, all files in source_path are included. |
| `exclude` | list of strings | No | `[]` | Glob patterns for files to exclude. Excludes take priority over includes. |

### Sync modes

- **`mirror`** (default): The remote path is an exact mirror of the local source. Files that exist remotely but not locally are deleted. This ensures SharePoint always matches the Git state exactly.

- **`additive`**: Files are added and updated but never deleted from SharePoint. Use this when the remote folder contains files from other sources that should be preserved.

### Include/exclude pattern syntax

Patterns use Python `fnmatch` glob syntax:

| Pattern | Matches |
|---------|---------|
| `*.py` | Any `.py` file in the source root |
| `**/*.py` | Any `.py` file at any depth |
| `tools/**` | Everything under the `tools/` directory |
| `__pycache__/**` | Everything under any `__pycache__/` directory |
| `.git/` | The `.git` directory and all its contents |

**Precedence:** Excludes are checked first. If a file matches any exclude pattern, it is skipped regardless of include patterns.

## Multiple deployment targets

A single repo can deploy to multiple SharePoint destinations (or a mix of SharePoint and future targets like S3 or Lambda):

```yaml
deploy:
  # Sync standards JSONs to the standards folder
  - target: sharepoint
    site_library: "Tycom Internal/Agents"
    remote_path: "gdb-standards-agent/standards"
    source_path: "standards"
    include:
      - "master-*.json"
      - "label-engine-standards.json"

  # Sync scripts to a different folder
  - target: sharepoint
    site_library: "Tycom Internal/Agents"
    remote_path: "gdb-standards-agent/scripts"
    source_path: "scripts"
    include:
      - "*.py"
```

Each target is processed independently. Failures in one target do not prevent others from running (though any failure causes the overall job to exit non-zero).

## Future targets

The schema is designed to accommodate additional targets beyond SharePoint. These are not yet implemented but the `.tycom-deploy.yml` format will support them:

- **`s3`**: Sync to an S3 bucket (for static website hosting)
- **`lambda`**: Package and deploy an AWS Lambda function
- **`supabase`**: Transform JSON and upload to Supabase tables

When these are implemented, they will be handled by separate scripts or additional steps in the reusable workflow. The SharePoint deploy script ignores non-SharePoint targets gracefully.

## Complete example: pyt-tools

```yaml
version: 1
repo: pyt-tools
agent: dev-tools

deploy:
  - target: sharepoint
    site_library: "Tycom Internal/Agents"
    remote_path: "dev-tools-agent/pyt-tools"
    source_path: "."
    sync_mode: mirror
    include:
      - "tycom-standard-tools/**/*.pyt"
      - "tycom-standard-tools/**/*.pyt.xml"
      - "tycom-standard-tools/**/*.py"
      - "modular-tools/**/*.pyt"
      - "modular-tools/**/*.pyt.xml"
      - "modular-tools/**/*.py"
      - "client-specific-tools/**/*.pyt"
      - "client-specific-tools/**/*.pyt.xml"
      - "client-specific-tools/**/*.py"
    exclude:
      - "__pycache__/**"
      - "*.pyc"
      - ".agent/**"
      - "tests/**"
      - ".git/**"
      - ".github/**"
      - "CLAUDE.md"
      - ".tycom-deploy.yml"
      - ".gitignore"
```

## Complete example: gdb-standards (multi-target)

```yaml
version: 1
repo: gdb-standards
agent: gdb-standards

deploy:
  - target: sharepoint
    site_library: "Tycom Internal/Agents"
    remote_path: "gdb-standards-agent/standards"
    source_path: "standards"
    sync_mode: mirror
    include:
      - "master-*.json"
      - "label-engine-standards.json"
      - "templates/**"
    exclude:
      - "projects/**"
      - "benchmarks/**"

  # Future: push schema data to Supabase for API access
  # - target: supabase
  #   source_path: "standards/master-schema-standards.json"
  #   transform: "json-to-tables"
```
