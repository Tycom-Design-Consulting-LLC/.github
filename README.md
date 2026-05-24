# .github -- Tycom Design & Consulting LLC

Org-level GitHub configuration for all Tycom repositories.

## What lives here

- **Reusable workflows** (`.github/workflows/`) -- callable workflows that individual repos reference instead of duplicating CI/CD logic
- **Deploy scripts** (`.github/scripts/`) -- Python scripts invoked by the reusable workflows
- **Documentation** (`.github/docs/`) -- schema docs and examples for repo-level configuration

## Reusable workflows

### `sharepoint-deploy.yml`

Mirrors repo files to SharePoint via Microsoft Graph API. Called by each repo's own `deploy.yml` workflow.

**How it works:**
1. The calling repo pushes to `main`
2. The calling repo's workflow invokes this reusable workflow
3. This workflow reads `.tycom-deploy.yml` from the calling repo's root
4. It authenticates to Graph API using repo-level secrets
5. It diffs local files against SharePoint and syncs changes (add/update/delete)
6. It posts a summary to the GitHub Actions job log

**Required secrets in the calling repo:**
- `GRAPH_TENANT_ID`
- `GRAPH_CLIENT_ID`
- `GRAPH_CLIENT_SECRET`

## Adding a new repo to the deploy pipeline

1. Copy the example workflow from `.github/docs/example-calling-workflow.yml` into your repo at `.github/workflows/deploy.yml`
2. Create `.tycom-deploy.yml` in your repo root (see `.github/docs/deploy-config-schema.md` for format)
3. Add the three Graph API secrets to your repo's Settings > Secrets > Actions
4. Push to `main` -- the deploy workflow triggers automatically
