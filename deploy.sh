#!/usr/bin/env bash
# ==============================================================================
# Auto Deployment Script for Google Apps Script (Tirwin Email List Validator)
#
# Secret Resolution Order:
#   1. Local `.env` file (if present in the project root)
#   2. Git Secrets / CI Environment Variables (injected via GitHub Actions `secrets.*`,
#      `git secret reveal`, or local `git config` fallback)
#   3. CLI flags (`--script-id`, `--sheet-id`, `--deployment-id`, `--desc`)
#
# Also normalizes CLASPRC_JSON across both @google/clasp v2 (`rc.token`) and
# @google/clasp v3 (`rc.tokens.default`) so it never fails with:
#   "TypeError: Cannot read properties of undefined (reading 'access_token')"
# ==============================================================================

set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$PROJECT_DIR"

PROJECT_TITLE="Tirwin Email List Validator"
DEPLOY_DESC="Auto-deployment $(date '+%Y-%m-%d %H:%M:%S')"
CLI_SCRIPT_ID=""
CLI_SHEET_ID=""
CLI_DEPLOYMENT_ID=""

# Colors for terminal output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

log_info()  { echo -e "${BLUE}[INFO]${NC} $1"; }
log_ok()    { echo -e "${GREEN}[OK]${NC} $1"; }
log_warn()  { echo -e "${YELLOW}[WARN]${NC} $1"; }
log_error() { echo -e "${RED}[ERROR]${NC} $1"; }

# Parse CLI arguments
while [[ $# -gt 0 ]]; do
  case "$1" in
    --script-id)
      CLI_SCRIPT_ID="$2"
      shift 2
      ;;
    --sheet-id)
      CLI_SHEET_ID="$2"
      shift 2
      ;;
    --deployment-id)
      CLI_DEPLOYMENT_ID="$2"
      shift 2
      ;;
    --desc)
      DEPLOY_DESC="$2"
      shift 2
      ;;
    -h|--help)
      echo "Usage: ./deploy.sh [--script-id <SCRIPT_ID>] [--sheet-id <GOOGLE_SHEET_ID>] [--deployment-id <DEPLOYMENT_ID>] [--desc <DESCRIPTION>]"
      exit 0
      ;;
    *)
      log_error "Unknown argument: $1"
      exit 1
      ;;
  esac
done

echo "================================================================================"
echo "  Tirwin Email List Validator — Google Apps Script Auto-Deployment"
echo "================================================================================"

# ==============================================================================
# 1. Load Secrets: Check `.env` first; if not found, fall back to Git Secrets
#    Uses a Node.js parser instead of bash `source` so unescaped JSON quotes in
#    `CLASPRC_JSON="{...}"` never get mangled by bash!
# ==============================================================================
ENV_FILE="$PROJECT_DIR/.env"
SECRET_SOURCE=""

load_dotenv_file() {
  local file="$1"
  eval "$(
    ENV_PATH="$file" node -e '
      const fs = require("fs");
      const lines = fs.readFileSync(process.env.ENV_PATH, "utf8").split(/\r?\n/);
      for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line || line.startsWith("#")) continue;
        const eqIdx = line.indexOf("=");
        if (eqIdx === -1) continue;
        const key = line.slice(0, eqIdx).trim();
        if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) continue;
        let val = line.slice(eqIdx + 1).trim();
        if ((val.startsWith("\"") && val.endsWith("\"")) || (val.startsWith("\x27") && val.endsWith("\x27"))) {
          val = val.slice(1, -1);
        }
        // Safely single-quote for bash eval
        const bashSafe = "\x27" + val.replace(/\x27/g, "\x27\"\\\x27\"\x27") + "\x27";
        console.log(`export ${key}=${bashSafe}`);
      }
    '
  )"
}

if [[ -f "$ENV_FILE" ]]; then
  log_info "Found local '.env' file. Loading configuration and secrets from '.env'..."
  load_dotenv_file "$ENV_FILE"
  SECRET_SOURCE=".env file"
else
  log_warn "No '.env' file found at $ENV_FILE. Looking for Git Secrets..."

  # Case A: Check if `git-secret` encrypted file (.env.secret) exists in the repo
  if [[ -f "$PROJECT_DIR/.env.secret" ]] && command -v git-secret >/dev/null 2>&1; then
    log_info "Found encrypted '.env.secret'. Revealing via 'git secret reveal'..."
    git secret reveal -f
    if [[ -f "$ENV_FILE" ]]; then
      load_dotenv_file "$ENV_FILE"
      SECRET_SOURCE="git-secret (.env.secret)"
    fi
  fi

  # Case B: Check GitHub Actions / Git CI environment secrets (e.g. ${{ secrets.SCRIPT_ID }})
  if [[ -z "$SECRET_SOURCE" ]]; then
    if [[ -n "${CLASPRC_JSON:-}" ]] || [[ -n "${SCRIPT_ID:-}" ]] || [[ -n "${GITHUB_ACTIONS:-}" ]]; then
      log_ok "Detected Git Secrets from CI/GitHub Actions environment variables."
      SECRET_SOURCE="Git / GitHub Actions Secrets"
    elif command -v git >/dev/null 2>&1 && git config --get tirwin.scriptId >/dev/null 2>&1; then
      log_info "Loading secrets from local git config (tirwin.*)..."
      SCRIPT_ID="${SCRIPT_ID:-$(git config --get tirwin.scriptId || true)}"
      DEPLOYMENT_ID="${DEPLOYMENT_ID:-$(git config --get tirwin.deploymentId || true)}"
      PARENT_SHEET_ID="${PARENT_SHEET_ID:-$(git config --get tirwin.parentSheetId || true)}"
      ALLOWED_DOMAIN="${ALLOWED_DOMAIN:-$(git config --get tirwin.allowedDomain || true)}"
      VERIFICATION_PROVIDER="${VERIFICATION_PROVIDER:-$(git config --get tirwin.verificationProvider || true)}"
      SMTP_API_KEY="${SMTP_API_KEY:-$(git config --get tirwin.smtpApiKey || true)}"
      SECRET_SOURCE="git config secrets"
    else
      log_warn "Neither '.env' nor pre-populated Git Secrets were found. Proceeding with defaults / interactive login."
      SECRET_SOURCE="defaults"
    fi
  fi
fi

# Apply defaults & CLI overrides
SCRIPT_ID="${CLI_SCRIPT_ID:-${SCRIPT_ID:-}}"
PARENT_SHEET_ID="${CLI_SHEET_ID:-${PARENT_SHEET_ID:-}}"
DEPLOYMENT_ID="${CLI_DEPLOYMENT_ID:-${DEPLOYMENT_ID:-}}"
CLASPRC_JSON="${CLASPRC_JSON:-}"
ALLOWED_DOMAIN="${ALLOWED_DOMAIN:-tirwin.in}"
BATCH_SIZE="${BATCH_SIZE:-100}"
VERIFICATION_PROVIDER="${VERIFICATION_PROVIDER:-NONE}"
SMTP_API_KEY="${SMTP_API_KEY:-}"

log_ok "Active secret source: $SECRET_SOURCE (Allowed Domain: @${ALLOWED_DOMAIN}, Batch Size: ${BATCH_SIZE})"

# Helper to persist newly created SCRIPT_ID or DEPLOYMENT_ID back into .env if .env exists
save_to_dotenv_if_present() {
  local key="$1"
  local val="$2"
  if [[ -f "$ENV_FILE" ]] && [[ -n "$val" ]]; then
    KEY="$key" VAL="$val" ENV_PATH="$ENV_FILE" node -e '
      const fs = require("fs");
      const file = process.env.ENV_PATH;
      const key = process.env.KEY;
      const val = process.env.VAL;
      let content = fs.readFileSync(file, "utf8");
      const regex = new RegExp("^" + key + "=.*$", "m");
      if (regex.test(content)) {
        content = content.replace(regex, `${key}="${val}"`);
      } else {
        content += `\n${key}="${val}"\n`;
      }
      fs.writeFileSync(file, content, "utf8");
    '
  fi
}

# ==============================================================================
# 2. Check Node.js & Install @google/clasp
# ==============================================================================
if ! command -v node >/dev/null 2>&1 || ! command -v npx >/dev/null 2>&1; then
  log_error "Node.js and npx are required. Please install Node.js (https://nodejs.org)."
  exit 1
fi

if [[ -x "$PROJECT_DIR/node_modules/.bin/clasp" ]]; then
  CLASP_CMD="$PROJECT_DIR/node_modules/.bin/clasp"
elif command -v clasp >/dev/null 2>&1; then
  CLASP_CMD="clasp"
else
  log_info "Installing Google Apps Script CLI (@google/clasp) locally..."
  npm install --no-audit --no-fund --save-dev @google/clasp
  CLASP_CMD="$PROJECT_DIR/node_modules/.bin/clasp"
fi
log_ok "Using clasp: $($CLASP_CMD --version 2>/dev/null || echo 'installed')"

# ==============================================================================
# 3. Verify Required Source Files
# ==============================================================================
for required_file in "Code.gs" "Index.html" "appsscript.json"; do
  if [[ ! -f "$PROJECT_DIR/$required_file" ]]; then
    log_error "Required file '$required_file' is missing in $PROJECT_DIR"
    exit 1
  fi
done
log_ok "Verified source files: Code.gs, Index.html, appsscript.json"

# ==============================================================================
# 4. Normalize & Provision ~/.clasprc.json
#    Converts between clasp v3 (`{"tokens":{"default":{...}}}`) and
#    clasp v2 (`{"token":{...},"oauth2ClientSettings":{...},"isLocalCreds":false}`)
#    so both versions work seamlessly without `reading 'access_token'` errors.
# ==============================================================================
normalize_clasprc_file() {
  RAW_CLASPRC="${1:-}" HOME_CLASPRC="$HOME/.clasprc.json" node -e '
    const fs = require("fs");
    const targetPath = process.env.HOME_CLASPRC;
    let rawInput = (process.env.RAW_CLASPRC || "").trim();

    if (!rawInput && fs.existsSync(targetPath)) {
      rawInput = fs.readFileSync(targetPath, "utf8").trim();
    }
    if (!rawInput) process.exit(0);

    // Support Base64 encoded string if it does not start with "{"
    if (!rawInput.startsWith("{")) {
      try {
        const decoded = Buffer.from(rawInput, "base64").toString("utf8").trim();
        if (decoded.startsWith("{")) rawInput = decoded;
      } catch (e) {}
    }

    let data;
    try {
      data = JSON.parse(rawInput);
    } catch (err) {
      // Repair unquoted keys if someone had bash-stripped quotes in ~/.clasprc.json
      const repaired = rawInput.replace(/([{,]\s*)([a-zA-Z0-9_]+)\s*:/g, "$1\"$2\":");
      data = JSON.parse(repaired);
    }

    const DEFAULT_CLIENT_ID = "1072944905499-vm2v2i5dvn0a0d2o4ca36i1vge8cvbn0.apps.googleusercontent.com";
    const DEFAULT_CLIENT_SECRET = "v6V3fKV_zWU7iw1DrpO1rknX";

    // Extract token fields whether input is clasp v3 (tokens.default), clasp v2 (token), or flat
    const v3Def = (data.tokens && data.tokens.default) ? data.tokens.default : null;
    const v2Tok = data.token ? data.token : null;
    const src = v2Tok || v3Def || data;

    const accessToken = src.access_token || "";
    const refreshToken = src.refresh_token || "";
    const clientId = (v3Def && v3Def.client_id) || (data.oauth2ClientSettings && data.oauth2ClientSettings.clientId) || DEFAULT_CLIENT_ID;
    const clientSecret = (v3Def && v3Def.client_secret) || (data.oauth2ClientSettings && data.oauth2ClientSettings.clientSecret) || DEFAULT_CLIENT_SECRET;

    const normalized = {
      token: {
        access_token: accessToken,
        refresh_token: refreshToken,
        scope: src.scope || "https://www.googleapis.com/auth/script.deployments https://www.googleapis.com/auth/script.projects https://www.googleapis.com/auth/script.webapp.deploy https://www.googleapis.com/auth/drive.metadata.readonly https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/service.management https://www.googleapis.com/auth/logging.read https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/cloud-platform",
        token_type: src.token_type || "Bearer",
        ...(src.expiry_date ? { expiry_date: src.expiry_date } : {})
      },
      oauth2ClientSettings: {
        clientId: clientId,
        clientSecret: clientSecret,
        redirectUri: (data.oauth2ClientSettings && data.oauth2ClientSettings.redirectUri) || "http://localhost"
      },
      isLocalCreds: Boolean(data.isLocalCreds),
      tokens: {
        default: {
          client_id: clientId,
          client_secret: clientSecret,
          type: "authorized_user",
          refresh_token: refreshToken,
          access_token: accessToken
        }
      }
    };

    fs.writeFileSync(targetPath, JSON.stringify(normalized, null, 2), { mode: 0o600 });
  '
}

if [[ -n "$CLASPRC_JSON" ]] || [[ -f "$HOME/.clasprc.json" ]]; then
  log_info "Validating and normalizing ~/.clasprc.json OAuth token structure (clasp v2/v3 compatible)..."
  normalize_clasprc_file "$CLASPRC_JSON"
fi

if [[ ! -f "$HOME/.clasprc.json" ]]; then
  if [[ -n "${CI:-}" ]] || [[ -n "${GITHUB_ACTIONS:-}" ]]; then
    log_error "Running in CI without CLASPRC_JSON secret! Please add CLASPRC_JSON to your .env or GitHub Actions Secrets."
    exit 1
  fi
  log_warn "No active Google clasp login found (~/.clasprc.json)."
  log_info "Opening browser to sign in with your @${ALLOWED_DOMAIN} Google Workspace account..."
  log_info "(Note: Ensure the Apps Script API is enabled at https://script.google.com/home/usersettings)"
  "$CLASP_CMD" login
  normalize_clasprc_file ""
fi
log_ok "Google clasp credentials ready."

# ==============================================================================
# 5. Create or Link Apps Script Project (.clasp.json)
# ==============================================================================
if [[ -n "$SCRIPT_ID" ]]; then
  log_info "Linking to Script ID from $SECRET_SOURCE: $SCRIPT_ID"
  cat > "$PROJECT_DIR/.clasp.json" <<EOF
{
  "scriptId": "$SCRIPT_ID",
  "rootDir": "$PROJECT_DIR"
}
EOF
fi

if [[ ! -f "$PROJECT_DIR/.clasp.json" ]] || ! grep -q '"scriptId"' "$PROJECT_DIR/.clasp.json"; then
  if [[ -n "$PARENT_SHEET_ID" ]]; then
    log_info "Creating new Google Apps Script project bound to Google Sheet ID: $PARENT_SHEET_ID ..."
    "$CLASP_CMD" create --title "$PROJECT_TITLE" --parentId "$PARENT_SHEET_ID" --rootDir "$PROJECT_DIR"
  else
    log_info "Creating new standalone Google Apps Script Web App project: '$PROJECT_TITLE' ..."
    "$CLASP_CMD" create --type webapp --title "$PROJECT_TITLE" --rootDir "$PROJECT_DIR"
  fi
  log_ok "Created Google Apps Script project (.clasp.json)."
fi

SCRIPT_ID=$(node -e "const fs = require('fs'); const c = JSON.parse(fs.readFileSync('.clasp.json', 'utf8')); console.log(c.scriptId || '');")
if [[ -z "$SCRIPT_ID" ]]; then
  log_error "Could not read 'scriptId' from .clasp.json."
  exit 1
fi
save_to_dotenv_if_present "SCRIPT_ID" "$SCRIPT_ID"
log_ok "Target Script ID: $SCRIPT_ID"

# ==============================================================================
# 6. Inject Runtime Secrets into Code.gs, Push to Apps Script, and Restore Clean File
# ==============================================================================
BACKUP_CODE_GS="$PROJECT_DIR/.Code.gs.bak"
cp "$PROJECT_DIR/Code.gs" "$BACKUP_CODE_GS"

cleanup_backup() {
  if [[ -f "$BACKUP_CODE_GS" ]]; then
    mv "$BACKUP_CODE_GS" "$PROJECT_DIR/Code.gs"
  fi
}
trap cleanup_backup EXIT

log_info "Injecting secrets from $SECRET_SOURCE into Code.gs for deployment..."
ALLOWED_DOMAIN="$ALLOWED_DOMAIN" \
BATCH_SIZE="$BATCH_SIZE" \
VERIFICATION_PROVIDER="$VERIFICATION_PROVIDER" \
SMTP_API_KEY="$SMTP_API_KEY" \
node -e '
  const fs = require("fs");
  let code = fs.readFileSync("Code.gs", "utf8");
  const esc = s => String(s || "").replace(/\\/g, "\\\\").replace(/\x27/g, "\\\x27");

  code = code.replace(/ALLOWED_DOMAIN:\s*\x27[^\x27]*\x27/, `ALLOWED_DOMAIN: \x27${esc(process.env.ALLOWED_DOMAIN)}\x27`);
  code = code.replace(/BATCH_SIZE:\s*\d+/, `BATCH_SIZE: ${parseInt(process.env.BATCH_SIZE, 10) || 100}`);
  code = code.replace(/VERIFICATION_PROVIDER:\s*\x27[^\x27]*\x27/, `VERIFICATION_PROVIDER: \x27${esc(process.env.VERIFICATION_PROVIDER)}\x27`);
  code = code.replace(/API_KEY:\s*\x27[^\x27]*\x27/, `API_KEY: \x27${esc(process.env.SMTP_API_KEY)}\x27`);

  fs.writeFileSync("Code.gs", code, "utf8");
'

log_info "Pushing files (Code.gs, Index.html, appsscript.json) to Google Apps Script..."
"$CLASP_CMD" push --force

# Immediately restore clean Code.gs so secrets are never left in the git working tree
cleanup_backup
trap - EXIT
log_ok "All project files pushed and local Code.gs restored cleanly."

# ==============================================================================
# 7. Create or Update Web App Deployment
# ==============================================================================
LAST_DEPLOYMENT_FILE="$PROJECT_DIR/.last_deployment_id"
if [[ -z "$DEPLOYMENT_ID" ]] && [[ -f "$LAST_DEPLOYMENT_FILE" ]]; then
  DEPLOYMENT_ID=$(tr -d '[:space:]' < "$LAST_DEPLOYMENT_FILE")
fi

DEPLOY_OUTPUT=""
if [[ -n "$DEPLOYMENT_ID" ]]; then
  log_info "Updating existing Web App deployment ($DEPLOYMENT_ID)..."
  if DEPLOY_OUTPUT=$("$CLASP_CMD" deploy --deploymentId "$DEPLOYMENT_ID" --description "$DEPLOY_DESC" 2>&1); then
    log_ok "Updated existing deployment: $DEPLOYMENT_ID"
  else
    log_warn "Could not update deployment $DEPLOYMENT_ID. Creating a new deployment instead..."
    DEPLOYMENT_ID=""
  fi
fi

if [[ -z "$DEPLOYMENT_ID" ]]; then
  log_info "Creating new versioned Web App deployment..."
  DEPLOY_OUTPUT=$("$CLASP_CMD" deploy --description "$DEPLOY_DESC" 2>&1)
  echo "$DEPLOY_OUTPUT"

  NEW_ID=$(echo "$DEPLOY_OUTPUT" | grep -Eo 'AKfycb[a-zA-Z0-9_-]+' | tail -n 1 || true)
  if [[ -n "$NEW_ID" ]]; then
    DEPLOYMENT_ID="$NEW_ID"
  fi
fi

if [[ -n "$DEPLOYMENT_ID" ]]; then
  echo "$DEPLOYMENT_ID" > "$LAST_DEPLOYMENT_FILE"
  save_to_dotenv_if_present "DEPLOYMENT_ID" "$DEPLOYMENT_ID"
  log_ok "Active Deployment ID: $DEPLOYMENT_ID"
fi

echo "================================================================================"
log_ok "AUTO-DEPLOYMENT COMPLETED SUCCESSFULLY!"
echo "--------------------------------------------------------------------------------"
echo "  Secret Source     : ${SECRET_SOURCE}"
echo "  Script Editor URL : https://script.google.com/d/${SCRIPT_ID}/edit"
if [[ -n "$DEPLOYMENT_ID" ]]; then
  echo "  Web App URL       : https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec"
  echo "  Domain Web App URL: https://script.google.com/a/macros/${ALLOWED_DOMAIN}/s/${DEPLOYMENT_ID}/exec"
  echo ""
  echo "  -> Copy the Domain Web App URL above and embed it into your Google Site"
  echo "     (Insert -> Embed -> By URL)."
fi
echo "================================================================================"
