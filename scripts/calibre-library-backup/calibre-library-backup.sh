#!/bin/bash
set -euo pipefail

# =============================================================================
# Calibre & Kobo Backup Script
# Backs up: Calibre library (books + metadata + config) and Kobo annotations
# Destination: Proton Drive
# =============================================================================

# --- Config ------------------------------------------------------------------
CALIBRE_LIBRARY="$HOME/Calibre"
CALIBRE_CONFIG="$HOME/Library/Preferences/calibre"
PROTON=$(echo "$HOME/Library/CloudStorage/ProtonDrive-"*"-folder" 2>/dev/null | awk '{print $1}')
BACKUP_ROOT="$PROTON/Backups/Calibre"
DATE=$(date +%Y-%m-%d)
KOBO_DB="/Volumes/KOBOeReader/.kobo/KoboReader.sqlite"
LOG_DIR="$HOME/.local/share/calibre-backup"
LOCAL_LOG="$LOG_DIR/backup.log"
LOCKDIR="$LOG_DIR/calibre-backup.lock"
ERRORS=0
TMP_DB=""
TMP_LOG=""

# --- Log dir and temp file cleanup -------------------------------------------
mkdir -p "$LOG_DIR"
chmod 700 "$LOG_DIR"
trap 'rm -f "${TMP_DB:-}" "${TMP_LOG:-}" 2>/dev/null; rmdir "${LOCKDIR:-}" 2>/dev/null || true' EXIT

# --- Preflight: verify required tools are present ----------------------------
for cmd in rsync sqlite3; do
  if ! command -v "$cmd" > /dev/null 2>&1; then
    echo "ERROR: required tool '$cmd' not found — aborting" >&2
    exit 1
  fi
done

# --- Guard: prevent concurrent runs ------------------------------------------
if ! mkdir "$LOCKDIR" 2>/dev/null; then
  echo "ERROR: Another backup instance is running — aborting" >&2
  exit 1
fi

# --- Logging -----------------------------------------------------------------
log() { echo "[$(date +%H:%M:%S)] $1" | tee -a "$LOCAL_LOG"; }

echo "" >> "$LOCAL_LOG"
log "===== Backup started: $(date) ====="

# --- Guard: Proton Drive must be mounted -------------------------------------
if [ ! -d "$PROTON" ]; then
  log "ERROR: Proton Drive not mounted at $PROTON — aborting"
  osascript -e 'display notification "Proton Drive not mounted — backup aborted" with title "Calibre Backup"' 2>/dev/null || true
  exit 1
fi

# --- Guard: Calibre library must exist and not be empty ----------------------
if [ ! -f "$CALIBRE_LIBRARY/metadata.db" ]; then
  log "ERROR: Calibre library not found at $CALIBRE_LIBRARY — aborting"
  exit 1
fi

# --- Setup dirs --------------------------------------------------------------
mkdir -p "$BACKUP_ROOT/books"
mkdir -p "$BACKUP_ROOT/config"
mkdir -p "$BACKUP_ROOT/kobo"
chmod 700 "$BACKUP_ROOT"

# --- 1. Books & Metadata (no --delete: keeps deleted books in backup) --------
log "Syncing Calibre library..."
if rsync -a \
  --include="*.epub" \
  --include="*.pdf" \
  --include="*.azw3" \
  --include="*.mobi" \
  --include="metadata.opf" \
  --include="cover.jpg" \
  --include="*/" \
  --exclude="*" \
  "$CALIBRE_LIBRARY/" \
  "$BACKUP_ROOT/books/" >> "$LOCAL_LOG" 2>&1; then
  log "OK: Books synced"
else
  log "ERROR: Books sync failed"
  ERRORS=$((ERRORS + 1))
fi

# --- 2. Calibre metadata.db (warn if Calibre is running) ---------------------
log "Backing up Calibre database..."
if pgrep -i "calibre" > /dev/null 2>&1; then
  log "WARNING: Calibre is running — metadata.db backup may be inconsistent"
fi
TMP_DB=$(mktemp "$LOG_DIR/metadata.db.XXXXXX")
if cp "$CALIBRE_LIBRARY/metadata.db" "$TMP_DB" 2>> "$LOCAL_LOG"; then
  mv "$TMP_DB" "$BACKUP_ROOT/books/metadata.db"
  TMP_DB=""
  log "OK: Calibre database backed up"
else
  rm -f "$TMP_DB"
  TMP_DB=""
  log "ERROR: Calibre database backup failed"
  ERRORS=$((ERRORS + 1))
fi

# --- 3. Calibre config -------------------------------------------------------
log "Backing up Calibre config..."
if rsync -a \
  "$CALIBRE_CONFIG/" \
  "$BACKUP_ROOT/config/" >> "$LOCAL_LOG" 2>&1; then
  log "OK: Config backed up"
else
  log "ERROR: Config backup failed"
  ERRORS=$((ERRORS + 1))
fi

# --- 4. Kobo annotations (only if Kobo is mounted) ---------------------------
if [ -f "$KOBO_DB" ]; then
  log "Kobo detected — backing up..."
  KOBO_BACKUP="$BACKUP_ROOT/kobo/KoboReader_${DATE}.sqlite"

  # Always overwrite today's sqlite (handles multiple sessions per day).
  # sqlite3 .backup produces a consistent snapshot even when WAL is active.
  if sqlite3 "$KOBO_DB" ".backup '$KOBO_BACKUP'" 2>> "$LOCAL_LOG"; then
    log "OK: Kobo database backed up"
  else
    log "ERROR: Kobo database backup failed"
    ERRORS=$((ERRORS + 1))
  fi

  # Always re-export annotations (captures any new highlights since last run)
  log "Exporting annotations to CSV..."
  if sqlite3 "$KOBO_DB" -csv \
    "SELECT b.Title, b.Attribution, bm.Text, bm.Annotation, bm.DateCreated
     FROM Bookmark bm
     JOIN Content b ON bm.VolumeID = b.ContentID
     WHERE bm.Text IS NOT NULL
     ORDER BY b.Title, bm.DateCreated;" \
    > "$BACKUP_ROOT/kobo/annotations_${DATE}.csv" 2>> "$LOCAL_LOG"; then
    ANNOTATION_COUNT=$(awk 'END{print NR}' "$BACKUP_ROOT/kobo/annotations_${DATE}.csv")
    log "OK: $ANNOTATION_COUNT annotations exported"
  else
    log "ERROR: Annotation export failed"
    ERRORS=$((ERRORS + 1))
  fi

  # Keep only last 30 Kobo backups (filenames are date-stamped, no spaces)
  find "$BACKUP_ROOT/kobo" -name "KoboReader_*.sqlite" | \
    sort -r | tail -n +31 | xargs rm -f 2>/dev/null || true
  find "$BACKUP_ROOT/kobo" -name "annotations_*.csv" | \
    sort -r | tail -n +31 | xargs rm -f 2>/dev/null || true
else
  log "Kobo not mounted — skipping annotation backup"
fi

# --- Rotate local log (keep last 500 lines) ----------------------------------
TMP_LOG=$(mktemp)
tail -n 500 "$LOCAL_LOG" > "$TMP_LOG" && mv "$TMP_LOG" "$LOCAL_LOG"
TMP_LOG=""

# --- Copy log to backup ------------------------------------------------------
cp "$LOCAL_LOG" "$BACKUP_ROOT/backup.log" 2>/dev/null || true

# --- Done --------------------------------------------------------------------
if [ "$ERRORS" -gt 0 ]; then
  log "===== Backup COMPLETED WITH $ERRORS ERROR(S): $(date) ====="
  osascript -e "display notification \"Calibre backup finished with $ERRORS error(s) — check log\" with title \"Calibre Backup\"" 2>/dev/null || true
else
  log "===== Backup complete: $(date) ====="
  osascript -e 'display notification "Calibre backup complete" with title "Calibre Backup"' 2>/dev/null || true
fi

log "Backup location: $BACKUP_ROOT"
