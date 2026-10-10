#!/bin/bash
# Ensure the .NET 8 SDK is present in cloud sessions, then restore.
set -euo pipefail
[ "${CLAUDE_CODE_REMOTE:-}" = "true" ] || exit 0
if ! command -v dotnet >/dev/null 2>&1; then
  apt-get update -qq
  apt-get install -y -qq dotnet-sdk-8.0
fi
cd "${CLAUDE_PROJECT_DIR:-.}"
dotnet restore PivotWright.sln >/dev/null
