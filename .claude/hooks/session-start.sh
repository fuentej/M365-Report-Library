#!/bin/bash
# Installs PowerShell 7 and Pester 5 in cloud sessions so ./shared/tests and
# ./reports/*/tests can run with Invoke-Pester. PSGallery (www.powershellgallery.com,
# codeload.github.com) is denied by the cloud sandbox's egress proxy, so Pester comes
# from the nuget.org flat-container feed instead. GitHub release downloads and
# api.nuget.org are both allowed.
set -euo pipefail

# Only cloud sessions hit this network path; local machines install PowerShell and
# Pester through their own package manager or PSGallery, which work there.
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

PWSH_VERSION="7.4.6"
PESTER_VERSION="5.7.1"
PWSH_HOME="/opt/pwsh"
PWSH_BIN="/usr/local/bin/pwsh"
PESTER_MODULE_DIR="/usr/local/share/powershell/Modules/Pester/${PESTER_VERSION}"

if command -v pwsh >/dev/null 2>&1; then
  echo "session-start: pwsh already installed: $(pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()')"
else
  echo "session-start: installing PowerShell ${PWSH_VERSION}..."
  curl -fsSL -o /tmp/pwsh.tar.gz \
    "https://github.com/PowerShell/PowerShell/releases/download/v${PWSH_VERSION}/powershell-${PWSH_VERSION}-linux-x64.tar.gz"
  mkdir -p "$PWSH_HOME"
  tar -xzf /tmp/pwsh.tar.gz -C "$PWSH_HOME"
  chmod +x "$PWSH_HOME/pwsh"
  ln -sf "$PWSH_HOME/pwsh" "$PWSH_BIN"
  rm -f /tmp/pwsh.tar.gz
  echo "session-start: installed pwsh $(pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()')"
fi

if [ -f "$PESTER_MODULE_DIR/Pester.psd1" ]; then
  echo "session-start: Pester ${PESTER_VERSION} already installed at ${PESTER_MODULE_DIR}"
else
  echo "session-start: installing Pester ${PESTER_VERSION} from nuget.org..."
  curl -fsSL -o /tmp/pester.nupkg \
    "https://api.nuget.org/v3-flatcontainer/pester/${PESTER_VERSION}/pester.${PESTER_VERSION}.nupkg"
  rm -rf /tmp/pesterx
  python3 -c "import zipfile; zipfile.ZipFile('/tmp/pester.nupkg').extractall('/tmp/pesterx')"
  mkdir -p "$PESTER_MODULE_DIR"
  cp -r /tmp/pesterx/tools/. "$PESTER_MODULE_DIR/"
  rm -f /tmp/pester.nupkg
  rm -rf /tmp/pesterx
  echo "session-start: installed Pester ${PESTER_VERSION}"
fi

exit 0
