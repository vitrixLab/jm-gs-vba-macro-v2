# This script is intentionally disabled.
# The prior version wrote ten COA_MAP mappings and marked them MAPPED without
# accounting approval evidence, then ran RefreshGL. Do not use it to modify
# the workbook. Apply only mappings approved by the responsible accountant.
$ErrorActionPreference = "Stop"
throw "DISABLED: unsafe hard-coded COA mappings removed. No workbook was opened or modified."
