#!/usr/bin/env bash
# The purpose of this script is to build the Decisions Module zip without having the
# Decisions Code Base checked out. Linux counterpart to BuildModule.ps1 (build-only;
# local deploy/service restart is Windows-specific and not covered here).

set -euo pipefail

find_module_name() {
    local build_proj="$1"
    grep -oP '(?<=-buildmodule )\S+' "$build_proj" | head -n1
}

get_compile_target() {
    local base_path="$1"
    local guess="$base_path/build.proj"
    if [[ -f "$guess" ]]; then
        echo "$guess"
        return
    fi
    echo "Could not find a build.proj file, please create one." >&2
    exit 1
}

base_path="$(pwd)"
echo "Using basePath - $base_path"

compile_target="$(get_compile_target "$base_path")"
module_name="$(find_module_name "$compile_target")"

echo "Building $module_name"
echo "Compiling Project by build.proj"
echo "Found Compile Target - $compile_target"

dotnet build "$compile_target"
