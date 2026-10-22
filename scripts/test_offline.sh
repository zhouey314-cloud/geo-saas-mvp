#!/bin/sh
set -eu
cd "$(dirname "$0")/.."
python_bin=$(command -v "${PYTHON_BIN:-python3}")
test_home=$(mktemp -d "${TMPDIR:-/tmp}/geo-saas-tests.XXXXXX")
trap 'rm -rf "$test_home"' EXIT HUP INT TERM
# Do not inherit API credentials, proxy settings, or the user's Streamlit secrets.
env -i PATH="$(dirname "$python_bin"):/usr/bin:/bin" HOME="$test_home" \
    TMPDIR="$test_home" PYTHONDONTWRITEBYTECODE=1 \
    "$python_bin" -m unittest discover -s tests -v
