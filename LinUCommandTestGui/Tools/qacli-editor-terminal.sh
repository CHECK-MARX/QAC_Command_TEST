#!/usr/bin/env bash
set -e

TERMINAL="$(command -v gnome-terminal)"
EDITOR_BIN="$(command -v vi)"

if [ -z "$TERMINAL" ]; then
    echo "gnome-terminal が見つかりません。" >&2
    exit 127
fi

if [ -z "$EDITOR_BIN" ]; then
    echo "vi が見つかりません。" >&2
    exit 127
fi

exec "$TERMINAL" --wait -- "$EDITOR_BIN" "$@"
