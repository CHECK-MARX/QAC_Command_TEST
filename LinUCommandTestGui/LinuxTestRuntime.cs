using System;
using System.IO;

namespace LinUCommandTestGui;

public static class LinuxTestRuntime
{
    public static string? FindOriginalScript(string root)
    {
        foreach (var name in new[] { "testloop.sh", "test_loop.sh" })
        {
            var path = Path.Combine(root, name);
            if (File.Exists(path)) return path;
        }
        return null;
    }

    public static bool IsTerminalDiagnostic(string line) =>
        line.TrimStart().StartsWith("[TERMINAL-DIAGNOSTIC] ", StringComparison.Ordinal);

    public const string EditorScript = """
        #!/bin/sh
        # qacli may pipe stdin to its editor; reconnect all streams to the terminal.
        exec /usr/bin/vi "$@" </dev/tty >/dev/tty 2>&1
        """;

    public const string ShellBridge = """
        # GUI terminal bridge. Ordinary commands run unchanged, without retries.
        # This file is generated automatically; edit LinuxTestRuntime.cs instead.
        qacli() {
          # Interactive editors must have a visible terminal, not captured output.
          if [ "${1:-}" = project ] && [ "${2:-}" = config-file ]; then
            local qacli_arg
            for qacli_arg in "$@"; do
              case "$qacli_arg" in
                --edit|--edit=*|-E|-E?*)
                  printf '[EDITOR] 別端末で設定ファイルを開きます。エディタを終了するとテストを続行します。\n'
                  local qacli_editor_result qacli_editor_rc qacli_terminal_log qacli_terminal_rc
                  qacli_editor_result="$(mktemp "${TMPDIR:-/tmp}/linu-editor.XXXXXX")" || return 1
                  qacli_terminal_log="$(mktemp "${TMPDIR:-/tmp}/linu-terminal.XXXXXX")" || { rm -f "$qacli_editor_result"; return 1; }
                  local -a qacli_editor_command=(env EDITOR="$LINU_TERMINAL_EDITOR" VISUAL="$LINU_TERMINAL_EDITOR"
                    sh -c 'result_file=$1; shift; "$@"; rc=$?; printf "%s\n" "$rc" > "$result_file"; exit 0'
                    qacli-editor "$qacli_editor_result" "$LINU_REAL_QACLI" "$@")
                  if command -v ptyxis >/dev/null 2>&1; then
                    ptyxis --standalone --working-directory="$PWD" -- "${qacli_editor_command[@]}" 2>"$qacli_terminal_log"
                  elif command -v gnome-terminal >/dev/null 2>&1; then
                    gnome-terminal --wait --working-directory="$PWD" -- "${qacli_editor_command[@]}" 2>"$qacli_terminal_log"
                  elif command -v xterm >/dev/null 2>&1; then
                    xterm -e "${qacli_editor_command[@]}" 2>"$qacli_terminal_log"
                  else
                    printf '[ERR] 編集用端末がありません。ptyxis、gnome-terminal、xtermのいずれかが必要です。\n' >&2
                  fi
                  qacli_terminal_rc=$?
                  # Keep terminal diagnostics in the log with their origin, separate from QAC failures.
                  while IFS= read -r qacli_diagnostic || [ -n "$qacli_diagnostic" ]; do
                    printf '[TERMINAL-DIAGNOSTIC] %s\n' "$qacli_diagnostic"
                  done < "$qacli_terminal_log"
                  rm -f "$qacli_terminal_log"
                  if [ "$qacli_terminal_rc" -ne 0 ]; then
                    rm -f "$qacli_editor_result"
                    printf '[ERR] 編集用端末の実行に失敗しました (exit=%s)\n' "$qacli_terminal_rc" >&2
                    return "$qacli_terminal_rc"
                  fi
                  qacli_editor_rc="$(cat "$qacli_editor_result")"
                  rm -f "$qacli_editor_result"
                  case "$qacli_editor_rc" in
                    ''|*[!0-9]*) printf '[ERR] 編集用端末が正常に完了しませんでした。\n' >&2; return 1 ;;
                  esac
                  printf '[EDITOR] 編集終了 (exit=%s)\n' "$qacli_editor_rc"
                  if [ "$qacli_editor_rc" -ne 0 ]; then
                    printf '[ERR] 編集コマンドが失敗しました (exit=%s)\n' "$qacli_editor_rc" >&2
                  fi
                  return "$qacli_editor_rc"
                  ;;
              esac
            done
          fi
          # Preserve terminal I/O, output timing, arguments and status. Execute exactly once.
          "$LINU_REAL_QACLI" "$@"
        }
        export -f qacli
        """;
}
