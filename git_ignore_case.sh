#!/bin/bash
# -----------------------------------------------------------------------------
# git_ignore_case.sh
# -----------------------------------------------------------------------------
# Normaliza diferencias “cosméticas” antes de actualizar el working copy desde
# el índice, para que Git no marque cambios por:
#   - Mayúsculas/minúsculas (incluyendo después de un '.': .Pos vs .pos)
#   - Retornos de carro CR (Windows) y líneas en blanco
#
# Flujo:
#   1) Saca el archivo del índice (ORIGFILE).
#   2) Hace un diff case-insensitive y tolerante a CR/blank contra el working.
#   3) Aplica al ORIGFILE solo los cambios “reales” (no los de casing/CR).
#   4) Copia el ORIGFILE parcheado sobre el working file y fuerza CRLF.
#
# Importante:
#   - Forzamos LC_ALL=C para que -i (ignore case) se comporte de forma estable
#     en ASCII (esto corrige el caso .Pos vs .pos, Invent.Object vs invent.Object).
#   - No cambiamos el contenido real salvo lo que indique el patch.
#
# Requisitos:
#   - diff, patch, git, unix2dos
#
# Compatibilidad probada: Bash en Windows (MSYS/MinGW/Git Bash) y Linux.
# -----------------------------------------------------------------------------

set -euo pipefail

# Asegurar comparación ASCII pura (evita problemas de locale)
export LC_ALL=C
export LANG=C

# ---------------------------------------------------------------------
# Función auxiliar: genera PATCHFILE entre dos archivos respetando:
#   - ignore case (-i)
#   - ignore CR finales (--strip-trailing-cr)
#   - ignore blank lines (--ignore-blank-lines)
# Salida: crea/llena el archivo PATCHFILE y no corta el script si “diff” sale 1.
# ---------------------------------------------------------------------
make_patch() {
  local left="$1"   # archivo base (del índice)
  local right="$2"  # archivo del working copy
  local outpatch="$3"

  # Nota: Usamos diff normal (no streams procesados) para que el patch
  # aplique bien al archivo original (conserva contenido original).
  diff -u -i --strip-trailing-cr --ignore-blank-lines "$left" "$right" > "$outpatch" || true
}

# ---------------------------------------------------------------------
# Sección 1: Formularios, clases, módulos, diseñadores
# ---------------------------------------------------------------------
for file in $(git status --porcelain | grep -E "^.{1}M" | grep -Ev "^R" | cut -c 4- \
              | grep -E -e "\.frm$" -e "\.bas$" -e "\.cls$" -e "\.Dsr$"); do
  ORIGFILE=$(mktemp)
  PATCHFILE=$(mktemp)

  # Contenido del índice (HEAD)
  git cat-file -p ":$file" > "$ORIGFILE"

  # Generar el patch ignorando casing/CR/blank
  make_patch "$ORIGFILE" "$file" "$PATCHFILE"

  # Aplicar cambios "reales" al ORIGFILE y volcar al working
  if [ -s "$PATCHFILE" ]; then
    patch -s "$ORIGFILE" < "$PATCHFILE"
  fi

  cp "$ORIGFILE" "$file"
  unix2dos --quiet "$file"

  rm -f "$ORIGFILE" "$PATCHFILE"
done

# ---------------------------------------------------------------------
# Sección 2: Proyectos VB6 (.vbp)
#   - Mantiene tu lógica de comparar solo hasta el 3er '#' por línea,
#     pero ahora usando el mismo esquema de diff robusto.
# ---------------------------------------------------------------------
for file in $(git status --porcelain | cut -c 4- | grep -E "\.vbp$"); do
  ORIGFILE=$(mktemp)
  PATCHFILE=$(mktemp)
  LEFT_CUT=$(mktemp)
  RIGHT_CUT=$(mktemp)

  git cat-file -p ":$file" > "$ORIGFILE"

  # Tomamos solo los primeros 3 campos separados por '#'
  # (como tenías), tanto de índice como del working.
  # Ojo: preservamos el original (ORIGFILE) para parcharlo con PATCHFILE.
  awk -F'#' '{print $1"#"$2"#"$3}' "$ORIGFILE" > "$LEFT_CUT"  2>/dev/null || true
  awk -F'#' '{print $1"#"$2"#"$3}' "$file"     > "$RIGHT_CUT" 2>/dev/null || true

  # Generamos el patch a partir de los recortes, pero con contexto del ORIGINAL.
  # Estrategia: si los recortes no difieren (ignorando case/CR/blank), no tocamos.
  # Si difieren, generamos un patch normal entre ORIGFILE y el working file
  # completo, pero con las mismas banderas (para no romper coherencia).
  if ! diff -q -i --strip-trailing-cr --ignore-blank-lines "$LEFT_CUT" "$RIGHT_CUT" >/dev/null; then
    make_patch "$ORIGFILE" "$file" "$PATCHFILE"
    if [ -s "$PATCHFILE" ]; then
      patch -s "$ORIGFILE" < "$PATCHFILE"
    fi
  fi

  cp "$ORIGFILE" "$file"
  unix2dos --quiet "$file"

  rm -f "$ORIGFILE" "$PATCHFILE" "$LEFT_CUT" "$RIGHT_CUT"
done
