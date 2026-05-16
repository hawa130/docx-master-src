# Math regression corpus

Inputs that the in-tree MathML → OMML converter
(`lib/edit/math/mml-to-omml/`) must keep handling. One `.tex` per case
in `cases/`, optional snapshot `<case>.expected.omml` next to it.

## Running

```bash
bun run test:math            # check: schema-validate + snapshot diff
bun run test:math:update     # regenerate every <case>.expected.omml
```

The runner does, per `.tex`:

1. temml → MathML (display mode, `throwOnError: true`)
2. `convertMathMLToOMML(...)` → OMML string
3. Validate against `shared-math.xsd` via `xmllint-wasm` (libxml2)
4. If `<case>.expected.omml` exists, diff (whitespace-normalized)

A regression is either a schema error or a diff against a previously
snapshot-locked expected. **Word visual verification is still a manual
step** — schema validity is necessary but not sufficient. After any
change to a fixture's output, open the rendered .docx in Microsoft Word
and confirm the equation renders as intended.

## Adding a case

1. Drop a `name.tex` into `cases/`.
2. `bun run test:math` — schema-only check first, makes sure the new
   case doesn't throw and produces valid OMML.
3. Open the produced OMML inside a Word document by hand and verify
   visually.
4. Once visually correct: `bun run test:math:update` to lock the
   snapshot.

## When `expected.omml` diffs

A diff means *either* a real regression *or* an intentional
improvement. Inspect the new output:

- New shape is wrong → fix the converter; do not update the snapshot.
- New shape is better (e.g. nesting tightened, redundant wrapper
  removed) → eyeball-verify in Word, then `--update` to re-snapshot.

Don't blindly `--update` on a red corpus — each snapshot represents
Word-verified ground truth at the time it was committed.
