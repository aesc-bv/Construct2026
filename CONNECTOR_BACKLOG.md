# Connector — Bug & Improvement Backlog

Triaged backlog for the Connector module, produced 2026-05-19 ahead of focused bug-fix work.

**All line references are in the canonical `Connector2/` module** (the shipping module as of
the 2026-05-19 cutover — see `.claude/agents/connector-specialist.md`). Every item below
also exists identically in the frozen legacy `Connector/` (Connector2 is a near-copy); do
**not** fix it there — `Connector/` is slated for deletion.

Re-read the cited code before fixing; line numbers drift. There is no automated test path —
every geometry change must be verified manually in SpaceClaim.

Priority: **P1** = produces wrong geometry / silent incorrect output; **P2** = robustness or
missing feature that bites real users; **P3** = cleanup / polish / dead code.

> **2026-05-19 — review pass 2:** added P2-6, P2-7, P3-8; sharpened P1-1 and P3-6; fixed the
> stale-`dirZ_w` bug (see **Resolved**). That fix added ~4 lines to
> `ComputePlanarConnectorBodies`, so `Connector2Control.xaml.cs` line numbers below ~720 have
> shifted by roughly +4. Re-grep the cited symbols — do not trust absolute line numbers.

---

## Resolved

### R-1 · Stale `dirZ_w` cast the dynamic-height ray the wrong way — FIXED 2026-05-19
- **Was:** `ComputePlanarConnectorBodies` computed `dirZ_w` from `dirZ_m` *before* the
  `ContainsPoint` test that can flip `dirZ_m` (face normal pointing into the solid). The
  dynamic-height ray `ComputeAvailableHeightAlongRay(..., -dirZ_w, ...)` then probed along
  the **pre-flip** direction → Dynamic height mis-clipped (wrong side / wrong distance) on
  any planar face whose normal initially points into the body.
- **Fix:** moved the `dirZ_w` computation to *after* the flip block so it is always derived
  from the final `dirZ_m` (one-line move + WHY comment). Compiles clean. **Not yet manually
  verified in SpaceClaim** — test Dynamic height on a face whose normal points inward.

---

## P1 — correctness / silent wrong output

### P1-1 · `OffsetFaces` failure swallowed → cutter not inflated by Tolerance, no feedback
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:1445` (cylindrical non-rect cut:
  `try { cutBody.OffsetFaces(cutBody.Faces, tolM); } catch { Logger.Log(...) }`),
  `:840-846` (planar inflate path, same swallow + `inflated?.Dispose()`),
  `:1550-1554` (cylindrical inflate path). Underlying op also at `Connector2/Connector2.cs:742`.
- **Symptom:** when `OffsetFaces` throws, the catch only logs and disposes the temp
  `inflated` copy — then **execution falls straight through to
  `return new PlanarConnectorResult { … CutBody = cutBody … }`** with `cutBody` still the
  original un-inflated body (or `null`). The caller treats the result as success, so the
  subtracted cavity is exactly the connector size — zero clearance, no warning; only a log
  line records it. (Mechanism confirmed at the planar inflate path: catch at the
  `OffsetFaces inflate failed` log line, immediately followed by the `PlanarConnectorResult`
  return — no failure flag is set.)
- **Risk:** High. Produces physically wrong assemblies that look fine in the viewport.
- **Fix direction:** On offset failure, either (a) abort the placement with a localized
  `Application.ReportStatus` / status message and skip committing that connector, or
  (b) fall back to the rectangular widened-cutter path (`max(halfW1,halfW2)+tol`) so
  clearance is still applied. Do **not** silently continue with a zero-clearance cutter.
- **Effort:** M (touches 3 sites + a user-facing message + a decision on fallback policy).

### P1-2 · `UpdateGenerateEnabled` is a no-op → Create never disabled on invalid input
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:2460` (def), `:2462-2464`
  (`FindName("btnCreateConnector") ?? FindName("btnCreate") ?? FindName("generateButton")`).
- **Symptom:** the XAML Create button has no `x:Name`, so all three `FindName` probes return
  null and the method never toggles button state. Invalid combinations (e.g. EndRelief on a
  trapezoidal connector) are only stopped by the `IsReliefAllowed` hard guard at the top of
  `createConnector` (`:919`) — the button still looks clickable, so the user clicks and gets
  a bail-out status message instead of a disabled control.
- **Risk:** Medium (no wrong geometry — the hard guard holds — but poor UX and a latent trap
  if anyone removes the hard guard trusting the button state).
- **Fix direction:** give the Create button an `x:Name` in `Connector2Control.xaml` and use
  it directly (drop the `FindName` probe list), or delete `UpdateGenerateEnabled` entirely
  and rely on the hard guard plus a clear status message. Pick one; don't leave a method
  that pretends to gate the button.
- **Effort:** S.

### P1-3 · Edge-width validation only covers a single selected edge
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:3005` (`CheckSelectedEdgeWidth`),
  called from `createConnector` `:932`.
- **Symptom:** `Width1 > availableEdgeLength` is only checked for the single-edge case.
  Multi-edge selections and pattern placements (`ComputePatternCentersOcc`, `:2051`) are not
  pre-validated, so an over-wide connector on a short edge in a batch silently produces
  degenerate/zero geometry or a failed boolean for that placement.
- **Risk:** Medium-High (silent bad/missing connectors in batch runs).
- **Fix direction:** validate width per resolved placement inside the per-edge loop (after
  `getFacesFromSelection`), accumulate skipped placements, and report a summary status
  ("3 of 12 placements skipped: connector wider than edge").
- **Effort:** M.

---

## P2 — robustness / feature gaps

### P2-1 · Cylindrical 3D preview not implemented (silently skipped)
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:2870` (`UpdatePreview`), filter at
  `:2964` (`if (!(bigFace.Shape.Geometry is Plane)) continue;`).
- **Symptom:** enabling 3D preview with a cylindrical edge selected shows nothing; the
  preview just `continue`s past non-planar faces with no message. Users assume the connector
  won't be created.
- **Fix direction:** either implement the cylindrical preview by reusing the cylindrical
  loft path from `createConnector` (`:~1142`), or, minimum, surface a status note
  ("3D preview is planar-only; cylindrical connectors are still created on Create").
- **Effort:** L (full preview) / S (status note).

### P2-2 · Dynamic-height probe radius hard-coded 0.5 mm
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:578` (`double probeR = 0.0005;`),
  related step floor `:513` (`Math.Max(0.0005, maxT/50.0)`), used by
  `ComputeAvailableHeightAlongRay` (`:560`).
- **Symptom:** the collision probe is a fixed 0.5 mm cylinder. Thin obstructions (< ~1 mm
  features, thin sheet) can be missed → connector clipped too long and interferes; or the
  coarse step over/under-shoots the true exit distance.
- **Fix direction:** scale the probe radius/step from the connector footprint (e.g. a
  fraction of `min(Width1,Width2)` or the wall thickness) instead of a constant; consider a
  finer binary-search tolerance.
- **Effort:** M (geometry-sensitive — needs manual SpaceClaim validation on tight models).

### P2-3 · `CylInfo.GetCoaxialCylPair` throws instead of degrading
- **Where:** `Connector2/CylInfo2.cs:27,28,32,36` (`ArgumentNullException` / `ArgumentException`
  for null body, null face, unresolvable master, non-cylindrical seed).
- **Symptom:** a bad/edge-case selection on the cylindrical branch throws out of the
  per-edge `WriteBlock` instead of being caught and reported as "skip this edge".
- **Fix direction:** wrap the cylindrical-branch call in `createConnector` so a throw logs +
  reports + skips that edge (consistent with the planar branch's defensive style), or have
  `GetCoaxialCylPair` return a success-tuple instead of throwing.
- **Effort:** S–M.

### P2-4 · Per-edge `WriteBlock` can commit partial geometry on mid-loop failure
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:1034`
  (`WriteBlock.ExecuteTask("Connectors (per-edge commit)", …)`).
- **Symptom:** everything is one undo step, but if a boolean op throws on placement N after
  placements 1..N-1 are applied, the caught exception lets the block complete — leaving a
  partially-applied result that a single Undo removes wholesale (no per-edge recovery).
- **Fix direction:** decide and document the policy — either fully transactional (roll back
  the whole block on any failure) or per-edge commit with a summary of which edges
  succeeded/failed. Currently it's neither cleanly.
- **Effort:** M.

### P2-5 · Preset apply ignores boolean fields
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:391` (`ApplyPresetToUi`), invoked from
  `ConnectorPresetCombo_SelectionChanged` (`:367`).
- **Symptom:** picking a preset overwrites only Height/Width1/Width2/Radius/Tolerance + the
  radius/chamfer radio. DynamicHeight, CornerCutout, RectangularCut, Pattern, etc. keep the
  user's prior state — so the same preset yields different geometry depending on prior UI
  state. Surprising and hard to support.
- **Fix direction:** decide whether presets are full snapshots (apply booleans too — the
  `ConnectorPresetRecord` already has the columns) or explicitly numeric-only (then document
  it in the preset help text). Recommend full snapshot for predictability.
- **Effort:** S–M.

### P2-6 · Unused/irrelevant form fields can block ALL connector creation
- **Where:** `Connector2/Connector2.cs:105-106` (`cornerCutoutRadius`,
  `radiusInCutOut_Radius`) and `:118` (`patternQty`) — `CreateConnector` strictly parses
  these via `ParseWithTrace` **unconditionally**, regardless of whether Pattern / Corner
  cutout are checked. `connectorCornerCutoutRadiusValue` is in the `Visibility="Hidden"`
  XAML row. `patternQty` is also re-parsed tolerantly later in `createConnector` (~`:960`).
- **Symptom:** if any of those textboxes is blank or non-numeric (e.g. user clears the
  Pattern box while Pattern is off), `CreateConnector` returns `null` → generic
  `Connector_Msg_FillValidValues` warning → **no connector can be created at all**, with no
  hint that an unused/hidden field is the cause.
- **Risk:** Medium-High UX (classic "it just won't work and I don't know why" support call).
- **Fix direction:** only parse a field when its feature is enabled (mirror the gating in
  `createConnector`), default-on-blank to `0`/feature value, and drop the redundant strict
  `patternQty` parse here in favour of the tolerant one in `createConnector`.
- **Effort:** S.

### P2-7 · Orphaned modeler `Body` on the extrude sign-flip path
- **Where:** `Connector2/Connector2.cs:643-650` (`CreateGeometry`: first `main` is
  re-extruded and the original is never disposed when the extrude direction is wrong). Same
  pattern at `:180` (`GetDynamicHeight`) and `:443` (`DrawWireBox`).
- **Symptom:** every placement whose profile extrudes the "wrong" way (≈half of them, by
  geometry) leaks one intermediate `Body`. Not user-visible per-run, but accumulates over a
  session / large patterns and is sloppy in a tight per-edge `WriteBlock`.
- **Risk:** Low-Medium (resource/handle pressure over long sessions).
- **Fix direction:** dispose the first `main` before re-extruding (or compute the correct
  sign first and extrude once). Apply the same to the two sibling sites.
- **Effort:** S.

---

## P3 — cleanup / polish / dead code

### P3-1 · CSV fallback path uses singular `\Connector\` (wrong folder)
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:275` (`LoadConnectorPresets`, fallback
  path construction). Settings default + the real file are
  `C:\ProgramData\AESCConstruct\Connectors\ConnectorProperties.csv` (**plural**).
- **Fix:** make the fallback use `Connectors\` (plural) to match `Settings.settings:39` and
  the on-disk folder. Effort: S.

### P3-2 · `NormalizeStemSuffixes` logs every component name
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:1673`.
- **Fix:** gate the per-component `Logger.Log` behind a debug flag (like `LOG_CYL`). Chatty
  on large assemblies; deep-recurses on pathological name collisions. Effort: S.

### P3-3 · `s_neighbourDecisionCache` is dead code
- **Where:** `Connector2/UI/Connector2Control.xaml.cs:63` (declared, `.Clear()`ed at `:890`,
  never written/read otherwise).
- **Fix:** remove it, or wire the intended per-neighbour decision reuse. Effort: S.

### P3-4 · `ConnectorStraight` vestigial flag + commented-out checkbox
- **Where:** model prop consumed at `Connector2/UI/Connector2Control.xaml.cs:1206-1207,
  1272-1273`; commented `//connectorStraight` at `:2337`; XAML checkbox `connectorStraight`
  around `Connector2/UI/Connector2Control.xaml:484-485`. Always `false` from UI.
- **Fix:** either re-expose the "straight vs radial on cylinders" option (it has live
  geometry effect) or remove the dead UI + simplify the branches. Decide intent first.
  Effort: M.

### P3-5 · Vestigial overlapping XAML rows (`Grid.Row="12"` twice)
- **Where:** `Connector2/UI/Connector2Control.xaml:398-400` (hidden "cutout radius" grid,
  `Visibility="Hidden"`) and `:430` (second `Grid.Row="12"`, Click-location). They coexist by
  overlap, not row index.
- **Fix:** if the hidden top-pair-cutout-radius UI is never coming back, delete the hidden
  grid and reflow rows; otherwise document the intentional overlap. Effort: S.

### P3-6 · `AreCoaxial` dead math — and a latent exception hazard
- **Where:** `Connector2/CylInfo2.cs:95-106`. Only `a.IsCoincident(b)` gates the return;
  `Direction.Cross(a.Direction, b.Direction)` (`:101`) and `cos` (`:103`) are computed and
  discarded. The `angTol`/`linearTol` XML-doc params don't exist on the method.
- **Hazard:** `Direction.Cross` of two **parallel** axes (exactly the coaxial inner/outer
  wall case this is called for) produces a ~zero vector; SpaceClaim's `Direction` rejects a
  zero-length direction. So `:101` is not just dead — for the normal tube case it is a
  potential throw masked only by call timing. Cylindrical connectors work today, so verify
  whether `Direction.Cross` is being tolerated or silently caught upstream.
- **Fix:** delete `:101` and `:103` (and the bogus param docs); keep only the
  `IsCoincident` test. If a parallel-but-offset tolerance is ever wanted, implement it
  explicitly without an unconditional cross of expected-parallel axes. Effort: S.

### P3-7 · `Connector_Tooltip_EndRelief` near-duplicate translation entry
- **Where:** `languageConstruct.csv` (+ `_translation_staging*` working files). Near-duplicate
  EndRelief tooltip strings.
- **Fix:** dedupe; ensure the XAML tooltip and the localized key resolve to one source.
  Effort: S.

### P3-8 · `GetDynamicHeight` is broken dead code
- **Where:** `Connector2/Connector2.cs:140-186`.
- **Symptom:** the method computes `dynamicHeight = 0`, builds and `ExtrudeProfile`s a
  `Body` (twice on the sign-flip path) that is **never disposed and never used**, and
  unconditionally `return height;` (the input, unchanged). `midPoint`, `p01`, `p06`,
  `sign1` are computed and unused. Real dynamic-height logic lives in
  `ComputeAvailableHeightAlongRay` (code-behind), so this appears entirely dead — but if
  anything ever calls it, it is a pure no-op that leaks modeler bodies (see also P2-7).
- **Fix:** delete the method (confirm no callers first — grep `GetDynamicHeight`). Effort: S.

0. ~~R-1 stale `dirZ_w`~~ — **done 2026-05-19** (pending manual SpaceClaim verification).
1. **P1-1** (wrong geometry, highest user impact) → **P2-6** (blocks creation, cheap fix) →
   **P1-3** → **P1-2**.
2. **P2-2** and **P2-1** (dynamic height / cylindrical preview — most-noticed gaps).
3. **P2-3 / P2-4 / P2-5 / P2-7** robustness.
4. P3 sweep incl. **P3-6** (verify the `Direction.Cross` hazard first) and **P3-8**
   (mostly mechanical; bundle into one cleanup commit).

After Connector2 is validated in SpaceClaim, do the deferred legacy removal: drop the
`Connector\*` `<Compile>`/`<Page>` entries from `AESCConstruct2026.csproj`, delete
`Connector/`, and update `WEEKLY_UPDATE_NL.md:61,94`.
