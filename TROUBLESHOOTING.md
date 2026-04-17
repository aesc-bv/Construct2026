# Troubleshooting

## Duplicate "AESC Construct" sidebar panel (one empty, one with content)

### Symptom

The SpaceClaim right sidebar shows **two** "AESC Construct" panel tabs side by side. The left one is empty (no content, no close behaviour). The right one has the actual panel content (Settings, Frame Generator, etc.). The empty tab cannot be removed through the UI.

### Root cause

The addin is registered with `host="NewAppDomain"` in `AESCConstruct2026.xml`. SpaceClaim occasionally reloads the addin in a fresh AppDomain during a running session — not only on startup, but also on events like document switching or internal addin refreshes.

When an AppDomain is torn down, SpaceClaim *sometimes* calls `Construct2026.Disconnect()` and *sometimes* does not. If `Disconnect()` is not called (or does no cleanup), the old `PanelTab` remains registered in SpaceClaim's persisted layout file:

```
C:\Users\<user>\AppData\Local\SpaceClaim\SpaceClaim\barlayout2.xml
```

Each un-cleaned reload adds another `<bar name="AESCConstruct2026.ConstructPanel" ...>` entry. On the next startup, SpaceClaim recreates all of them, and the addin additionally calls `PanelTab.Create` once more — producing N+1 tabs.

### Current mitigation (commit 2e9a2cd)

`Construct2026.Disconnect()` now calls `UIManager.ClosePanelAndDispose()`, which calls `PanelTab.Close()` and disposes the `ElementHost`. On the shutdown paths where SpaceClaim does call `Disconnect()`, the bar entry is properly removed from `barlayout2.xml` and no orphan accumulates.

This is a **partial** fix. Empirically SpaceClaim calls `Disconnect()` on ~1 in 5 AppDomain reloads. In-session silent reloads still leak.

### Recovery when it occurs

1. Close SpaceClaim completely (verify no `SpaceClaim.exe` in Task Manager).
2. Move the layout file out of the way:
   ```
   move "C:\Users\%USERNAME%\AppData\Local\SpaceClaim\SpaceClaim\barlayout2.xml" "barlayout2.xml.bak"
   ```
3. Start SpaceClaim. It regenerates `barlayout2.xml` from `barlayout_default2.xml` (which contains **zero** `AESCConstruct2026.ConstructPanel` entries), and our `RegisterConstructPanel` creates exactly one clean tab.

### Diagnostic

To check if orphans have accumulated:
```
grep -c "AESCConstruct2026.ConstructPanel" "%LOCALAPPDATA%\SpaceClaim\SpaceClaim\barlayout2.xml"
```
- Expected value: **2** (one `<bar>` and one `<item>` for the live panel).
- Higher value: orphans have accumulated. Run the recovery steps.

### Full fix (not yet viable)

Switching to `host="SameAppDomain"` in `AESCConstruct2026.xml` eliminates the AppDomain reload entirely and would prevent all orphans. Tried in April 2026 — it broke the ribbon (generic SpaceClaim sketch tools appeared instead of our custom ribbon buttons), likely due to an assembly-loading conflict with the main SpaceClaim AppDomain. Needs dedicated investigation of which references / types conflict before that path is available.

### Code entry points

- `Construct2026.cs` — `Disconnect()`
- `UIMain/UIManager.cs` — `ClosePanelAndDispose()`, `RegisterConstructPanel()`

Do not remove the Disconnect cleanup or downgrade it to a no-op — it is load-bearing for this issue even though the benefit isn't visible until the next SpaceClaim reload.
