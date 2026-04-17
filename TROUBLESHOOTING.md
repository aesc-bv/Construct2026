# Troubleshooting

## Duplicate "AESC Construct" sidebar panel (one empty, one with content)

### Symptom

The SpaceClaim right sidebar shows **two or more** "AESC Construct" panel tabs. The extra ones are empty; one has the actual panel content. The empty tabs cannot be removed through the UI.

### Root cause (confirmed by `barlayout2.xml` inspection)

`UIManager.EnsureConstructPanel()` used to silently call `PanelTab.Create()` every time the SpaceClaim remoting proxy reported `_constructPanelTab.IsDeleted == true` **or** threw a `RemotingException`. That report was unreliable — the visual tab stayed alive in the sidebar, so each recreation stacked a fresh `PanelTab` on top of the still-existing one. Over time `barlayout2.xml` accumulated multiple `<bar name="AESCConstruct2026.ConstructPanel">` entries, each corresponding to an orphaned visual tab.

A secondary contributor: SpaceClaim skips calling `IExtensibility.Disconnect()` on ~80% of its in-session AppDomain reloads, so cleanup was also not reliably reached from that path.

### Fix (commit de7533a)

1. **Primary** — `UIManager.EnsureConstructPanel` no longer recreates the `PanelTab` based on `IsDeleted`. If the proxy misreports, we log it and reuse the existing tab. Only the `ElementHost` is recreated (if null), attaching into the same `PanelTab` slot.
2. **Fallback** — `Construct2026.Connect` registers `AppDomain.DomainUnload` and `AppDomain.ProcessExit` handlers that call the same cleanup as `Disconnect()`, catching the reload paths SpaceClaim doesn't notify.

### Recovery if orphans exist from earlier builds

If you upgrade from a pre-de7533a build, the old `barlayout2.xml` may still contain multiple bars from past silent recreations. One-time recovery:

1. Close SpaceClaim completely (check Task Manager for any `SpaceClaim.exe`).
2. Move the layout file aside:
   ```
   move "%LOCALAPPDATA%\SpaceClaim\SpaceClaim\barlayout2.xml" "%LOCALAPPDATA%\SpaceClaim\SpaceClaim\barlayout2.xml.bak"
   ```
3. Start SpaceClaim — it regenerates `barlayout2.xml` from `barlayout_default2.xml` (which has zero `AESCConstruct` entries). Your addin creates exactly one clean tab.

### Diagnostic

Check accumulated bars in the layout file:
```
grep -c "AESCConstruct2026.ConstructPanel" "%LOCALAPPDATA%\SpaceClaim\SpaceClaim\barlayout2.xml"
```
- Expected value: **2** (one `<bar>`, one `<item>`).
- Higher value: the fix is not running or old orphans are still there. Run the recovery steps.

Check the addin log for the diagnostic line:
```
grep "IsDeleted\|RECREATING" "%PROGRAMDATA%\AESCConstruct\AESCConstruct2026_Log.txt"
```
- Expected: `IsDeleted=true reported (ignored — not recreating)` entries.
- If you see `RECREATING PanelTab` lines, the old code path is still in use — check the deployed DLL version.

### Code entry points (do not regress)

- `Construct2026.cs` — `Connect()` registers `DomainUnload`/`ProcessExit`; `Disconnect()` calls `UIManager.ClosePanelAndDispose()`
- `UIMain/UIManager.cs` — `EnsureConstructPanel()` must NOT call `PanelTab.Create` a second time; `ClosePanelAndDispose()` is the shared cleanup
