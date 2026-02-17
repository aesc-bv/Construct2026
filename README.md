# AESCConstruct2026

SpaceClaim AddIn for ANSYS Discovery (v251) providing structural modeling tools: frame/profile generation, fasteners, plates, connectors, rib cutouts, engraving, and custom component properties.

## Build

Build with Visual Studio 2017+ or `msbuild AESCConstruct2026.sln`. See [CLAUDE.md](CLAUDE.md) for full build configuration and architecture documentation.

## AddIn Registration

`AESCConstruct2026.xml` must exist in the output directory alongside the DLL. If missing after build, copy it manually:

```xml
<?xml version="1.0" encoding="utf-8" ?>
<AddIns>
  <AddIn
    name="AESC Construct 2026"
    description="A custom add-in for SpaceClaim."
    assembly="AESCConstruct2026.dll"
    typename="AESCConstruct2026.Construct2026"
    host="NewAppDomain" />
</AddIns>
```
