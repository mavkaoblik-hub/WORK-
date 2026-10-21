# Offline v8: fail fast when Excel Trust Center denies VBA project changes.
# This file tracks the VBA permission check that is embedded in customer ZIP v8.
# Dot-source this function after opening a disposable workbook in Excel COM.
function Assert-TransportVbaAccess($guardWorkbook) {
    $component = $null
    $components = $null
    try {
        $project = $guardWorkbook.VBProject
        if ($null -eq $project) { throw 'VBProject is null.' }
        $components = $project.VBComponents
        if ($null -eq $components) { throw 'VBComponents is null.' }
        $component = $components.Add(1)
        if ($null -eq $component) { throw 'VBComponents.Add returned null.' }
        $component.Name = 'TC_AccessPreflight'
        $code = $component.CodeModule
        if ($null -eq $code) { throw 'CodeModule is null.' }
        $code.AddFromString("Option Explicit`r`nPublic Sub TC_AccessPreflight()`r`nEnd Sub")
        return $true
    } catch {
        throw ('VBA_ACCESS_BLOCKED: ' + $_.Exception.Message +
          ' Enable Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to VBA project object model, then retry. Disable again after successful installation.')
    } finally {
        if ($null -ne $component -and $null -ne $components) {
            try { $components.Remove($component) } catch { Write-Warning 'Could not remove temporary VBA probe.' }
        }
    }
}
