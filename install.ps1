#Requires -Version 3.0

[CmdletBinding(SupportsShouldProcess)]
param ()

dynamicparam {
    $Attributes = [Collections.ObjectModel.Collection[Attribute]]::new()
    $ParamAttribute = [Parameter]::new()
    $ParamAttribute.Position = 0
    $ParamAttribute.ParameterSetName = '__AllParameterSets'
    $Attributes.Add($ParamAttribute)

    [string[]]$FontNames = Join-Path $PSScriptRoot patched-fonts | Get-ChildItem -Directory -Name
    $Attributes.Add([ValidateSet]::new(($FontNames)))

    $Parameter = [Management.Automation.RuntimeDefinedParameter]::new('FontName', [string[]], $Attributes)
    $RuntimeParams = [Management.Automation.RuntimeDefinedParameterDictionary]::new()
    $RuntimeParams.Add('FontName', $Parameter)

    return $RuntimeParams
}

end {
    Function New-FontResourceType {
        $fontCSharpCode = @'
            using System;
            using System.Collections.Generic;
            using System.Text;
            using System.IO;
            using System.Runtime.InteropServices;
    
            namespace FontResource {
                public class AddFonts {
                    private static IntPtr HWND_BROADCAST = new IntPtr(0xffff);
    
                    [DllImport("gdi32.dll")]
                    static extern int AddFontResource(string lpFilename);
    
                    [DllImport("gdi32.dll")]
                    static extern int RemoveFontResource(string lpFileName);
    
                    [return: MarshalAs(UnmanagedType.Bool)]
                    [DllImport("user32.dll", SetLastError = true)]
                    private static extern bool PostMessage(IntPtr hWnd, WM Msg, IntPtr wParam, IntPtr lParam);
    
                    public enum WM : uint {
                        FONTCHANGE = 0x001D
                    }
                
                    public static int AddFont(string fontFilePath) {
                        FileInfo fontFile = new FileInfo(fontFilePath);
                        if (!fontFile.Exists) { throw new FileNotFoundException("Font file not found"); }
                        try {
                            int retVal = AddFontResource(fontFilePath);
                            bool posted = PostMessage(HWND_BROADCAST, WM.FONTCHANGE, IntPtr.Zero, IntPtr.Zero);
                            return retVal;
                        }
                        catch { throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error()); }
                    }
                }
            }
'@

        Add-Type $fontCSharpCode
    }

    $FontName = $PSBoundParameters.FontName
    if (-not $FontName) { $FontName = '*' }

    $fontFiles = [Collections.Generic.List[System.IO.FileInfo]]::new()

    Join-Path $PSScriptRoot patched-fonts | Push-Location
    foreach ($aFontName in $FontName) {
        Get-ChildItem $aFontName -Filter "*.ttf" -Recurse | Foreach-Object { $fontFiles.Add($_) }
        Get-ChildItem $aFontName -Filter "*.otf" -Recurse | Foreach-Object { $fontFiles.Add($_) }
    }
    Pop-Location

    try { $LoadedType = [FontResource.AddFonts].ispublic } catch { $LoadedType = $false }
    if (-not $LoadedType) { New-FontResourceType }

    $fontRegistryPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    $fontFileTypes = @{
        '.fon' = ''
        '.fnt' = ''
        '.ttf' = ' (TrueType)'
        '.ttc' = ' (TrueType)'
        '.otf' = ' (OpenType)'
    }

    $shellApp = New-Object -ComObject shell.application
    $fonts = $shellApp.NameSpace(0x14)
    $fontsFolder = $fonts.self.path
    foreach ($fontFile in $fontFiles) {
        $folderObj = (New-Object -ComObject shell.application).namespace($fontFile.directoryname)
        $fileObj = $folderObj.Items().Item($fontFile.Name)
        $fontName = $folderObj.GetDetailsOf($fileObj, 21)

        Copy-Item $fontFile.FullName -Destination $fontsFolder -Force

        $fontFinalPath = Join-Path $fontsFolder $fontFile.Name
        $retVal = [FontResource.AddRemoveFonts]::AddFont($fontFinalPath)

        if ($retVal -eq 0) {
            Write-Host "Font resource, '$($fontFile.FullName)', installation failed"
        }
        else {
            Write-Host "Font resource, '$($fontFile.Name)', installed successfully" -ForegroundColor Green
            Set-ItemProperty -Path "$($fontRegistryPath)" -Name "$fontName$($fontFileTypes.item($fontFile.Extension))" -Value "$($fontFile.Name)" -Type STRING -Force
        }
    }
}