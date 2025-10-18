# FreeCut PowerPoint插件安装脚本
# 需要管理员权限运行

param(
    [switch]$Uninstall
)

# 检查管理员权限
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "此脚本需要管理员权限。正在重新启动..." -ForegroundColor Yellow
    Start-Process PowerShell -Verb RunAs "-File `"$PSCommandPath`" $PSBoundParameters"
    exit
}

$pluginName = "FreeCut"
$pluginPath = Join-Path $PSScriptRoot "deploy\FreeCut.dll"
$manifestPath = Join-Path $PSScriptRoot "deploy\FreeCut.dll.manifest"

# PowerPoint版本检测
$powerpointVersions = @("16.0", "15.0", "14.0") # Office 2019/2016, 2013, 2010

function Install-Plugin {
    Write-Host "开始安装FreeCut插件..." -ForegroundColor Green

    # 检查文件是否存在
    if (-not (Test-Path $pluginPath)) {
        Write-Host "错误：找不到插件文件 $pluginPath" -ForegroundColor Red
        Write-Host "请先运行 build.bat 编译项目" -ForegroundColor Yellow
        return
    }

    $installed = $false

    foreach ($version in $powerpointVersions) {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Office\$version\PowerPoint\Addins\$pluginName"
        $regPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$version\PowerPoint\Addins\$pluginName"

        try {
            # 64位注册表
            if (Test-Path "HKLM:\SOFTWARE\Microsoft\Office\$version\PowerPoint") {
                New-Item -Path $regPath -Force | Out-Null
                Set-ItemProperty -Path $regPath -Name "Description" -Value "FreeCut - PPT自动裁剪PDF导出插件"
                Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "FreeCut"
                Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
                Set-ItemProperty -Path $regPath -Name "Manifest" -Value $manifestPath
                Write-Host "已注册到PowerPoint $version (64位)" -ForegroundColor Green
                $installed = $true
            }

            # 32位注册表
            if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$version\PowerPoint") {
                New-Item -Path $regPath32 -Force | Out-Null
                Set-ItemProperty -Path $regPath32 -Name "Description" -Value "FreeCut - PPT自动裁剪PDF导出插件"
                Set-ItemProperty -Path $regPath32 -Name "FriendlyName" -Value "FreeCut"
                Set-ItemProperty -Path $regPath32 -Name "LoadBehavior" -Value 3 -Type DWord
                Set-ItemProperty -Path $regPath32 -Name "Manifest" -Value $manifestPath
                Write-Host "已注册到PowerPoint $version (32位)" -ForegroundColor Green
                $installed = $true
            }
        } catch {
            Write-Host "注册PowerPoint $version 时出错: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($installed) {
        Write-Host "`n安装完成！" -ForegroundColor Green
        Write-Host "请重启PowerPoint，然后在功能区查找'FreeCut'标签页。" -ForegroundColor Cyan
        Write-Host "`n如果插件未显示，请检查PowerPoint的加载项设置：" -ForegroundColor Yellow
        Write-Host "文件 → 选项 → 加载项 → 管理COM加载项 → 确认FreeCut已启用" -ForegroundColor Yellow
    } else {
        Write-Host "未找到兼容的PowerPoint版本。" -ForegroundColor Red
    }
}

function Uninstall-Plugin {
    Write-Host "开始卸载FreeCut插件..." -ForegroundColor Yellow

    $uninstalled = $false

    foreach ($version in $powerpointVersions) {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Office\$version\PowerPoint\Addins\$pluginName"
        $regPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$version\PowerPoint\Addins\$pluginName"

        try {
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force
                Write-Host "已从PowerPoint $version (64位)卸载" -ForegroundColor Green
                $uninstalled = $true
            }

            if (Test-Path $regPath32) {
                Remove-Item -Path $regPath32 -Recurse -Force
                Write-Host "已从PowerPoint $version (32位)卸载" -ForegroundColor Green
                $uninstalled = $true
            }
        } catch {
            Write-Host "卸载PowerPoint $version 时出错: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($uninstalled) {
        Write-Host "`n卸载完成！请重启PowerPoint。" -ForegroundColor Green
    } else {
        Write-Host "未找到已安装的插件。" -ForegroundColor Yellow
    }
}

# 主逻辑
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    FreeCut PowerPoint插件安装程序" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($Uninstall) {
    Uninstall-Plugin
} else {
    Install-Plugin
}

Write-Host "`n按任意键退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")