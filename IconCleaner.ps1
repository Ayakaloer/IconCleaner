# ============================================================================
#  IconCleaner.ps1  ——  Windows "此电脑/桌面" 流氓图标清理工具
# ----------------------------------------------------------------------------
#  功能:
#    1. 扫描"此电脑"与"桌面"上的第三方命名空间图标(CLSID 命名空间挂载项)
#    2. 自动识别常见流氓图标(百度网盘 / 阿里云盘 / 夸克 / 迅雷 / 360 / 2345 等)
#    3. 由用户勾选要删除的项(可自定义), 删除前自动备份, 可一键恢复
#    4. 支持"免疫防复活"、导出报告、重启资源管理器、注册表定位
#
#  用法:
#    双击 "启动图标清理工具.bat"(自动以管理员身份运行)
#    命令行扫描:  powershell -NoProfile -ExecutionPolicy Bypass -File IconCleaner.ps1 -ScanOnly
#    自检:        powershell -NoProfile -ExecutionPolicy Bypass -File IconCleaner.ps1 -SelfTest
# ============================================================================
param(
    [switch]$ScanOnly,   # 只扫描并输出到控制台, 不打开界面
    [switch]$SelfTest    # 对备份/删除/恢复/免疫逻辑做自检(在 HKCU 下创建临时键)
)

$ErrorActionPreference = 'SilentlyContinue'

# ---------------------------------------------------------------------------
# 0. 基础路径与数据目录
# ---------------------------------------------------------------------------
$script:Root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if ([string]::IsNullOrWhiteSpace($script:Root)) { $script:Root = (Get-Location).Path }
$script:DataDir = Join-Path $script:Root 'data'
try {
    New-Item -ItemType Directory -Force -Path $script:DataDir | Out-Null
} catch {
    $script:DataDir = Join-Path $env:LOCALAPPDATA 'IconCleaner'
    New-Item -ItemType Directory -Force -Path $script:DataDir | Out-Null
}
$script:BackupDir   = Join-Path $script:DataDir 'backups'
$script:LogFile     = Join-Path $script:DataDir 'log.txt'
$script:ConfigFile  = Join-Path $script:DataDir 'config.json'
New-Item -ItemType Directory -Force -Path $script:BackupDir | Out-Null

$script:LogBox       = $null
$script:SearchText   = ''
$script:ShowSystem   = $false
$script:ImmuneMarker = 'IconCleaner 免疫占位键'

# ---------------------------------------------------------------------------
# 1. 日志
# ---------------------------------------------------------------------------
function Write-Log([string]$Message) {
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    try {
        [System.IO.File]::AppendAllText($script:LogFile, $line + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
    } catch {}
    if ($script:LogBox) {
        $script:LogBox.AppendText($line + "`r`n")
        $script:LogBox.ScrollToEnd()
    }
    Write-Host $line
}

# ---------------------------------------------------------------------------
# 2. 特征库(内置 + 可编辑的 signatures.json)
# ---------------------------------------------------------------------------
$script:BuiltInWhitelist = @(
    '{20D04FE0-3AEA-1069-A2D8-08002B30309D}',  # 此电脑
    '{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}',  # 桌面
    '{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}',  # 文档
    '{374DE290-123F-4565-9164-39C4925E467B}',  # 下载
    '{1CF1260C-4DD0-4EB8-811F-33C572699FDE}',  # 音乐
    '{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}',  # 图片
    '{A0953C92-50DC-43BF-BE83-3742FED03C9C}',  # 视频
    '{0DB7E03F-FC29-4DC6-9020-FF41B59B631A}',  # 3D 对象
    '{018D5C66-4533-4307-9B53-224DE2ED1FE6}',  # OneDrive
    '{031E4825-7B94-4DC3-B131-E946B44C8DD5}',  # 库
    '{59031A47-3F72-44A7-89C5-5595FE6B30EE}',  # 用户文件夹 (DelegateFolders)
    '{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}',  # 网络 (DelegateFolders)
    '{26EE0668-A00A-44D7-9371-BEB064C98683}',  # 控制面板 (DelegateFolders)
    '{21EC2020-3AEA-1069-A2DD-08002B30309D}',  # 控制面板
    '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}',  # 所有控制面板项
    '{645FF040-5081-101B-9F08-00AA002F954E}',  # 回收站
    '{450D8FBA-AD25-11D0-98A8-0800361B1103}',  # 我的文档(旧)
    '{871C5380-42A0-1069-A2EA-08002B30309D}',  # Internet Explorer
    '{208D2C60-3AEA-1069-A2D7-08002B30309D}',  # 网上邻居
    '{B4FB3F98-C1EA-428D-A78A-D1F5659CBA93}',  # Other Users
    '{2559A1F1-21D7-11D4-BDAF-00C04F60B9F0}',  # 搜索
    '{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}',  # 运行
    '{2559A1F4-21D7-11D4-BDAF-00C04F60B9F0}',  # 帮助
    '{D20EA4E1-3957-11D2-A40B-0C5020524153}',  # 管理工具
    '{BD84B380-8CA2-1069-AB1D-08000948F534}',  # 字体
    '{0AFACED1-E828-11D1-9187-B532F1E9575D}',  # 文件夹快捷方式
    '{323CA680-C24D-4099-B94D-446DD2D7249E}',  # 收藏夹
    '{BFB9D5A0-C6A9-404C-B2B2-AE6DB6AF4968}',  # 链接
    # --- Windows 11 系统项(实测于 Win11 24H2, 不同版本略有差异) ---
    '{088E3905-0323-4B02-9826-5D99428E115F}',  # 下载(Win11)
    '{24AD3AD4-A569-4530-98E1-AB02F9417AA8}',  # 图片(旧)
    '{D3162B92-9365-467A-956B-92703ACA08AF}',  # 文档(旧)
    '{3DFDF296-DBEC-4FB4-81D1-6A3438BCF4DE}',  # 系统
    '{F86FA3AB-70D2-4FC7-9C99-FCBF05467F3A}',  # 网络位置
    '{04731B67-D933-450A-90E6-4ACD2E9408FE}',  # 搜索主页
    '{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}',  # 系统
    '{11016101-E366-4D22-BC06-4ADA335C892B}',  # 系统
    '{2F6CE85C-F9EE-43CA-90C7-8A9BD53A2467}',  # 系统
    '{4336A54D-038B-4685-AB02-99BB52D3FB8B}',  # 库(Win7)
    '{5B934B42-522B-4C34-BBFE-37A3EF7B9C90}',  # 系统
    '{64693913-1C21-4F30-A98F-4E52906D3B56}',  # 系统
    '{89D83576-6BD1-4C86-9454-BEB04E94C819}',  # 系统
    '{8FD8B88D-30E1-4F25-AC2B-553D3D65F0EA}',  # 系统
    '{9343812E-1C37-4A49-A12E-4B2D810D956B}',  # 搜索
    '{98F275B4-4FFF-11E0-89E2-7B86DFD72085}',  # 系统
    '{A00EE528-EBD9-48B8-944A-8942113D46AC}',  # 系统
    '{BD7A2E7B-21CB-41B2-A086-B309680C6B7E}',  # 系统
    '{DAF95313-E44D-46AF-BE1B-CBACEA2C3065}',  # 回收站(Win10)
    '{E345F35F-9397-435C-8F95-4E922C26259E}',  # 系统
    '{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}',  # 快速访问
    '{EDC978D6-4D53-4B2F-A265-5805674BE568}',  # 系统
    '{F8278C54-A712-415B-B593-B77A2BE0DDA9}',  # 系统
    '{F874310E-B6B7-47DC-BC84-B9E6B38F5903}',  # 库(Win11)
    '{289AF617-1CC3-42A6-926C-E6A863F0E3BA}',  # 系统委托文件夹
    '{35786D3C-B075-49B9-88DD-029876E11C01}',  # 系统委托文件夹
    '{9113A02D-00A3-46B9-BC5F-9C04DADDD5D7}',  # 系统委托文件夹
    '{B155BDF8-02F0-451E-9A26-AE317CFD7779}',  # 网络委托文件夹
    '{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}',  # OneDrive 委托文件夹
    '{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}',  # OneDrive 委托文件夹
    '{640167B4-59B0-47A6-B335-A6B3C0695AEA}'   # 便携媒体设备 (Windows Media Player)
)
# 系统项名称精确匹配(防止未本地化的 CLSID 令牌名被误判为第三方)
$script:BuiltInSystemNames = @(
    'this pc','computer','my computer','此电脑','我的电脑',
    'nethood','streambackedfolder','portable devices','removable drives',
    'search','public','mapi','dlna','csc','appinstancefolder','file','enhanced','dxp',
    'recycle bin','回收站','network','网络','libraries','library','库',
    'downloads','下载','documents','文档','pictures','图片','music','音乐','videos','视频',
    'desktop','桌面','control panel','控制面板','users files','用户文件夹',
    '3d objects','3d 对象','history','start menu','开始菜单','my documents','我的文档',
    'administrative tools','管理工具','fonts','字体','folder options','文件夹选项',
    'send to','favorites','收藏夹','links','链接','network explorer','connections'
)
$script:BuiltInRogueKeywords = @(
    '百度网盘','百度云','百度网','baidunetdisk',
    '阿里云盘','aliyundrive',
    '夸克网盘','夸克','quark',
    '天翼云盘','电信云盘','移动云盘','和彩云',
    '迅雷','thunder',
    '腾讯微云','微云','weiyun',
    'wps网盘','wps云盘','wps云','wps网','金山快盘',
    '360','2345','好压','快压','kuaizip','压缩宝',
    '驱动精灵','驱动人生','鲁大师','金山毒霸','电脑管家',
    '猎豹','搜狗','万能驱动','联想电脑管家',
    '爱奇艺','优酷','腾讯视频','酷狗','qq音乐','网易云音乐',
    '小鸟壁纸','元气壁纸'
)
$script:BuiltInThirdPartyKeywords = @(
    '网盘','云盘','云空间','cloud','drive','sync',
    '助手','管家','卫士','工具箱','加速器','游戏','手机助手','备份','transfer','share'
)

function Load-Signatures {
    $script:WhiteList          = @($script:BuiltInWhitelist)
    $script:RogueKeywords      = @($script:BuiltInRogueKeywords)
    $script:ThirdPartyKeywords = @($script:BuiltInThirdPartyKeywords)
    $sigFile = Join-Path $script:Root 'signatures.json'
    if (-not (Test-Path -LiteralPath $sigFile)) {
        try {
            $s = [ordered]@{
                rogueKeywords      = @($script:RogueKeywords)
                thirdPartyKeywords = @($script:ThirdPartyKeywords)
                whitelistClsid     = @($script:WhiteList)
            }
            [System.IO.File]::WriteAllText($sigFile, ($s | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($true)))
            Write-Log "已生成默认特征库: $sigFile (可自行编辑)"
        } catch {
            Write-Log "特征库写入失败: $($_.Exception.Message)"
        }
        return
    }
    try {
        $s = Get-Content -Raw -Encoding UTF8 -LiteralPath $sigFile | ConvertFrom-Json
        if ($s.rogueKeywords)      { $script:RogueKeywords      = @($s.rogueKeywords) }
        if ($s.thirdPartyKeywords) { $script:ThirdPartyKeywords = @($s.thirdPartyKeywords) }
        if ($s.whitelistClsid)     { $script:WhiteList += @($s.whitelistClsid) }
        Write-Log "已加载特征库: $sigFile"
    } catch {
        Write-Log "signatures.json 读取失败, 使用内置特征库: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# 3. 界面配置
# ---------------------------------------------------------------------------
function Load-Config {
    $script:ShowSystem = $false
    $script:ImmuneDefault = $false
    if (Test-Path -LiteralPath $script:ConfigFile) {
        try {
            $c = Get-Content -Raw -Encoding UTF8 -LiteralPath $script:ConfigFile | ConvertFrom-Json
            $script:ShowSystem = [bool]$c.ShowSystem
            $script:ImmuneDefault = [bool]$c.Immune
        } catch {}
    }
}
function Save-Config {
    try {
        $c = [ordered]@{ ShowSystem = [bool]$ChkShowSystem.IsChecked; Immune = [bool]$ChkImmune.IsChecked }
        [System.IO.File]::WriteAllText($script:ConfigFile, ($c | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($true)))
    } catch {}
}

# ---------------------------------------------------------------------------
# 4. 扫描位置定义
# ---------------------------------------------------------------------------
$script:ScanLocations = @(
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace';                                    Area = '此电脑';       Source = 'HKLM (64位)' }
    @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace';                        Area = '此电脑';       Source = 'HKLM (32位)' }
    @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace';                                    Area = '此电脑';       Source = 'HKCU' }
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace';                                       Area = '桌面';         Source = 'HKLM (64位)' }
    @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace';                           Area = '桌面';         Source = 'HKLM (32位)' }
    @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace';                                       Area = '桌面';         Source = 'HKCU' }
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\DelegateFolders';                    Area = '此电脑(委托)'; Source = 'HKLM (64位)' }
    @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\DelegateFolders';        Area = '此电脑(委托)'; Source = 'HKLM (32位)' }
    @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\DelegateFolders';                    Area = '此电脑(委托)'; Source = 'HKCU' }
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\DelegateFolders';                       Area = '桌面(委托)';   Source = 'HKLM (64位)' }
    @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\DelegateFolders';           Area = '桌面(委托)';   Source = 'HKLM (32位)' }
    @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\DelegateFolders';                       Area = '桌面(委托)';   Source = 'HKCU' }
)

# ---------------------------------------------------------------------------
# 5. 名称解析与分类
# ---------------------------------------------------------------------------
$script:IndirectAvailable = $false
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class IconCleanerWin32 {
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
    public static extern int SHLoadIndirectString(string pszSource, StringBuilder pszBuf, int cchBuf, IntPtr ppvReserved);
}
"@ -ErrorAction Stop
    $script:IndirectAvailable = $true
} catch {}

function Resolve-IndirectString([string]$Source) {
    if (-not $script:IndirectAvailable) { return '' }
    try {
        $sb = New-Object System.Text.StringBuilder 512
        $hr = [IconCleanerWin32]::SHLoadIndirectString($Source, $sb, 512, [IntPtr]::Zero)
        if ($hr -eq 0) { return $sb.ToString() }
    } catch {}
    return ''
}

function Resolve-ClsidName([string]$Clsid, [string]$Hint) {
    if (-not [string]::IsNullOrWhiteSpace($Hint)) { return $Hint.Trim() }
    $paths = @(
        "HKLM:\SOFTWARE\Classes\CLSID\$Clsid",
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\$Clsid",
        "HKCU:\Software\Classes\CLSID\$Clsid"
    )
    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p) {
            $k = Get-Item -LiteralPath $p
            $v = [string]$k.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
            $ls = [string]$k.GetValue('LocalizedString', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            if ($ls -like '@*') {
                $resolved = Resolve-IndirectString -Source $ls
                if (-not [string]::IsNullOrWhiteSpace($resolved)) { return $resolved }
            }
        }
    }
    return ''
}

function Get-Category([string]$Name, [string]$Clsid) {
    if ($script:WhiteList -contains $Clsid) { return '系统默认' }
    $n = $Name.ToLowerInvariant()
    if ($n -match '^clsid_') { return '系统默认' }
    if ($script:BuiltInSystemNames -contains $n.Trim()) { return '系统默认' }
    if ($n -match 'microsoft|windows|onedrive|outlook|visual studio|office') { return '微软' }
    foreach ($kw in $script:RogueKeywords) {
        if ($kw -and $n.Contains($kw.ToLowerInvariant())) { return '疑似流氓' }
    }
    foreach ($kw in $script:ThirdPartyKeywords) {
        if ($kw -and $n.Contains($kw.ToLowerInvariant())) { return '第三方' }
    }
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name -eq '(无法解析名称)') { return '未知' }
    return '第三方'
}

# ---------------------------------------------------------------------------
# 6. 扫描
# ---------------------------------------------------------------------------
function Get-IconEntries {
    $results = @()
    foreach ($loc in $script:ScanLocations) {
        if (-not (Test-Path -LiteralPath $loc.Path)) { continue }
        $children = Get-ChildItem -LiteralPath $loc.Path -ErrorAction SilentlyContinue
        foreach ($c in $children) {
            $guid = $c.PSChildName
            # 跳过非 GUID 的容器子键(如 DelegateFolders)
            $tmpGuid = [guid]::Empty
            if (-not [guid]::TryParse($guid, [ref]$tmpGuid)) { continue }
            $key  = Get-Item -LiteralPath $c.PSPath
            $rawDefault = [string]$key.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            $desc = [string]$key.GetValue('Description', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            $tip  = [string]$key.GetValue('InfoTip', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            $detail = (@($desc, $tip) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' | '
            if ([string]::IsNullOrWhiteSpace($detail)) { $detail = '(无说明)' }

            if ($rawDefault -eq $script:ImmuneMarker) {
                $name = $rawDefault
                $cat  = '已免疫'
                $detail = '该键已被本工具写入拒绝权限(免疫), 防止软件重新创建图标。删除此键即解除免疫。'
            } else {
                $name = Resolve-ClsidName -Clsid $guid -Hint $rawDefault
                if ([string]::IsNullOrWhiteSpace($name)) { $name = '(无法解析名称)' }
                $cat = Get-Category -Name $name -Clsid $guid
            }

            $results += [pscustomobject]@{
                Selected = ($cat -eq '疑似流氓')
                Category = $cat
                Name     = $name
                Clsid    = $guid
                Area     = $loc.Area
                Source   = $loc.Source
                RegPath  = (Join-Path $loc.Path $guid)
                Detail   = $detail
            }
        }
    }
    $rank = @{ '疑似流氓' = 0; '第三方' = 1; '未知' = 2; '已免疫' = 3; '微软' = 4; '系统默认' = 5 }
    return @($results | Sort-Object { $rank[[string]$_.Category] }, Name)
}

# ---------------------------------------------------------------------------
# 7. 注册表操作: 备份 / 恢复 / 权限修复 / 强制删除 / 免疫
# ---------------------------------------------------------------------------
function Export-RegTree([string]$ProviderPath) {
    if (-not (Test-Path -LiteralPath $ProviderPath)) { return $null }
    $key = Get-Item -LiteralPath $ProviderPath
    # 注意: 值名称可能为空字符串(默认值), 不能用作 JSON 键,
    # 因此 Values/Children 一律用数组结构保存
    $valuesList = @()
    foreach ($n in $key.GetValueNames()) {
        $kind = $key.GetValueKind($n)
        # DoNotExpandEnvironmentNames: 防止 ExpandString 值中的 %VAR% 被展开
        $data = $key.GetValue($n, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        $isNull = ($null -eq $data)
        if ($kind -eq [Microsoft.Win32.RegistryValueKind]::Binary -and -not $isNull) {
            $data = [Convert]::ToBase64String([byte[]]$data)
        }
        if ($kind -eq [Microsoft.Win32.RegistryValueKind]::MultiString -and -not $isNull) {
            $data = @($data)
        }
        if ($kind -eq [Microsoft.Win32.RegistryValueKind]::String -or $kind -eq [Microsoft.Win32.RegistryValueKind]::ExpandString) {
            $data = [string]$data
        }
        $valuesList += @{ Name = $n; Kind = $kind.ToString(); Data = $data; IsNull = $isNull }
    }
    $obj = [ordered]@{ Path = $ProviderPath; Values = $valuesList; Children = @() }
    foreach ($child in $key.GetSubKeyNames()) {
        $obj.Children += Export-RegTree -ProviderPath (Join-Path $ProviderPath $child)
    }
    return $obj
}

function Restore-RegTree($Tree) {
    if ($null -eq $Tree) { return }
    $path = [string]$Tree.Path
    $hive = if ($path -like 'HKLM:*') { [Microsoft.Win32.Registry]::LocalMachine } else { [Microsoft.Win32.Registry]::CurrentUser }
    $sub  = $path.Substring($path.IndexOf(':') + 1).TrimStart('\')
    $k = $hive.CreateSubKey($sub)
    foreach ($v in @($Tree.Values)) {
        if ($null -eq $v) { continue }
        $kindName = [string]$v.Kind
        $kind = [Microsoft.Win32.RegistryValueKind]$kindName
        if ($kind -eq [Microsoft.Win32.RegistryValueKind]::None -or $kind -eq [Microsoft.Win32.RegistryValueKind]::Unknown) {
            if ($v.IsNull) { $k.SetValue([string]$v.Name, '') } else { $k.SetValue([string]$v.Name, $v.Data) }
            continue
        }
        $data = $v.Data
        switch ($kindName) {
            'DWord'       { $data = [int]$data }
            'QWord'       { $data = [long]$data }
            'Binary'      { $data = [Convert]::FromBase64String([string]$data) }
            'MultiString' { $data = [string[]]@($data) }
            default       { $data = [string]$data }
        }
        if ($v.IsNull) { $data = '' }
        $k.SetValue([string]$v.Name, $data, $kind)
    }
    $k.Close()
    foreach ($child in @($Tree.Children)) {
        Restore-RegTree -Tree $child
    }
}

function Fix-RegKeyAcl([string]$ProviderPath) {
    if (-not (Test-Path -LiteralPath $ProviderPath)) { return $false }
    $hive = if ($ProviderPath -like 'HKLM:*') { [Microsoft.Win32.Registry]::LocalMachine } else { [Microsoft.Win32.Registry]::CurrentUser }
    $sub  = $ProviderPath.Substring($ProviderPath.IndexOf(':') + 1).TrimStart('\')
    try {
        $rights = [System.Security.AccessControl.RegistryRights]::TakeOwnership -bor [System.Security.AccessControl.RegistryRights]::ChangePermissions
        $k = $hive.OpenSubKey($sub, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, $rights)
        if (-not $k) { return $false }
        $me = [System.Security.Principal.NTAccount]([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
        $acl = $k.GetAccessControl()
        try { $acl.SetOwner($me); $k.SetAccessControl($acl) } catch {}
        $acl = $k.GetAccessControl()
        $acl.SetAccessRuleProtection($false, $false)
        foreach ($r in @($acl.GetAccessRules($true, $true, [System.Security.Principal.NTAccount]))) {
            if (-not $r.IsInherited) { [void]$acl.RemoveAccessRule($r) }
        }
        $grant = New-Object System.Security.AccessControl.RegistryAccessRule($me, 'FullControl', 'ContainerInherit', 'None', 'Allow')
        $acl.AddAccessRule($grant)
        $k.SetAccessControl($acl)
        $k.Close()
        return $true
    } catch {
        return $false
    }
}

function Remove-RegKeyForced([string]$ProviderPath) {
    if (-not (Test-Path -LiteralPath $ProviderPath)) { return $true }
    try {
        Remove-Item -LiteralPath $ProviderPath -Recurse -Force -ErrorAction Stop
        return $true
    } catch {
        Write-Log "  直接删除失败, 尝试修复权限后重试: $ProviderPath"
        if (Fix-RegKeyAcl -ProviderPath $ProviderPath) {
            try {
                Remove-Item -LiteralPath $ProviderPath -Recurse -Force -ErrorAction Stop
                return $true
            } catch {
                Write-Log "  权限修复后仍无法删除: $($_.Exception.Message)"
            }
        } else {
            Write-Log "  权限修复失败(可能被软件完全锁死), 可尝试进入安全模式删除"
        }
        return $false
    }
}

function Set-RegKeyImmune([string]$ProviderPath) {
    try {
        if (-not (Test-Path -LiteralPath $ProviderPath)) {
            New-Item -Path $ProviderPath -Force | Out-Null
        }
        $hive = if ($ProviderPath -like 'HKLM:*') { [Microsoft.Win32.Registry]::LocalMachine } else { [Microsoft.Win32.Registry]::CurrentUser }
        $sub  = $ProviderPath.Substring($ProviderPath.IndexOf(':') + 1).TrimStart('\')
        $rights = [System.Security.AccessControl.RegistryRights]::ChangePermissions -bor [System.Security.AccessControl.RegistryRights]::SetValue
        $k = $hive.OpenSubKey($sub, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, $rights)
        if (-not $k) { return $false }
        $k.SetValue('', $script:ImmuneMarker, [Microsoft.Win32.RegistryValueKind]::String)
        $acl = $k.GetAccessControl()
        $acl.SetAccessRuleProtection($false, $false)
        $denyRights = [System.Security.AccessControl.RegistryRights]::WriteKey -bor [System.Security.AccessControl.RegistryRights]::SetValue -bor [System.Security.AccessControl.RegistryRights]::CreateSubKey -bor [System.Security.AccessControl.RegistryRights]::Delete
        $deny = New-Object System.Security.AccessControl.RegistryAccessRule('Everyone', $denyRights, 'ContainerInherit', 'None', 'Deny')
        $acl.AddAccessRule($deny)
        $k.SetAccessControl($acl)
        $k.Close()
        return $true
    } catch {
        Write-Log "  免疫失败: $($_.Exception.Message)"
        return $false
    }
}

# ---------------------------------------------------------------------------
# 8. 其他功能: 备份文件 / 重启资源管理器 / 注册表定位
# ---------------------------------------------------------------------------
function Save-Backup($Rows) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $file = Join-Path $script:BackupDir "backup_$ts.json"
    $trees = @($Rows | ForEach-Object { Export-RegTree -ProviderPath $_.RegPath } | Where-Object { $_ })
    $backup = [ordered]@{
        Time    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Comment = '此电脑/桌面图标清理备份'
        Count   = $trees.Count
        Entries = $trees
    }
    [System.IO.File]::WriteAllText($file, ($backup | ConvertTo-Json -Depth 100), (New-Object System.Text.UTF8Encoding($true)))
    Write-Log "已备份 $($trees.Count) 项到: $file"
    return $file
}

function Restart-Explorer {
    Write-Log '正在重启资源管理器...'
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
    if (-not (Get-Process -Name explorer -ErrorAction SilentlyContinue)) {
        Start-Process explorer.exe
    }
    Write-Log '资源管理器已重启。'
}

function Open-RegeditAt([string]$ProviderPath) {
    try {
        $last = $ProviderPath
        $last = $last -replace '^HKLM:\\', '计算机\HKEY_LOCAL_MACHINE\'
        $last = $last -replace '^HKCU:\\', '计算机\HKEY_CURRENT_USER\'
        New-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit' -Force | Out-Null
        Set-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit' -Name LastKey -Value $last
        Start-Process regedit.exe
    } catch {}
}

# ---------------------------------------------------------------------------
# 9. 界面 XAML
# ---------------------------------------------------------------------------
$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="此电脑/桌面 流氓图标清理工具 v1.0"
        Height="700" Width="1060" MinHeight="520" MinWidth="860"
        WindowStartupLocation="CenterScreen"
        FontFamily="Microsoft YaHei UI, SimSun" FontSize="13">
    <DockPanel Margin="10">
        <WrapPanel DockPanel.Dock="Top" Margin="0,0,0,8">
            <Button Name="BtnScan" Content="重新扫描" Width="92" Margin="0,0,8,0" Padding="8,4"/>
            <Button Name="BtnCheckAll" Content="全选" Width="64" Margin="0,0,8,0" Padding="8,4"/>
            <Button Name="BtnCheckNone" Content="全不选" Width="64" Margin="0,0,8,0" Padding="8,4"/>
            <Button Name="BtnDelete" Content="删除选中项" Width="104" Margin="0,0,8,0" Padding="8,4" FontWeight="Bold"/>
            <Button Name="BtnRestore" Content="恢复备份..." Width="104" Margin="0,0,8,0" Padding="8,4"/>
            <Button Name="BtnExport" Content="导出报告..." Width="104" Margin="0,0,8,0" Padding="8,4"/>
            <Button Name="BtnExplorer" Content="重启资源管理器" Width="124" Margin="0,0,16,0" Padding="8,4"/>
            <CheckBox Name="ChkImmune" Content="免疫防复活" VerticalAlignment="Center" Margin="0,0,14,0"
                      ToolTip="删除后写入拒绝权限, 防止软件下次启动时重新创建图标"/>
            <CheckBox Name="ChkShowSystem" Content="显示系统项" VerticalAlignment="Center" Margin="0,0,14,0"
                      ToolTip="显示 Windows 自带的系统图标项(默认隐藏)"/>
            <TextBox Name="TxtSearch" Width="170" VerticalContentAlignment="Center"
                     ToolTip="按名称 / CLSID / 注册表路径筛选"/>
        </WrapPanel>
        <TextBlock DockPanel.Dock="Top" Margin="0,0,0,6" Foreground="#666666" TextWrapping="Wrap">
            扫描"此电脑"与"桌面"上的命名空间图标(CLSID 挂载项)。红色加粗为疑似流氓项(默认已勾选)，请核对后再删除；删除前自动备份，可随时通过"恢复备份"还原。双击某行或使用右键菜单可在注册表中定位。
        </TextBlock>
        <StatusBar DockPanel.Dock="Bottom">
            <StatusBarItem>
                <TextBlock Name="TxtStatus" Text="就绪"/>
            </StatusBarItem>
        </StatusBar>
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="170"/>
            </Grid.RowDefinitions>
            <ListView Name="LvIcons" Grid.Row="0" SelectionMode="Extended">
                <ListView.View>
                    <GridView>
                        <GridViewColumn Width="40">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                              HorizontalAlignment="Center"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="分类" Width="88" DisplayMemberBinding="{Binding Category}"/>
                        <GridViewColumn Header="名称" Width="230" DisplayMemberBinding="{Binding Name}"/>
                        <GridViewColumn Header="CLSID" Width="230" DisplayMemberBinding="{Binding Clsid}"/>
                        <GridViewColumn Header="位置" Width="120" DisplayMemberBinding="{Binding Area}"/>
                        <GridViewColumn Header="来源" Width="90" DisplayMemberBinding="{Binding Source}"/>
                        <GridViewColumn Header="注册表路径" Width="280" DisplayMemberBinding="{Binding RegPath}"/>
                    </GridView>
                </ListView.View>
                <ListView.ItemContainerStyle>
                    <Style TargetType="{x:Type ListViewItem}">
                        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding Category}" Value="疑似流氓">
                                <Setter Property="Foreground" Value="#C62828"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding Category}" Value="系统默认">
                                <Setter Property="Foreground" Value="#9E9E9E"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding Category}" Value="微软">
                                <Setter Property="Foreground" Value="#607D8B"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding Category}" Value="已免疫">
                                <Setter Property="Foreground" Value="#2E7D32"/>
                                <Setter Property="FontStyle" Value="Italic"/>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </ListView.ItemContainerStyle>
                <ListView.ContextMenu>
                    <ContextMenu>
                        <MenuItem Header="在注册表中定位"/>
                        <Separator/>
                        <MenuItem Header="复制 CLSID"/>
                        <MenuItem Header="复制注册表路径"/>
                        <MenuItem Header="复制条目完整信息"/>
                    </ContextMenu>
                </ListView.ContextMenu>
            </ListView>
            <Border Grid.Row="1" Margin="0,8,0,0" Padding="8" Background="#F5F5F5" CornerRadius="4">
                <TextBlock Name="TxtDetail" Text="未选中条目" TextWrapping="Wrap" Foreground="#333333"/>
            </Border>
            <GroupBox Grid.Row="2" Header="操作日志" Margin="0,8,0,0">
                <TextBox Name="TxtLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"/>
            </GroupBox>
        </Grid>
    </DockPanel>
</Window>
'@

$script:WpfLoaded = $false
function New-XamlWindow([string]$XamlText) {
    if (-not $script:WpfLoaded) {
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase
        $script:WpfLoaded = $true
    }
    $xr = New-Object System.Xml.XmlNodeReader ([xml]$XamlText)
    return [Windows.Markup.XamlReader]::Load($xr)
}

# ---------------------------------------------------------------------------
# 10. 自检(备份/删除/恢复/免疫 往返 + 界面 XAML 校验)
# ---------------------------------------------------------------------------
if ($SelfTest) {
    $base = 'HKCU:\Software\IconCleanerSelfTest'
    Write-Log '=== 自检开始 ==='
    try {
        if (Test-Path -LiteralPath $base) { [void](Remove-RegKeyForced -ProviderPath $base) }

        # --- 测试1: 备份/删除/恢复 往返(覆盖全部值类型) ---
        $k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\IconCleanerSelfTest\Parent')
        $k.SetValue('', '默认值测试', [Microsoft.Win32.RegistryValueKind]::String)
        $k.SetValue('Str', '中文字符串测试', [Microsoft.Win32.RegistryValueKind]::String)
        $k.SetValue('Exp', '%PATH% 扩展字符串', [Microsoft.Win32.RegistryValueKind]::ExpandString)
        $k.SetValue('Dw', 123456, [Microsoft.Win32.RegistryValueKind]::DWord)
        $k.SetValue('Qw', [long]12345678901234, [Microsoft.Win32.RegistryValueKind]::QWord)
        $k.SetValue('Bin', [byte[]](1, 2, 3, 4, 255), [Microsoft.Win32.RegistryValueKind]::Binary)
        $k.SetValue('Multi', [string[]]@('第一行', '第二行', 'third line'), [Microsoft.Win32.RegistryValueKind]::MultiString)
        $c = $k.CreateSubKey('Child')
        $c.SetValue('SubValue', '子键值', [Microsoft.Win32.RegistryValueKind]::String)
        $c.Close(); $k.Close()

        $tree = Export-RegTree -ProviderPath "$base\Parent"
        if ($null -eq $tree) { throw '导出失败(返回空)' }
        $back = $tree | ConvertTo-Json -Depth 100 | ConvertFrom-Json
        if (-not (Remove-RegKeyForced -ProviderPath "$base\Parent")) { throw '删除失败' }
        if (Test-Path -LiteralPath "$base\Parent") { throw '删除后键仍存在' }
        Restore-RegTree -Tree $back
        $k2 = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\IconCleanerSelfTest\Parent')
        if ($null -eq $k2) { throw '恢复后键不存在' }
        if ([string]$k2.GetValue('') -ne '默认值测试') { throw '默认值往返不一致' }
        if ([string]$k2.GetValue('Str') -ne '中文字符串测试') { throw '字符串往返不一致' }
        if ([string]$k2.GetValue('Exp', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) -ne '%PATH% 扩展字符串') { throw '扩展字符串往返不一致' }
        if ([int]$k2.GetValue('Dw') -ne 123456) { throw 'DWord 往返不一致' }
        if ([long]$k2.GetValue('Qw') -ne 12345678901234) { throw 'QWord 往返不一致' }
        if ([Convert]::ToBase64String([byte[]]$k2.GetValue('Bin')) -ne 'AQIDBP8=') { throw '二进制往返不一致' }
        if (([string[]]$k2.GetValue('Multi') -join '|') -ne '第一行|第二行|third line') { throw '多字符串往返不一致' }
        $k3 = $k2.OpenSubKey('Child')
        if ($null -eq $k3 -or [string]$k3.GetValue('SubValue') -ne '子键值') { throw '子键往返不一致' }
        if ($k3) { $k3.Close() }
        $k2.Close()
        Write-Log '[PASS] 测试1: 备份/删除/恢复 往返'

        # --- 测试2: 免疫 + 强制删除 ---
        New-Item -Path "$base\Immune" -Force | Out-Null
        if (-not (Set-RegKeyImmune -ProviderPath "$base\Immune")) { throw '免疫写入失败' }
        if (-not (Remove-RegKeyForced -ProviderPath "$base\Immune")) { throw '免疫键强制删除失败' }
        if (Test-Path -LiteralPath "$base\Immune") { throw '免疫键删除后仍存在' }
        Write-Log '[PASS] 测试2: 免疫 + 强制删除'

        # --- 测试3: 在免疫键上恢复(先解锁再恢复) ---
        $k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\IconCleanerSelfTest\Round2')
        $k.SetValue('V', '免疫恢复测试', [Microsoft.Win32.RegistryValueKind]::String)
        $k.Close()
        $t2 = Export-RegTree -ProviderPath "$base\Round2"
        [void](Remove-RegKeyForced -ProviderPath "$base\Round2")
        New-Item -Path "$base\Round2" -Force | Out-Null
        [void](Set-RegKeyImmune -ProviderPath "$base\Round2")
        [void](Fix-RegKeyAcl -ProviderPath "$base\Round2")
        Restore-RegTree -Tree ($t2 | ConvertTo-Json -Depth 100 | ConvertFrom-Json)
        $k4 = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\IconCleanerSelfTest\Round2')
        if ($null -eq $k4 -or [string]$k4.GetValue('V') -ne '免疫恢复测试') { throw '免疫键恢复失败' }
        if ($k4) { $k4.Close() }
        Write-Log '[PASS] 测试3: 免疫键解锁恢复'

        # --- 测试4: 界面 XAML 解析与控件齐全 ---
        $w = New-XamlWindow -XamlText $Xaml
        foreach ($n in @('BtnScan','BtnCheckAll','BtnCheckNone','BtnDelete','BtnRestore','BtnExport','BtnExplorer','ChkImmune','ChkShowSystem','TxtSearch','TxtStatus','LvIcons','TxtDetail','TxtLog')) {
            if ($null -eq $w.FindName($n)) { throw "XAML 控件未找到: $n" }
        }
        Write-Log '[PASS] 测试4: 界面 XAML 解析'

        [void](Remove-RegKeyForced -ProviderPath $base)
        Write-Log '=== 自检全部通过 ==='
        Write-Host 'SELFTEST: PASS'
        exit 0
    } catch {
        Write-Log "自检失败: $($_.Exception.Message)"
        Write-Host "SELFTEST: FAIL - $($_.Exception.Message)"
        try { [void](Remove-RegKeyForced -ProviderPath $base) } catch {}
        exit 1
    }
}

# ---------------------------------------------------------------------------
# 11. 命令行扫描模式
# ---------------------------------------------------------------------------
if ($ScanOnly) {
    Load-Signatures
    $entries = Get-IconEntries
    if (-not $entries -or $entries.Count -eq 0) {
        Write-Host '未发现命名空间图标项。'
        exit 0
    }
    Write-Host ("共发现 {0} 项:" -f $entries.Count)
    $entries | Select-Object Category, Name, Clsid, Area, Source, RegPath |
        Format-Table -AutoSize | Out-String -Width 400 | Write-Host
    Write-Host '提示: 双击 启动图标清理工具.bat 打开图形界面进行勾选删除。'
    exit 0
}

# ---------------------------------------------------------------------------
# 12. 管理员权限检查与自提权
# ---------------------------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`""
        exit 0
    }
}

# ---------------------------------------------------------------------------
# 13. 图形界面
# ---------------------------------------------------------------------------
Load-Signatures
Load-Config

try {
    $window = New-XamlWindow -XamlText $Xaml
} catch {
    Write-Host "界面加载失败: $($_.Exception.Message)"
    Write-Host '可改用命令行扫描: powershell -ExecutionPolicy Bypass -File IconCleaner.ps1 -ScanOnly'
    if ($Host.Name -eq 'ConsoleHost') { Read-Host '按回车退出' }
    exit 1
}

$BtnScan      = $window.FindName('BtnScan')
$BtnCheckAll  = $window.FindName('BtnCheckAll')
$BtnCheckNone = $window.FindName('BtnCheckNone')
$BtnDelete    = $window.FindName('BtnDelete')
$BtnRestore   = $window.FindName('BtnRestore')
$BtnExport    = $window.FindName('BtnExport')
$BtnExplorer  = $window.FindName('BtnExplorer')
$ChkImmune    = $window.FindName('ChkImmune')
$ChkShowSystem= $window.FindName('ChkShowSystem')
$TxtSearch    = $window.FindName('TxtSearch')
$TxtStatus    = $window.FindName('TxtStatus')
$TxtDetail    = $window.FindName('TxtDetail')
$TxtLog       = $window.FindName('TxtLog')
$LvIcons      = $window.FindName('LvIcons')
$script:LogBox = $TxtLog

$ChkShowSystem.IsChecked = $script:ShowSystem
$ChkImmune.IsChecked     = $script:ImmuneDefault

function Update-List {
    $q = $script:SearchText
    $all = @($script:AllRows)
    $items = @($all | Where-Object {
        (($script:ShowSystem) -or ($_.Category -ne '系统默认')) -and
        ([string]::IsNullOrEmpty($q) -or
         $_.Name.ToLowerInvariant().Contains($q.ToLowerInvariant()) -or
         $_.Clsid.ToLowerInvariant().Contains($q.ToLowerInvariant()) -or
         $_.RegPath.ToLowerInvariant().Contains($q.ToLowerInvariant()) -or
         $_.Area.Contains($q))
    })
    $LvIcons.ItemsSource = $items
    $selCount = @($all | Where-Object { $_.Selected }).Count
    $TxtStatus.Text = "共 $($all.Count) 项 (显示 $($items.Count) 项)  |  已勾选 $selCount 项"
}

function Start-Scan {
    Write-Log '开始扫描...'
    $script:AllRows = @(Get-IconEntries)
    Update-List
    $groups = @($script:AllRows | Group-Object Category | ForEach-Object { "$($_.Name) $($_.Count)" })
    $selCount = @($script:AllRows | Where-Object { $_.Selected }).Count
    $TxtStatus.Text = "共 $($script:AllRows.Count) 项  |  " + ($groups -join '  |  ') + "  |  已勾选 $selCount 项"
    Write-Log "扫描完成: 共 $($script:AllRows.Count) 项"
}

# ---- 事件绑定 ----
$BtnScan.Add_Click({ Start-Scan })

$BtnCheckAll.Add_Click({
    @($LvIcons.Items) | ForEach-Object { $_.Selected = $true }
    $LvIcons.ItemsSource = @($LvIcons.ItemsSource)
    Update-List
})
$BtnCheckNone.Add_Click({
    @($LvIcons.Items) | ForEach-Object { $_.Selected = $false }
    $LvIcons.ItemsSource = @($LvIcons.ItemsSource)
    Update-List
})

$BtnDelete.Add_Click({
    $sel = @($LvIcons.Items | Where-Object { $_.Selected })
    if ($sel.Count -eq 0) {
        [void][System.Windows.MessageBox]::Show($window, '请先在列表中勾选要删除的条目。', '提示', 'OK', 'Information')
        return
    }
    $lines  = ($sel | ForEach-Object { "  * $($_.Name)   [$($_.RegPath)]" }) -join "`n"
    $sysSel = @($sel | Where-Object { $_.Category -in @('系统默认', '微软') })
    $immSel = @($sel | Where-Object { $_.Category -eq '已免疫' })
    $warn = ''
    if ($sysSel.Count -gt 0) { $warn += "`n`n【警告】其中包含 $($sysSel.Count) 个系统/微软项, 删除可能影响系统功能!" }
    if ($immSel.Count -gt 0) { $warn += "`n`n【提示】其中包含 $($immSel.Count) 个已免疫项, 删除它们等于解除免疫。" }
    $immNote = ''
    if ($ChkImmune.IsChecked) { $immNote = "`n`n本次删除将同时启用【免疫防复活】(写入拒绝权限, 防止软件重建图标)。" }
    $r = [System.Windows.MessageBox]::Show($window, "即将处理以下 $($sel.Count) 个注册表项:`n`n$lines`n`n删除前会自动备份, 之后可通过【恢复备份】还原。$immNote$warn`n`n确定继续吗?", '删除确认', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }

    $null = Save-Backup -Rows $sel
    $okCount = 0; $failCount = 0
    foreach ($row in $sel) {
        if (Remove-RegKeyForced -ProviderPath $row.RegPath) {
            $okCount++
            Write-Log "已删除: $($row.Name)  [$($row.RegPath)]"
            if ($ChkImmune.IsChecked) {
                if (Set-RegKeyImmune -ProviderPath $row.RegPath) {
                    Write-Log "  已免疫: $($row.RegPath)"
                } else {
                    Write-Log "  免疫失败(图标可能被软件重建): $($row.RegPath)"
                }
            }
        } else {
            $failCount++
            Write-Log "删除失败: $($row.Name)  [$($row.RegPath)]"
        }
    }
    Start-Scan
    $ask = [System.Windows.MessageBox]::Show($window, "处理完成: 成功 $okCount 项, 失败 $failCount 项。`n是否立即重启资源管理器, 让更改立即生效?`n(注意: 会关闭所有已打开的资源管理器窗口)", '重启资源管理器', 'YesNo', 'Question')
    if ($ask -eq 'Yes') { Restart-Explorer }
})

$BtnRestore.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = '选择备份文件'
    $dlg.Filter = '备份文件 (*.json)|*.json|所有文件 (*.*)|*.*'
    if (Test-Path -LiteralPath $script:BackupDir) { $dlg.InitialDirectory = $script:BackupDir }
    if ($dlg.ShowDialog() -ne $true) { return }
    try {
        $backup = [System.IO.File]::ReadAllText($dlg.FileName, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    } catch {
        [void][System.Windows.MessageBox]::Show($window, "读取备份文件失败:`n$($_.Exception.Message)", '错误', 'OK', 'Error')
        return
    }
    $entries = @($backup.Entries | Where-Object { $_ })
    if ($entries.Count -eq 0) {
        [void][System.Windows.MessageBox]::Show($window, '该备份文件中没有可恢复的条目。', '提示', 'OK', 'Information')
        return
    }
    $r = [System.Windows.MessageBox]::Show($window, "备份时间: $($backup.Time)`n说明: $($backup.Comment)`n`n将恢复 $($entries.Count) 个注册表项。`n若对应键已开启免疫, 恢复时会自动解除权限限制。`n`n确定恢复吗?", '恢复确认', 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    $okCount = 0; $failCount = 0
    foreach ($tree in $entries) {
        try {
            if (Test-Path -LiteralPath $tree.Path) { [void](Fix-RegKeyAcl -ProviderPath $tree.Path) }
            Restore-RegTree -Tree $tree
            $okCount++
            Write-Log "已恢复: $($tree.Path)"
        } catch {
            $failCount++
            Write-Log "恢复失败: $($tree.Path) -> $($_.Exception.Message)"
        }
    }
    [void][System.Windows.MessageBox]::Show($window, "恢复完成: 成功 $okCount 项, 失败 $failCount 项。", '完成', 'OK', 'Information')
    Start-Scan
})

$BtnExport.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Title = '导出扫描报告'
    $dlg.Filter = 'CSV 文件 (*.csv)|*.csv'
    $dlg.FileName = '图标扫描报告_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv'
    if ($dlg.ShowDialog() -ne $true) { return }
    $rows = @($script:AllRows | Select-Object @{n='分类';e={$_.Category}}, @{n='名称';e={$_.Name}}, @{n='CLSID';e={$_.Clsid}}, @{n='位置';e={$_.Area}}, @{n='来源';e={$_.Source}}, @{n='注册表路径';e={$_.RegPath}}, @{n='说明';e={$_.Detail}})
    $csv = $rows | ConvertTo-Csv -NoTypeInformation | Out-String
    [System.IO.File]::WriteAllText($dlg.FileName, $csv, (New-Object System.Text.UTF8Encoding($true)))
    Write-Log "报告已导出: $($dlg.FileName)"
    [void][System.Windows.MessageBox]::Show($window, '导出成功。', '完成', 'OK', 'Information')
})

$BtnExplorer.Add_Click({
    $r = [System.Windows.MessageBox]::Show($window, '确定重启资源管理器吗? 会关闭所有已打开的资源管理器窗口。', '重启资源管理器', 'YesNo', 'Question')
    if ($r -eq 'Yes') { Restart-Explorer }
})

$TxtSearch.Add_TextChanged({
    $script:SearchText = $TxtSearch.Text
    Update-List
})

$ChkShowSystem.Add_Checked({ $script:ShowSystem = $true; Update-List })
$ChkShowSystem.Add_Unchecked({ $script:ShowSystem = $false; Update-List })

$LvIcons.Add_SelectionChanged({
    $it = $LvIcons.SelectedItem
    if ($it) {
        $TxtDetail.Text = "名称: $($it.Name)`n分类: $($it.Category)`nCLSID: $($it.Clsid)`n位置: $($it.Area) ($($it.Source))`n注册表: $($it.RegPath)`n说明: $($it.Detail)"
    } else {
        $TxtDetail.Text = '未选中条目'
    }
})

$LvIcons.Add_MouseDoubleClick({
    if ($LvIcons.SelectedItem) { Open-RegeditAt -ProviderPath $LvIcons.SelectedItem.RegPath }
})

$cm = $LvIcons.ContextMenu
$cm.Items[0].Add_Click({ if ($LvIcons.SelectedItem) { Open-RegeditAt -ProviderPath $LvIcons.SelectedItem.RegPath } })
$cm.Items[2].Add_Click({ if ($LvIcons.SelectedItem) { try { Set-Clipboard -Value $LvIcons.SelectedItem.Clsid } catch {} } })
$cm.Items[3].Add_Click({ if ($LvIcons.SelectedItem) { try { Set-Clipboard -Value $LvIcons.SelectedItem.RegPath } catch {} } })
$cm.Items[4].Add_Click({
    if ($LvIcons.SelectedItem) {
        $it = $LvIcons.SelectedItem
        try { Set-Clipboard -Value "名称: $($it.Name)`n分类: $($it.Category)`nCLSID: $($it.Clsid)`n位置: $($it.Area) ($($it.Source))`n注册表: $($it.RegPath)`n说明: $($it.Detail)" } catch {}
    }
})

$window.Add_Closed({ Save-Config })

# ---- 启动 ----
Write-Log '工具已启动。'
Start-Scan
$null = $window.ShowDialog()
Write-Log '工具已退出。'
