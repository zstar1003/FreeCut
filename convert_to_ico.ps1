# 从SVG文件中提取base64 PNG数据并转换为ICO
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# 读取SVG文件
$svgContent = Get-Content "ui\logo.svg" -Raw

# 提取base64数据（从data:img/png;base64,后面的部分）
$base64Pattern = 'xlink:href="data:img/png;base64,([^"]+)"'
if ($svgContent -match $base64Pattern) {
    $base64Data = $matches[1]

    # 解码base64为字节数组
    $imageBytes = [Convert]::FromBase64String($base64Data)

    # 创建内存流并加载图像
    $ms = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
    $ms.Write($imageBytes, 0, $imageBytes.Length)
    $image = [System.Drawing.Image]::FromStream($ms, $true)

    # 创建多个尺寸的图标（标准ICO格式包含多个尺寸）
    $sizes = @(16, 32, 48, 64, 128, 256)
    $iconStream = New-Object System.IO.MemoryStream
    $writer = New-Object System.IO.BinaryWriter($iconStream)

    # 写入ICO文件头
    $writer.Write([UInt16]0)  # 保留字段
    $writer.Write([UInt16]1)  # 图像类型（1=ICO）
    $writer.Write([UInt16]$sizes.Count)  # 图像数量

    # 准备图像数据
    $imageDataList = @()
    foreach ($size in $sizes) {
        $bitmap = New-Object System.Drawing.Bitmap($size, $size)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.DrawImage($image, 0, 0, $size, $size)

        $pngStream = New-Object System.IO.MemoryStream
        $bitmap.Save($pngStream, [System.Drawing.Imaging.ImageFormat]::Png)
        $pngBytes = $pngStream.ToArray()

        $imageDataList += ,@{
            Size = $size
            Data = $pngBytes
        }

        $graphics.Dispose()
        $bitmap.Dispose()
        $pngStream.Dispose()
    }

    # 写入图像目录
    $offset = 6 + (16 * $sizes.Count)  # 头部 + 目录条目
    foreach ($imageData in $imageDataList) {
        $size = $imageData.Size
        $data = $imageData.Data

        # 宽度（256表示为0）
        if ($size -eq 256) {
            $writer.Write([Byte]0)
        } else {
            $writer.Write([Byte]$size)
        }

        # 高度（256表示为0）
        if ($size -eq 256) {
            $writer.Write([Byte]0)
        } else {
            $writer.Write([Byte]$size)
        }

        $writer.Write([Byte]0)   # 调色板颜色数
        $writer.Write([Byte]0)   # 保留字段
        $writer.Write([UInt16]1) # 颜色平面数
        $writer.Write([UInt16]32) # 位深度
        $writer.Write([UInt32]$data.Length) # 图像数据大小
        $writer.Write([UInt32]$offset) # 图像数据偏移

        $offset += $data.Length
    }

    # 写入图像数据
    foreach ($imageData in $imageDataList) {
        $writer.Write($imageData.Data)
    }

    # 保存为icon.ico
    $writer.Flush()
    $icoBytes = $iconStream.ToArray()
    [System.IO.File]::WriteAllBytes("icon.ico", $icoBytes)

    # 清理资源
    $writer.Close()
    $iconStream.Close()
    $ms.Close()
    $image.Dispose()

    Write-Host "成功创建 icon.ico 文件！"
} else {
    Write-Host "错误：无法从SVG文件中提取base64数据"
    exit 1
}
