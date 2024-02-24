
function Compare-Version {
    <#
    .SYNOPSIS
        比對2個版本, $version1 大於 $version2 回傳$Ture , 等於或小於回傳$False
    .DESCRIPTION
        函數的詳細描述
    #>
    param (
      [Parameter(Mandatory = $true)]
      [string]$Version1, # 第一個版本
  
      [Parameter(Mandatory = $true)]
      [string]$Version2     # 第二個版本
    )
  
    # 將版本號拆分成陣列，以便逐個比較各個部分
    $version1Array = $Version1.Split('.')
    $version2Array = $Version2.Split('.')
  
    # 使用 foreach 迴圈遍歷每個部分進行比較
    foreach ($i in 0..$version1Array.Count) {
      if ([int]$version1Array[$i] -gt [int]$version2Array[$i]) {
        return $true    # 返回 $true 表示第一個版本號大於第二個版本號
      }
      elseif ([int]$version1Array[$i] -lt [int]$version2Array[$i]) {
        return $false   # 返回 $false 表示第一個版本號小於第二個版本號
      }
      else {
        # 如果當前部分相等，則繼續比較下一個部分
        continue
      }
    }
  
    # 如果完全相同，則表示版本號相同
    return $false    # 返回 $true 表示兩個版本號相同
  }

  
 