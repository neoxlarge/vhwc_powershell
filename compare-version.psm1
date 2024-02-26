
function Compare-Version {
    <#
    .SYNOPSIS
        ���2�Ӫ���, $version1 �j�� $version2 �^��$Ture , ����Τp��^��$False
    .DESCRIPTION
        ��ƪ��ԲӴy�z
    #>
    param (
      [Parameter(Mandatory = $true)]
      [string]$Version1, # �Ĥ@�Ӫ���
  
      [Parameter(Mandatory = $true)]
      [string]$Version2     # �ĤG�Ӫ���
    )
  
    # �N������������}�C�A�H�K�v�Ӥ���U�ӳ���
    $version1Array = $Version1.Split('.')
    $version2Array = $Version2.Split('.')
  
    # �ϥ� foreach �j��M���C�ӳ����i����
    foreach ($i in 0..$version1Array.Count) {
      if ([int]$version1Array[$i] -gt [int]$version2Array[$i]) {
        return $true    # ��^ $true ��ܲĤ@�Ӫ������j��ĤG�Ӫ�����
      }
      elseif ([int]$version1Array[$i] -lt [int]$version2Array[$i]) {
        return $false   # ��^ $false ��ܲĤ@�Ӫ������p��ĤG�Ӫ�����
      }
      else {
        # �p�G��e�����۵��A�h�~�����U�@�ӳ���
        continue
      }
    }
  
    # �p�G�����ۦP�A�h��ܪ������ۦP
    return $false    # ��^ $true ��ܨ�Ӫ������ۦP
  }

  
 