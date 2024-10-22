$mutexName = "Global\dutycheck-printers"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)

if ($mutex.WaitOne(0, $false) -eq $false) { 
    Write-Host "印表機檢查通知己在執行中,結束."
    exit 
}

write-host "灣橋印表機檢查通知Line notify"

# 1. 檢查排程 每天8點 和下午2點 檢查 L5100DN 和 TSC barcode
# 2. 從web介面取得印表機狀況, 如果不是以下狀熊就發通知.
#    - L5100DN normal_status = @("Sleep", "Deep Sleep", "Ready","No Paper T1", "No Paper T2", "Printing", "Please Wait","No Paper MP")
#    - TC barcode normal_status = @('Ready')
# 3. 重點印表機檢查排程, 每天8點到17點, 每0,15,30,45分, 檢查有always_on = $true的印表機.
#    例如急診,ICU,M5A等
# 4. 重點印表機, 必須在線, 網路連不上也會通知.

# 設定值
# debug log的資訊, Continue會顯示, SilentContinue不會顯示.
$DebugPreference = "Continue"

# line notify token
$line_apikey = "XkxO98qPwgpqoYQXsSsoSu94yHGA0TV9pZSVRkeZpqk"

# 定時的時間
$timer_hours = @(8..17) #8點到17點
$timer_minutes = @(0, 15, 30, 45)

# printer ip table
$L5100DNs = @{

    'wadm-mrr-pc02'    = @{'ip' = '172.20.2.253'
        'location'           = '病歷室'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ''
    }

    'wmed-com-pr01'    = @{'ip' = '172.20.2.194'
        'location'           = '社區營造'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ''
    }

    'wnur-csr-pr01'    = @{'ip' = '172.20.2.196'
        'location'           = '供應中心'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ''
    }


    'wnur-opd-pr01'    = @{'ip' = '172.20.9.21'
        'location'           = '診間101'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ''
    }

    'wnur-opd-pr02'    = @{'ip' = '172.20.9.22'
        'location'           = '診間103'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ''
    }

    'wnur-opd-pr03'    = @{'ip' = '172.20.9.23'
        'location'           = '診間105'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'C2EmRMB8'
    }
    
    'wnur-opd-pr04'    = @{'ip' = '172.20.9.24'
        'location'           = '診間110 婦產科'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = '&-:NWm45'
    }
                    
    'wnur-opd-pr05'    = @{'ip' = '172.20.9.25'
        'location'           = '診間109 骨科'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'K1kS6kre'
    }                     

    'wnur-opd-pr06'    = @{'ip' = '172.20.9.40'
        'location'           = '診間106 兒科'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = '4@M3E&Yi'
    }          
                    
    'wnur-opd-pr17'    = @{'ip' = '172.20.9.37'
        'location'           = '診間107 眼科'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'aCNhDdRF'
    }   

    'XXXwnur-opd-pr07' = @{'ip' = '172.20.9.27'
        'location'              = '診間10XXX 眼科'
        'password_vhwc'         = 'Us2791072'
        'password_factroy'      = ''
    } 

    'wnur-opd-pr08'    = @{'ip' = '172.20.9.28'
        'location'           = '診間108 精神科'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ''
        
    } 
 
    'wnur-opd-pra1'    = @{'ip' = '172.20.9.54'
        'location'           = '診間102 注射室'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ':WBBC2/V'
        
    }      

    'wadm-nhi-pr02'    = @{'ip' = '172.20.3.104'
        'location'           = '醫企室'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'a1LgVwid'
    }

    'wnur-lng-pr02'    = @{'ip' = '172.20.3.103'
        'location'           = '長照A據點'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = '5FL/DfU&'
    }
                    
    'wnur-lng-pr11'    = @{'ip' = '172.20.3.12'
        'location'           = '失智A據點'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'rQS?x/hS'
    }

    'wpha-sto-pr01'    = @{'ip' = '172.20.9.87'
        'location'           = '藥劑科藥庫'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'N+K$*sOc'

    }

    'wpha-pha-pr02'    = @{'ip' = '172.20.9.82'
        'location'           = '藥劑科中醫'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'TnUR%AUs'
        'always_on'          = $true
    }

    'wnur-erx-pr03'    = @{'ip' = '172.20.3.113'
        'location'           = '急診室 外側'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = '>aQxdNMd'
        'always_on'          = $true
    }

    'wnur-erx-pr02'    = @{'ip' = '172.20.3.44'
        'location'           = '急診室 內側'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'f1C7xdaT'
        'always_on'          = $true
    }                    

    'wmed-msh-pr01'    = @{'ip' = '172.20.3.253'
        'location'           = '醫療部辦公室'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'F1C7xdaT'
    }       
                    
    'wmed-msh-pr02'    = @{'ip' = '172.20.7.41'
        'location'           = '醫療部辦公室'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = '2mF1Afk$M'
    }   

    'wnur-hca-pc01'    = @{'ip' = '172.20.7.62'
        'location'           = '居家護理孟言'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'eN8Wn&Pa'
    }       
                    
    'wpsy-phc-pr01'    = @{'ip' = '172.20.3.219'
        'location'           = '精神部居家'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'NAW0j6aL'
    } 
                    
    'wpsy-psy-pr01'    = @{'ip' = '172.20.3.231'
        'location'           = '精神部辦公室'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ':x9@N#*m'
    }                     

    'wnur-icu-pr01'    = @{'ip' = '172.20.5.63'
        'location'           = 'ICU'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = '9>X-G2a&'
        'always_on'          = $true
    }
                    
    'wnur-icu-pr02'    = @{'ip' = '172.20.5.30'
        'location'           = 'ICU'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'FQK>1Ncx'
        'always_on'          = $true
    }                    


    'wnur-m3w-pr01'    = @{'ip' = '172.20.5.55'
        'location'           = 'M3'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ':Y@u2W7X'
    }    

    'wnur-m3w-pr02'    = @{'ip' = '172.20.5.64'
        'location'           = 'M3'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = '8#5s31u@'
    } 

    'wnur-orw-pr01'    = @{'ip' = '172.20.5.29'
        'location'           = '開刀房'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = ''
    } 

    'wnur-m5a-pr01'    = @{'ip' = '172.20.5.60'
        'location'           = 'M5A'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'i+3Rx7fs'
        'always_on'          = $true
    }    

    'wnur-m5a-pr02'    = @{'ip' = '172.20.5.61'
        'location'           = 'M5A'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = '#GhYeVTg'
        'always_on'          = $true
    }    


    'wnur-m5b-pr01'    = @{'ip' = '172.20.5.36'
        'location'           = 'M5B'
        'password_vhwc'      = 'Us2791072'
        'password_factroy'   = 'YC5@r>*p'
        'always_on'          = $true
    }                        

    'wmis-000-pr06'    = @{'ip' = '172.20.5.158'
        'location'           = '6F資訊室'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
    }       

    'wnur-opd-pr21'    = @{'ip' = '172.20.12.201'
        'location'           = '診間201 牙科'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = '0hgYu&ux'
    }                 
    
    'wnur-opd-pr22'    = @{'ip' = '172.20.12.202'
        'location'           = '診間202'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
    }  

    'wnur-opd-pr23'    = @{'ip' = '172.20.12.203'
        'location'           = '診間203'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
    }  

    'wreh-000-pr01'    = @{'ip' = '172.20.17.224'
        'location'           = '復建科櫃台'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = 'gF$g$Gec'
        
    }            

    'wreh-000-pr02'    = @{'ip' = '172.20.17.200'
        'location'           = '診間復建科'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = 'L-%KfwCe'
    }      
                    
    'wnur-a1w-pr01'    = @{'ip' = '172.20.17.69'
        'location'           = 'A1'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  
                    
    'wnur-a1w-pr02'    = @{'ip' = '172.20.17.201'
        'location'           = 'A1'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-a2w-pr01'    = @{'ip' = '172.20.17.70'
        'location'           = 'A2'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-a2w-pr02'    = @{'ip' = '172.20.17.202'
        'location'           = 'A2'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-a3w-pr01'    = @{'ip' = '172.20.17.71'
        'location'           = 'A3'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-a3w-pr02'    = @{'ip' = '172.20.17.203'
        'location'           = 'A3'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-a5w-pr01'    = @{'ip' = '172.20.17.72'
        'location'           = 'A5'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-a5w-pr02'    = @{'ip' = '172.20.17.205'
        'location'           = 'A5'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-b1w-pr01'    = @{'ip' = '172.20.2.121'
        'location'           = 'B1'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-b1w-pr02'    = @{'ip' = '172.20.2.100'
        'location'           = 'B1'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-b2w-pr01'    = @{'ip' = '172.20.2.122'
        'location'           = 'B2'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-b2w-pr02'    = @{'ip' = '172.20.2.101'
        'location'           = 'B2'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-b3w-pr01'    = @{'ip' = '172.20.2.123'
        'location'           = 'B3'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-b3w-pr02'    = @{'ip' = '172.20.2.102'
        'location'           = 'B3'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
        'always_on'          = $true
    }  

    'wnur-b5w-pr01'    = @{'ip' = '172.20.2.124'
        'location'           = 'B5'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = '3dkj43F7'
        'always_on'          = $true
    }  

    'wnur-b5w-pr02'    = @{'ip' = '172.20.2.104'
        'location'           = 'B5'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = 'uGif4Bpn'
        'always_on'          = $true
    }  

    'wdie-out-pr01'    = @{'ip' = '172.20.2.149'
        'location'           = '營養科'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
    }      
                    
    'wsec-ele-pr01'    = @{'ip' = '172.20.2.183'
        'location'           = '機電班'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
    }  
    
    'wnur-000-pr01'    = @{'ip' = '172.20.3.196'
        'location'           = '護理部辦公室憶玲'
        'password_vhwc'      = 'Us2791072'
        'password_factory'   = ''
    }  

}

$tc200s = @{

    'wnur-opd-pdb6' = @{
        'ip'       = '172.20.9.41'
        'location' = '診間106'
    }

    'wnur-erx-prb1' = @{
        'ip'        = '172.20.3.107'
        'location'  = '急診室'
        'always_on' = $true
    }

    'wmis-000-prb1' = @{
        'ip'       = '172.20.5.177'
        'location' = '6F資訊室'
        
    }

    'wnur-m5b-prb1' = @{
        'ip'        = '172.20.5.41'
        'location'  = 'M5B'
        'always_on' = $true
    }

    'wnur-m5a-prb1' = @{
        'ip'        = '172.20.5.40'
        'location'  = 'M5A'
        'always_on' = $true
    }

    'wnur-icu-prb1' = @{
        'ip'        = '172.20.5.42'
        'location'  = 'ICU'
        'always_on' = $true
    }

    'wnur-m3w-prb1' = @{
        'ip'       = '172.20.5.43'
        'location' = 'M3'
    }

    'wnur-a1w-ba01' = @{
        'ip'        = '172.20.17.211'
        'location'  = 'A1'
        'always_on' = $true
    }

    'wnur-a2w-prb1' = @{
        'ip'        = '172.20.17.212'
        'location'  = 'A2'
        'always_on' = $true
    }

    'wnur-opd-dp07' = @{
        'ip'        = '172.20.17.213'
        'location'  = 'A3'
        'always_on' = $true
    }

    'wnur-a5w-ba01' = @{
        'ip'        = '172.20.17.215'
        'location'  = 'A5'
        'always_on' = $true
    }

    'wnur-b1w-prb1' = @{
        'ip'        = '172.20.2.119'
        'location'  = 'B1'
        'always_on' = $true
    }
    
    'wnur-b2w-prb1' = @{
        'ip'        = '172.20.2.114'
        'location'  = 'B2'
        'always_on' = $true
    }

    'wnur-b3w-prb1' = @{
        'ip'        = '172.20.2.116'
        'location'  = 'B3'
        'always_on' = $true
    }

    'wnur-b5w-prb1' = @{
        'ip'        = '172.20.2.117'
        'location'  = 'B5'
        'always_on' = $true
    }
}


# 分散式列印的客戶端
# firewall裡有些開9080是錯的, 實際上是2788
$idsm_clients = @{
    'wmis-000-pc05'     = @{
        'ip'       = '172.20.5.185'
        'port'     = '2788' 
        'location' = '資訊室測試'
    }

    'wpha-pha-pc08'     = @{
        'ip'        = '172.20.9.78'
        'port'      = '2788'
        'location'  = '藥劑科 藥袋列印'
        'always_on' = $true
    }

    'wnur-a1w-pc04'     = @{
        'ip'        = '172.20.17.14'
        'port'      = '2788'
        'location'  = 'A1'
        'always_on' = $true
    }

    'wnur-a2w-pc04'     = @{
        'ip'        = '172.20.17.24'
        'port'      = '2788'
        'location'  = 'A2'
        'always_on' = $true
    }

    'wnur-a3w-pc05'     = @{
        'ip'        = '172.20.17.35'
        'port'      = '2788'
        'location'  = 'A3'
        'always_on' = $true
    }

    'wnur-a5w-pc02'     = @{
        'ip'        = '172.20.17.52'
        'port'      = '2788'
        'location'  = 'A5'
        'always_on' = $true
    }

    'wnur-b1w-pc04'     = @{
        'ip'        = '172.20.2.93'
        'port'      = '2788'
        'location'  = 'B1'
        'always_on' = $true
    }

    'wnur-b2w-pc05'     = @{
        'ip'        = '172.20.2.94'
        'port'      = '2788'
        'location'  = 'B2'
        'always_on' = $true
    }

    'wnur-b3w-pc04'     = @{
        'ip'        = '172.20.2.97'
        'port'      = '2788'
        'location'  = 'B3'
        'always_on' = $true
    }

    'wnur-b5w-pc05'     = @{
        'ip'        = '172.20.2.77' #109
        'port'      = '2788'
        'location'  = 'B5'
        'always_on' = $true
    }

    'wadm-mrr-pc02'     = @{
        'ip'       = '172.20.2.207'
        'port'     = '2788'
        'location' = '病歷室'
    }

    'wnur-erx-pc02'     = @{
        'ip'        = '172.20.3.3'
        'port'      = '2788'
        'location'  = '急診室'
        'always_on' = $true
    }

    'wnur-icu-pc01'     = @{
        'ip'        = '172.20.5.2'
        'port'      = '2788'
        'location'  = 'ICU'
        'always_on' = $true
    }

    'wnur-m3w-pc05'     = @{
        'ip'        = '172.20.5.31'
        'port'      = '2788'
        'location'  = 'M3'
        'always_on' = $true
    }

    'wnur-m5a-pc06'     = @{
        'ip'        = '172.20.5.26'
        'port'      = '2788'
        'location'  = 'M5A'
        'always_on' = $true
    }

    'wnur-m5b-pc06'     = @{
        'ip'        = '172.20.5.15'
        'port'      = '2788'
        'location'  = 'M5B'
        'always_on' = $true
    }

    'wlab-000-pc03'     = @{
        'ip'        = '172.20.7.18'
        'port'      = '2788'
        'location'  = '檢驗科血庫'
        'always_on' = $true
    }

    'wlab-000-pc06a'    = @{
        'ip'        = '172.20.3.211'
        'port'      = '2788'
        'location'  = '檢驗科'
        'always_on' = $true
    }

    <# 儀器連線用電腦, 目前應該沒有用到分散式列印, 暫時拿掉
    'wlab-000-pc09' = @{
        'ip'   = '172.20.3.149'
        'port' = '2788'
        'location' = '檢驗科 收信儀器連線'
    }
    #>
    
    'wadm-reg-pc04'     = @{
        'ip'       = '172.20.3.29'
        'port'     = '2788'
        'location' = '掛號室 代收嘉榮費用'
    }

    'wpsy-pnp-pc01'     = @{
        'ip'       = '172.20.3.158'
        'port'     = '2788'
        'location' = '精神科'
    }

    'wadm-reg-pc01'     = @{
        'ip'       = '172.20.3.77'
        'port'     = '2788'
        'location' = '掛號室 曉婷'
    }
    
    'wadm-reg-pc02'     = @{
        'ip'       = '172.20.3.1'
        'port'     = '2788'
        'location' = '掛號室 玉尊'
    }

    'wadm-reg-pc03'     = @{
        'ip'       = '172.20.3.13'
        'port'     = '2788'
        'location' = '掛號室 舒璇'
    }

    'wadm-reg-pc04xxxx' = @{
        'ip'       = '172.20.3.120'
        'port'     = '2788'
        'location' = '掛號室 第4櫃台'
    }

    'wpha-pha-pc06'     = @{
        'ip'        = '172.20.9.76'
        'port'      = '2788'
        'location'  = '藥劑科 藥局發藥櫃台'
        'always_on' = $true
    }

    'wreh-000-pc03'     = @{
        'ip'       = '172.20.17.63'
        'port'     = '2788'
        'location' = '復建科 櫃台'
    }
    
    'wdie-out-pc01'     = @{
        'ip'        = '172.20.2.138'
        'port'      = '2788'
        'location'  = '營養室外 餐卡'
        'always_on' = $true
    }
    
    'wdie-out-pc02'     = @{
        'ip'        = '172.20.2.150'
        'port'      = '2788'
        'location'  = '營養室 出餐單'
        'always_on' = $true
    }
}    


function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz", # Line Notify 存取權杖

        [Parameter(Mandatory = $true)]
        [string]$Message, # 要發送的訊息內容

        [string]$StickerPackageId, # 要一併傳送的貼圖套件 ID

        [string]$StickerId              # 要一併傳送的貼圖 ID
    )

    # Line Notify API 的 URI
    $uri = "https://notify-api.line.me/api/notify"

    # 設定 HTTP Header，包含 Line Notify 存取權杖
    $headers = @{ "Authorization" = "Bearer $Token" }

    # 設定要傳送的訊息內容
    $payload = @{
        "message" = $Message
    }

    # 如果要傳送貼圖，加入貼圖套件 ID 和貼圖 ID
    if ($StickerPackageId -and $StickerId) {
        $payload["stickerPackageId"] = $StickerPackageId
        $payload["stickerId"] = $StickerId
    }

    try {
        # 使用 Invoke-RestMethod 傳送 HTTP POST 請求
        $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payload
        
        # 訊息成功傳送
        Write-Debug "訊息已成功傳送。"
    }
    catch {
        # 發生錯誤，輸出錯誤訊息
        Write-Error $_.Exception.Message
    }
}


function Schedulecheck-L5100DN {

    param(
        $printers
    )

    # device status: 異常狀態
    # 底下為廠商印出的貼紙上所列的異常
    $warning_status = @("Replace Drum", "Drum End Soon", "No Drum Unit",
        "Toner Low", "Replace Toner", "Cartidge Error",
        "Jam Inside", "Jam Rear")

    # brother 官網上所列的異常
    # https://support.brother.com/g/s/id/htmldoc/printer/cv_hll5000d/use/manual/index.html#GUID-D508418E-CC5B-42EE-8001-EFFA0AFD6A51_168                

    # 為了避免漏抓, 不在以下正常的狀態就算異常.
    $normal_status = @("Sleep", "Deep Sleep", "Ready", "No Paper T1", "No Paper T2", "Printing", "Please Wait", "No Paper", "No Paper MP")


    # 定義要登入的網址
    # 一開始的登入畫面, 可不用輸入密碼.
    $url_login = "/general/status.html"
    # 詳細的資訊, 需要登入 
    $url_info = "/general/information.html?kind=item"

    foreach ($printer in $printers.keys) {
            
        $network_status = Test-Connection -IPAddress $printers.$printer.ip -Count 3 -Quiet
        Write-debug "printer ip: $($printers.$printer.ip) $($printers.$printer.location) network status: $network_status"

        if ($network_status -eq $true) {

            $response = Invoke-WebRequest -Uri "http://$($printers.$printer.ip)" -SessionVariable sess
            $scrftoken = $response.InputFields | Where-Object -FilterScript { $_.id -eq "CSRFToken" }
        
            # 使用正規表示式取得設備狀態
            # 使用正規表示式搜尋 <div id="moni_data"><span class="moni moniOk"> 及其後的內容, 直到遇到 </span> 為止。(?s) 允許 . 匹配換行符。
            # 從正規表示式的捕獲群組取得匹配的文字內容, 並去除前後空白字元。

            $deviceStatus = [regex]::Match($response.Content, '(?s)<div id="moni_data"><span class="moni moni(.*?)">(.*?)</span>').Groups[2].Value.Trim()
        
            Write-Debug "Device status: $deviceStatus"
        
            <# #####################此段為登入印表機的web介面, 取得更多內容, 但目前暫無用到.  ##############################

            #登入取得更多資訊
        
            $formData = @{
                "CSRFToken" = $scrftoken.value
                "B55d"      = $printers.$printer.password_vhwc
                "loginurl"  = $url_login
            }

            $response_login = Invoke-WebRequest -Uri "http://$($printers.$printer.ip)$($url_login)" -Body $formData -Method Post -WebSession $sess

            if ($response_login.Content.Contains("Password&#32;Error")) {
                Write-Debug "password_vhwc: $($formData.B55d) fail, try password_factory: $($printers.$printer.password_factroy)"
                $formData.B55d = $printers.$printer.password_factroy
            
                $response_login = Invoke-WebRequest -Uri "http://$($printers.$printer.ip)$($url_login)" -Body $formData -Method Post -WebSession $sess
            }

            $response_info = Invoke-WebRequest -Uri "http://$($printers.$printer.ip)$($url_info)" -WebSession $sess 

            

            # 將取得的網頁,存到檔案.
            # Out-File -InputObject $response_info.Content -FilePath "d:\$($printer).html"



            ########################################################################################################### #>

            if ($deviceStatus -notin $normal_status) {
                $msg = "🚨L5100DN `nName: $printer `n"
                $msg += "IP: $($printers.$printer.ip) `n"
                $msg += "Status: $deviceStatus `n"
                $msg += "Location: $($printers.$printer.location)"

                Send-LineNotifyMessage -Token $line_apikey -Message $msg
            }

        }
        else {

            if ($printers.$printer.always_on -eq $true) {
                $msg = "🚨L5100DN `nName: $printer `n"
                $msg += "IP: $($printers.$printer.ip) `n"
                $msg += "Status: Network Fail, 注意此機須在線! `n"
                $msg += "Location: $($printers.$printer.location)"

                Send-LineNotifyMessage -Token $line_apikey -Message $msg

            }
        }
    }
}




function schedulecheck-tc200 {
    param(
        $printers
    )
    
    # status page
    $url_status = "/cgi-bin/status.cgi"

    $normal_status = @('Ready', 'Printing')


    foreach ($printer in $printers.keys) {

        $network_status = Test-Connection -IPAddress $($printers.$printer.ip) -Count 3 -Quiet
        Write-debug "printer ip: $($printers.$printer.ip) $($printers.$printer.location) network status: $network_status"

        if ($network_status -eq $true) {
            
            # TSC TC200 的web介面, 用了http 0.9的 拹定, invoke-webrequest 用出現錯誤無法使用， 
            # 改用curl.exe 的方式取得網頁資料.

            # $response = & curl.exe --http0.9 http://172.20.5.177/cgi-bin/status.cgi
            # $response = & "curl.exe" "--http0.9" "http://$($printers.$printer.ip)$url_status"

            $response = Invoke-Command -ScriptBlock { & "curl.exe" "--http0.9" "http://$($printers.$printer.ip)$url_status" }

            $devicestatus = [regex]::Match($response, 'Printer Status</TD><TD></TD></TR><TR><TD class=(.*?)>(.*?)</TD>').Groups[2].Value.Trim()
            write-debug "device satus: $devicestatus"

            if ($deviceStatus -notin $normal_status) {
                $msg = "🚨TSC Barcode `nName: $printer `n"
                $msg += "IP: $($printers.$printer.ip) `n"
                $msg += "Status: $deviceStatus `n"
                $msg += "Location: $($printers.$printer.location)"

                Send-LineNotifyMessage -Token $line_apikey -Message $msg
            }

        }
        else {
            #network fail
            if ($printers.$printer.always_on -eq $true) {
                $msg = "🚨TSC Barcode `nName: $printer `n"
                $msg += "IP: $($printers.$printer.ip) `n"
                $msg += "Status: Network Fail, 注意此機須在線! `n"
                $msg += "Location: $($printers.$printer.location)"

                Send-LineNotifyMessage -Token $line_apikey -Message $msg

            }
        }
    }
}


function schedulecheck-idmsclients {

    param (
        $idms_clients
    )

    # 迴圈檢查每個客戶端
    foreach ($clientName in $idms_clients.Keys) {
        $clientInfo = $idsm_clients[$clientName]
        $ipAddress = $clientInfo.ip
        $portNumber = $clientInfo.port
        $location = $clientInfo.location
        $always_on = $clientInfo.always_on

        # 檢查IP是否可Ping通
        if (Test-Connection -ComputerName $ipAddress -Count 2 -Quiet) {
            Write-Debug "Ping $clientName ($ipAddress) 正常."
            
            # 檢查Port服務是否有回應
            $connectionTestResult = Test-NetConnection -ComputerName $ipAddress -Port $portNumber 
            if ($connectionTestResult.TcpTestSucceeded) {
                Write-Debug "  - Port $portNumber 回應正常."
            }
            else {
                Write-Debug "  - Port $portNumber 無回應.發送 Line 通知"
                
                # 發送 Line 通知
                $msg = "🚨分散式client `nName: $clientName `n"
                $msg += "IP Status: $($ipaddress) ping正常 `n"
                if ($always_on -eq $true) {
                    $msg += "Port Status: $portNumber 無回應, 可以未登入或程式末執行. !注意此機須在線! `n"
                }
                else {
                    $msg += "Port Status: $portNumber 無回應 `n" 
                }
                $msg += "Location: $location"
                
                Send-LineNotifyMessage -Token $line_apikey -Message $msg
            }
        }
        else {
            Write-Debug "$clientName ($ipAddress) Ping不到,發送 Line 通知 "
       
            # 發送 Line 通知
            $msg = "🚨分散式client `nName: $clientName `n"
            if ($always_on -eq $true) {
                $msg += "IP Status: $($ipaddress) Ping不到, 注意此機須在線! `n"
            }
            else {
                $msg += "IP Status: $($ipaddress) Ping不到 `n"
            }
            $msg += "Location: $location"

            Send-LineNotifyMessage -Token $line_apikey -Message $msg
        }
    }
}



# 把always_on的過慮出來. always_on = $true 表示這台必須隨時在線.
$L5100DNsWithAlwaysOn = @{}
foreach ($printer in $L5100DNs.keys) {
    if ($L5100DNs.$printer.always_on -eq $true) {
        $L5100DNsWithAlwaysOn.$printer = $L5100DNs.$printer
    }
}

$tc200sWithAlwaysOn = @{}
foreach ($printer in $tc200s.keys) {
    if ($tc200s.$printer.always_on -eq $true) {
        $tc200sWithAlwaysOn.$printer = $tc200s.$printer
    }
}


$idsm_clientswithAlwaysOn = @{}
foreach ($client in $idsm_clients.keys) {
    if ($idsm_clients.$client.always_on -eq $true) {
        $idsm_clientswithAlwaysOn.$client = $idsm_clients.$client
    }
}

while ($true) {
    $now = Get-Date
    if ($now.Hour -in (8, 14) -and $now.Minute -in (0)) {
        Write-debug "$now : Daily check L5100DN all"
        Schedulecheck-L5100DN -printers $L5100DNs        
    }
    elseif ($now.Hour -in $timer_hours -and $now.minute -in $timer_minutes) {
        write-debug "$now : Check L5100DN always on"
        Schedulecheck-L5100DN -printers $L5100DNsWithAlwaysOn
    }

    if ($now.Hour -in (8, 14) -and $now.Minute -in (0)) {
        Write-debug "$now : Daily check TSC barcode all"
        schedulecheck-tc200 -printers $tc200s       
    }
    elseif ($now.Hour -in $timer_hours -and $now.minute -in $timer_minutes) {
        write-debug "$now : Check TSC Barcode always on"
        schedulecheck-tc200 -printers $tc200sWithAlwaysOn
    }

    if ($now.Hour -in (8, 14) -and $now.Minute -in (0)) {
        Write-debug "$now : Daily check IDMS clients all"
        schedulecheck-idmsclients -idms_clients $idsm_clients
    }
    elseif ($now.Hour -in $timer_hours -and $now.minute -in $timer_minutes) {
        write-debug "$now : Check IDMS clients always on"
        schedulecheck-idmsclients -idms_clients $idsm_clientswithAlwaysOn
    }

    start-sleep -Seconds 60
}

