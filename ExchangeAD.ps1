# AUTHOR --> Thespartoos
# TARGET --> Install Microsoft Exchange on a Windows Server with all the requeriments. 
# Exchange 2019 --> https://www.microsoft.com/en-us/download/details.aspx?id=102900
function banner {

	$banner = @()
	$banner += ''
    $banner += '___________             .__                                      _____  ________   '
	$banner += '\_   _____/__  ___ ____ |  |__ _____    ____    ____   ____     /  _  \ \______ \  '
	$banner += ' |    __)_\  \/  // ___\|  |  \\__  \  /    \  / ___\_/ __ \   /  /_\  \ |    |  \ '
	$banner += ' |        \>    <\  \___|   Y  \/ __ \|   |  \/ /_/  >  ___/  /    |    \|    `   \'
	$banner += '/_______  /__/\_ \\___  >___|  (____  /___|  /\___  / \___  > \____|__  /_______  /'
	$banner += '        \/      \/    \/     \/     \/     \//_____/      \/          \/        \/ '
	$banner += ''
	$banner += '                                                  [By Alejandro Ruiz (Thespartoos)]'
	$banner += ''
	$banner | foreach-object {
		Write-Host $_ -ForegroundColor (Get-Random -Input @('Green','Cyan','Yellow','gray','white'))
	}

}

function usage {

    Write-Output ''
    Write-Host '[!] Usage:' -ForegroundColor "Red"
    Write-Output ''
    Write-Host "1 --> PS C:\> Import-Module .\ExchangeAD.ps1" -ForegroundColor "yellow"
    Write-Host "2 --> PS C:\> Requeriments (DC-Name, Net Framework 4.8, Visual C++, API)" -ForegroundColor "yellow"
    Write-Host "3 --> PS C:\> AD_Requeriments (AD Services, Rol IIS)" -ForegroundColor "yellow"
    Write-Host "4 --> PS C:\> ExchangeAD (Rol DNS, Install Exchange)" -ForegroundColor "yellow"
    #Write-Host "5 --> PS C:\> ExchangeConf (Set up Exchange)" -ForegroundColor "green"
    Write-Output ''
}

function Requeriments {
    
    Clear-Host
    Start-Sleep -Seconds 2

    # DC_NAME
    Write-Output ''
    $dc_name = Read-Host "[*] Deseas cambiar el nombre de tu equipo [S/N]"

    if ($dc_name -eq "S") {
        Write-Output ''
        $namePC = Read-Host "[*] Elige tu nombre de equipo"
        Write-Output ''

        Write-Host "[*] Cambiando el nombre de equipo a $namePC"
        Write-Output ''

        Rename-Computer -NewName $namePC

        Write-Output ''
        Write-Host "[V] Nombre de equipo cambiado exitosamente" -ForegroundColor "green"
        Write-Output ''

        Write-Host "[!] Es necesario reiniciar el equipo para que los cambios tengan efecto" -ForegroundColor "red"
        Write-Output ''
        Write-Host "[!] Al reiciarse ejecute de nuevo esta function" -ForegroundColor "yellow"
        Write-Output ''

        Start-Sleep -Seconds 7
        Restart-Computer
        Start-Sleep -Seconds 4
    }

    elseif ($dc_name -eq "N") {
        Write-Output ''
    }

    elseif ([string]::IsNullOrEmpty($AD_Service)) {
        Write-Output ''
        Write-Host "[!] Debes especificar una letra" -ForegroundColor "red"
        Write-Output ''
    }

    Write-Output ''
    Write-Host "[!] Ahora toca descargarse Net Framework 4.8 Para el Servidor Exchange" -ForegroundColor "green"
    Write-Output ''
    Write-Host "[*] Url: https://go.microsoft.com/fwlink/?linkid=2088631" -ForegroundColor "yellow"
    Write-Output ''
    Start-Sleep -Seconds 1

    Write-Host "[*] Cuando lo descargue inicie el instalador e instalelo" -ForegroundColor "yellow"
    Write-Output ''

    $net = Read-Host "[!] Cuando tengas instalado el NetFrameWork escriba [S] y pulse ENTER para continuar"

    if ($net -eq "S") {
        Write-Output ''
        Write-Host "[V] Instalado" -ForegroundColor "green"
        Write-Output ''
        Start-Sleep -Seconds 2
        Clear-Host
    }
    # VISUAL C++ --> https://www.microsoft.com/en-us/download/confirmation.aspx?id=40784
    Write-Output ''
    Write-Host "[!] Ahora hay que descargarse Visual C++ e instalarlo" -ForegroundColor "green"
    Write-Output ''

    Write-Host "[*] Url: https://www.microsoft.com/en-us/download/confirmation.aspx?id=40784" -ForegroundColor "yellow"
    Write-Output ''
    $visual = Read-Host "[!] Cuando tengas instalado el Visual C++ escriba [S] y pulse ENTER para continuar"

    if ($visual -eq "S") {
        Write-Output ''
        Write-Host "[V] Instalado" -ForegroundColor "green"
        Start-Sleep -Seconds 1
        Clear-Host
    }

    # API --> https://www.microsoft.com/en-us/download/confirmation.aspx?id=34992
    Write-Output ''
    Write-Host "[!] Ahora hay que descargarse API e instalarlo" -ForegroundColor "yellow"
    Write-Output ''

    Write-Host "[*] Url: https://www.microsoft.com/en-us/download/confirmation.aspx?id=34992" -ForegroundColor "yellow"
    Write-Output ''
    $API = Read-Host "[!] Cuando tengas instalado API escriba [S] y pulse ENTER para continuar"

    if ($API -eq "S") {
        Write-Output ''
        Write-Host "[V] Instalado" -ForegroundColor "green"
        Write-Output ''
        Start-Sleep -Seconds 1
        Clear-Host

        Write-Output ''
        Write-Host "[!] Es necesario volver a reiniciar" -ForegroundColor  "red"
        Write-Output ''
        Write-Host "[*] Al reiniciarse ejecute la funtion AD_Requeriments" -ForegroundColor "yellow"
        Write-Output ''

        Start-Sleep -Seconds 7

        Restart-Computer
    }

}

function AD_Requeriments {

    # Comprobar que está instalado Servicios de dominio de AD
    Clear-Host
    $AD_Service = Get-WindowsFeature | Where-Object Name -like "AD-Domain-Services" | Where-Object InstallState -eq "Installed"
    Start-Sleep -Seconds 1

    if ([string]::IsNullOrEmpty($AD_Service)) {
        Write-Output ''
        Write-Host "[!] AD Services are not installed" -ForegroundColor "red"
        Write-Output ''
        Start-Sleep -Seconds 3
        Clear-Host

        Write-Output ''
        Write-Host "[*] Instalando los servicios de dominio y configurando el dominio" -ForegroundColor "yellow"
        Write-Output ''

        Add-WindowsFeature RSAT-ADDS
        Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools

        Write-Output ''
        $domainName = Read-Host "[*] Elige tu nombre de dominio"
        Write-Output ''

        Start-Sleep -Seconds 10
        
        Write-Output ''
        Write-Host "[*] A continuacion, deberas proporcionar la password del usuario Administrador del dominio" -ForegroundColor "yellow"
        Write-Output ''

        Try { Install-ADDSForest -CreateDnsDelegation:$false -DatabasePath "C:\\Windows\\NTDS" -DomainMode "7" -DomainName $domainName -DomainNetbiosName "alejandrocorp" -ForestMode "7" -InstallDns:$true -LogPath "C:\\Windows\\NTDS" -NoRebootOnCompletion:$false -SysvolPath "C:\\Windows\\SYSVOL" -Force:$true } Catch { Restart-Computer }

        Write-Output ''
        Write-Host "[!] Se va a reiniciar el equipo. Deberas iniciar sesion como el usuario Administrador a nivel de dominio" -ForegroundColor "red"
        Write-Output ''
        Write-Host "[!] Al reiniciarse vuelva a ejecutar, Import-Module .\ExchangeAD.ps1 y luego AD_Requeriments" -ForegroundColor "yellow"

        Start-Sleep -Seconds 7
        Restart-Computer
    }
    else {
        Write-Output ''
        Write-Host "[+] AD Servies are installed" -ForegroundColor "green"
        Write-Output ''

        Start-Sleep -Seconds 2
        Clear-Host
    }

    # ROL IIS
    Write-Output ''
    Write-Host '[+] Instalando rol IIS con sus complementos obligatorios' -ForegroundColor "yellow"
    Write-Output ''
    Write-Host '[!] En Windows Server 2019, si se queda colgado pulse al ENTER' -ForegroundColor "yellow"
    Write-Output ''

    Install-WindowsFeature -name Web-Server -IncludeManagementTools
    Install-WindowsFeature -name Web-Http-Redirect
    Install-WindowsFeature -name Web-Request-Monitor
    Install-WindowsFeature -name Web-Client-Auth
    Install-WindowsFeature -name Web-Cert-Auth
    Install-WindowsFeature -name Web-WMI
    Install-WindowsFeature -name Web-Windows-Auth
    Install-WindowsFeature -name Web-Dyn-Compression
    Install-WindowsFeature -name Web-Basic-Auth
    Install-WindowsFeature -name Web-Digest-Auth
    Install-WindowsFeature -name Web-ISAPI-Filter
    Install-WindowsFeature -name Web-Mgmt-Service
    Install-WindowsFeature -name NET-WCF-HTTP-Activation45
    Install-WindowsFeature -name RSAT-Clustering
    Install-WindowsFeature -name RSAT-Clustering-Powershell
    Install-WindowsFeature -name RSAT-Clustering-Mgmt
    Install-WindowsFeature -name RSAT-Clustering-CmdInterface
    Install-WindowsFeature -name Web-Net-Ext45
    Install-WindowsFeature -name Web-ISAPI-Ext
    Install-WindowsFeature -name Web-ASP-NET45
    Install-WindowsFeature -name RPC-over-HTTP-proxy

    Write-Output ''
    Write-Host '[V] Instalado correctamente' -ForegroundColor "green"
    Write-Output ''

    Write-Output ''
    Write-Host "[*] Instalamos dependencias requeridas para Exchange 2016" -ForegroundColor "yellow"
    Write-Output ''

    Install-WindowsFeature Server-Media-Foundation
    Install-WindowsFeature RSAT-ADDS
    Write-Output ''
    Write-Host "[V] Instalado correctamente" -ForegroundColor "green"
    Start-Sleep -Seconds 3
    Clear-Host

    # Rewrite IIS --> https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_es-ES.msi

    Write-Output ''
    Write-Host "[*] Instalando Rewrite IIS..."
    Write-Output ''
    
    Invoke-WebRequest -uri "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_es-ES.msi" -OutFile rewrite_amd64.msi

    .\rewrite_amd64.msi
    Start-Sleep -Seconds 10
    Write-Output ''
    Write-Host '[!] En Windows Server 2019, si se queda colgado pulse al ENTER' -ForegroundColor "yellow"
    Write-Output ''
    Write-Host "[+] Instalado correctamente" -ForegroundColor "green"
    Write-Output ''

    Write-Host "[!] Se va a reiniciar el server para poder aplicar los cambios y proceder" -ForegroundColor "red"
    Write-Output ''
    Write-Host "[*] Al reiniciarse debes de volver a importar los modulos del script y ejecutar ExchangeAD para instarlo sin problema" -ForegroundColor "yellow"
    Write-Output ''
    Start-Sleep -Seconds 15

    Restart-Computer
}

function ExchangeAD {

    Clear-Host
    Start-Sleep -Seconds 2
    
    # Instalar ROL DNS

    Write-Host '[+] Instalando rol DNS en Windows Server' -ForegroundColor "yellow"
    $roleDNS = Get-WindowsFeature | Where-Object Name -like "DNS" | Where-Object InstallState -eq "Installed"

    if ([string]::IsNullOrEmpty($roleDNS)) {
        Write-Output ''
        Write-Host '[!] Is not installed' -ForegroundColor "red"
        Write-Output ''
        Start-Sleep -Seconds 2
        Write-Output ''
        Write-Host '[*] Instalando el servicio DNS'
        Install-WindowsFeature DNS
        Write-Output ''
        Write-Host "[V] Rol DNS instalado correctamente" -ForegroundColor "green"
        Write-Output ''
    }
    else {
        Write-Output ''
        Write-Host '[+] Rol DNS is installed' -ForegroundColor "green"
        Write-Output ''
        Start-Sleep -Seconds 2
    }
    Clear-Host
    Write-Output ''
    Write-Host "[*] Configurando previamente para su comienzo de instalacion" -ForegroundColor "green"
    Write-Output ''
    
    Write-Host '[!] Ahora toca descargarse la ISO de Exchange Server' -ForegroundColor "yellow"
    Write-Output ''
    Write-Host 'Exchange Windows Server 2019 --> [*] Url: https://www.microsoft.com/en-us/download/details.aspx?id=102900' -ForegroundColor "yellow"
    Write-Output ''
    Write-Host 'Exchange Windows Server 2016 --> [*] Url: https://www.microsoft.com/es-es/download/confirmation.aspx?id=57827' -ForegroundColor "yellow"
    Write-Output ''

    Start-Sleep -Seconds 2

    $exchange = Read-Host "[!] Cuando se descargue la iso de exchange escribe [S] y pulse ENTER para continuar"
    # Mount-DiskImage -ImagePath "C:\ubuntu.iso" --> montar
    # Get-Volume --> Ver discos
    # Dismount-DiskImage -ImagePath "C:\ubuntu.iso" --> desmontar
    if ($exchange -eq "S") {
        Write-Output ''
        Write-Host '[!] Ahora debes de montar la iso' -ForegroundColor "yellow"
        Write-Output ''

        Start-Sleep -Seconds 1
        $mount = Read-Host "[!] Cuando tengas montada la iso de Exchange escribe la letra de la unidad [ex --> E:]"
    }

    Set-Location $mount
    Clear-Host
    Write-Output ''
    Write-Host "[+] Iniciando instalacion Exchange preparativos" -ForegroundColor "yellow"
    Write-Output ''
    cmd /c setup/prepareschema /IAcceptExchangeServerLicenseTerms
    Clear-Host
    Write-Output ''
    cmd /c setup/prepareschema /IAcceptExchangeServerLicenseTerms_DiagnosticDataON
    Clear-Host
    Write-Output ''
    cmd /c setup/prepareAD /IAcceptExchangeServerLicenseTerms /OrganizationName:"Pruebas Exchange"
    Clear-Host
    Write-Output ''
    cmd /c setup/prepareAD /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /OrganizationName:"Pruebas Exchange"
    Clear-Host
    Write-Output ''
    cmd /c Setup
    Start-Sleep -Seconds 2
    Write-Output ''
    Write-Host "[V] Exchange instalado correctamente !" -ForegroundColor "green"
    Write-Output ''
    Start-Sleep -Seconds 2
     
}   

if ($args.Count -eq "0") {
    Clear-Host
    banner
    usage
}
else {
    Write-Output ''
    Write-Host "[!] Ejecuta sin argumentos" -ForegroundColor "red"
    Write-Output ''
}