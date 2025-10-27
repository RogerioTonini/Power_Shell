# ===================================================================
# Script: Inventario.ps1
# Objetivo: Coletar informações detalhadas do equipamento e S.O.
# ===================================================================

try {
	# --- Marca, Modelo, Número de Série ---
	$computador = Get-CimInstance Win32_ComputerSystem
	$marcaComputador = $computador.Manufacturer.Trim()
	$modeloComputador = $computador.Model.Trim()

	$bios       = Get-CimInstance Win32_BIOS
	$numSerieBIOS = $bios.SerialNumber.Trim()

	# --- Usuário Logado ---
	$nomeMaquina = ([string]$computador.Name).Trim()
	$usuarioLogado = if ($computador.UserName) {
		($computador.UserName -split '\\')[-1]
	} else {
		"N/D"
	}
	# $usuarioLogado = $computador.UserName

	# --- Sistema Operacional ---
	$so = Get-CimInstance Win32_OperatingSystem
	$versao_SO = "$($so.Caption) $($so.OSArchitecture) (Build $($so.BuildNumber))"

	# --- Placa Mãe ---
	$placaMae = Get-CimInstance Win32_BaseBoard
	$marcaPlacaMae = $placaMae.Manufacturer.Trim()
	$modeloPlacaMae = $placaMae.Product.Trim()
	$numSeriePlacaMae = "'" + ([string]$placaMae.SerialNumber).Trim()
	# $placaMae.SerialNumber.Trim()

	# --- Processador ---
	$cpu = Get-CimInstance Win32_Processor
	$infoProcessador = "$($cpu.Name.Trim())"

	# --- Memória RAM ---
	$memFisica = Get-CimInstance Win32_PhysicalMemory
	$memTotInstGB = [math]::Round(($memFisica | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)
	$memCapTotGB = [math]::Round($computador.TotalPhysicalMemory / 1GB, 2)

	# --- Armazenamento / Tipos de Dispositivos Físico (HDD / SSD / NVMe / M.2) ---
	$discos = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
	$armazenamento = ($discos | ForEach-Object {
		"$($_.DeviceID) - $([math]::Round($_.Size / 1GB, 2)) GB"
	}) -join " / "

	# try {
		# $discosFisicos = Get-PhysicalDisk | Select FriendlyName, MediaType, BusType
		# $tiposArmazenamento = ($discosFisicos | ForEach-Object {
			# $tipo = switch ($_.BusType) {
				# "NVMe" { "M.2 NVMe" }
				# "SATA" { if ($_.MediaType -eq "SSD") { "SSD SATA" } else { "HDD SATA" } }
				# "RAID" { "RAID" }
				# "USB"  { "USB / Externo" }
				# default { $_.MediaType }
			# }
			# "$($_.FriendlyName): $tipo"
		# }) -join "; "
	# } catch {
		# $tiposArmazenamento = "N/D"
	# }

	try {
		$discosFisicos = Get-PhysicalDisk |
			Where-Object {
				$_.FriendlyName -notmatch 'Virtual|Msft|Storage|RAM Disk|VHD|VDISK'
			} |
			Select-Object FriendlyName, MediaType, BusType, Model

		$tiposArmazenamento = ($discosFisicos | ForEach-Object {
			# Corrigir MediaType "Unspecified" baseado no nome do modelo
			$media = $_.MediaType
			if ($media -eq "Unspecified") {
				if ($_.Model -match "SSD|NVMe") {
					$media = "SSD"
				} elseif ($_.Model -match "HDD|ST|WD|Hitachi|Seagate|Toshiba") {
					$media = "HDD"
				}
			}

			# Classificar conforme o barramento
			$tipo = switch ($_.BusType) {
				"NVMe" { "M.2 NVMe" }
				"SATA" { if ($media -eq "SSD") { "SSD SATA" } else { "HDD SATA" } }
				"RAID" { "RAID" }
				"USB"  { "USB / Externo" }
				default { $media }
			}

			# Monta a descrição final
			"$($_.FriendlyName): $tipo"
		}) -join " / "
	}
	catch {
		$tiposArmazenamento = "N/D"
	}

	# --- Informações: Placa de Rede e Servidores: DHCP, DNS ---
	# DHCP e DNS
	$servidorDHCP = if ($adaptador.DHCPServer) { $adaptador.DHCPServer } else { "N/D" }

	if ($adaptador.DNSServerSearchOrder) {
		$dnsPrimario = $adaptador.DNSServerSearchOrder[0]
		$dnsSecundario = if ($adaptador.DNSServerSearchOrder.Count -gt 1) { $adaptador.DNSServerSearchOrder[1] } else { "N/D" }
	} else {
		$dnsPrimario = "N/D"
		$dnsSecundario = "N/D"
	}

	$adaptador = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object {
        $_.IPEnabled -eq $true -and
        $_.Description -notmatch 'Virtual|Hyper-V|Bluetooth|Loopback|VMware|TAP|Wi-Fi Direct'
    } | Select-Object -First 1

	if ($adaptador) {
		# Tipo de IP
		$tipoIP = if ($adaptador.DHCPEnabled) { "DHCP" } else { "Estatico" }

		# IP e MAC
		$enderecoIP = $adaptador.IPAddress[0]
		$enderecoMAC = $adaptador.MACAddress

		# DHCP
		$servidorDHCP = if ($adaptador.DHCPServer) { $adaptador.DHCPServer } else { "N/D" }

		# DNS
		if ($adaptador.DNSServerSearchOrder) {
			$dnsPrimario = $adaptador.DNSServerSearchOrder[0]
			$dnsSecundario = if ($adaptador.DNSServerSearchOrder.Count -gt 1) {
				$adaptador.DNSServerSearchOrder[1]
			} else {
				"N/D"
			}
		} else {
			$dnsPrimario = "N/D"
			$dnsSecundario = "N/D"
		}
	} else {
		$tipoIP = "N/D"
		$ip = "N/D"
		$mac = "N/D"
		$servidorDHCP = "N/D"
		$dnsPrimario = "N/D"
		$dnsSecundario = "N/D"
	}

	# --- Montar objeto final ---
	$info = [PSCustomObject]@{
		"Marca Equipamento"        = $marcaComputador
		"Modelo Equipamento"       = $modeloComputador
		"Num de Serie (BIOS)"      = $numSeirieBIOS
		"Fabricante Placa Mae"     = $marcaPlacaMae
		"Modelo Placa Mae"         = $modeloPlacaMae
		"Num de Serie Placa Mae"   = $numSeriePlacaMae
		"Nome da estacao"          = $nomeMaquina
		"Usuario Logado"           = $usuarioLogado
		"Versao S.O."              = $versao_SO
		"Processador"              = $infoProcessador
		"Cap. Mem. RAM (GB)"       = $memCapTotGB
		"Total RAM Instalada (GB)" = $memTotInstGB
		"Armazenamento"            = $armazenamento
		"Tipo de Armazenamento"    = $tiposArmazenamento
		"Tipo de IP"               = $tipoIP
		"Endereco IP"              = $enderecoIP
		"Endereco MAC"             = $enderecoMAC
		"Servidor DHCP"            = $servidorDHCP
		"DNS Primario"             = $dnsPrimario
		"DNS Secundario"           = $dnsSecundario
	}

	# --- Exibir resultado formatado ---
	$info | Format-List

	# Determina o diretório do script, mesmo se executado no ISE ou VSCode
	if ($PSScriptRoot) {
		$scriptPath = $PSScriptRoot
	} elseif ($MyInvocation.MyCommand.Path) {
		$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
	} else {
		$scriptPath = Get-Location
	}
	
	# Define o nome do arquivo de inventário
	$arquivoInventario = Join-Path $scriptPath "Inventario_Sistema.csv"

	# Gera data/hora da coleta
	$datahoraColeta = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")

	# Cabeçalho do arquivo
    $cabecalho = "Marca_Computador;Modelo_Computador;Num_serie_BIOS;Marca_Placa_Mae;Mod_Placa_Mae;Num_Serie_Placa_Mae;Nome_Estacao;Usuario_Logado;Versao_SO;Info_Processador;Cap_Total_RAM_GB;Tot_RAM_Instalada_GB;Tipo_Armazenamento;Cap_Armazenamento;Tipo_IP;Endereco_IP;Endereco_MAC;Serv_DHCP;DNS_Primario;DNS_Secundario;Data_Hora"

    # Linha com os dados coletados
	$dados = "$marcaComputador;$modeloComputador;$numSerieBIOS;$marcaPlacaMae;$modeloPlacaMae;$numSeriePlacaMae;$nomeMaquina;$usuarioLogado;$versao_SO;$infoProcessador;$memCapTotGB;$memTotInstGB;$tiposArmazenamento;$armazenamento;$tipoIP;$enderecoIP;$enderecoMAC;$servidorDHCP;$dnsPrimario;$dnsSecundario;$datahoraColeta"

    # Cria ou adiciona os dados no arquivo
    if (-not (Test-Path $arquivoInventario)) {
        $cabecalho | Out-File -FilePath $arquivoInventario -Encoding UTF8
    }
    $dados | Out-File -FilePath $arquivoInventario -Encoding UTF8 -Append

    Write-Host "Inventario coletado e gravado em: $arquivoInventario" -ForegroundColor Green

} catch {
    Write-Host "Erro ao coletar informações: $_" -ForegroundColor Red
}

# MachinePolicy  Undefined
# UserPolicy     AllSigned
# Process        Undefined
# CurrentUser    Undefined (AllSigned)
# LocalMachine   Restricted
