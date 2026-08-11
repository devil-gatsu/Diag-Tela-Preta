import subprocess
import json
import os
import platform
from datetime import datetime

# Dicionário com as legendas e ações recomendadas para cada erro
DICIONARIO_ERROS = {
    "1000": {"descricao": "Falha no processo explorer.exe (Crash).", "acao": "Rodar sfc /scannow para reparar arquivos do sistema."},
    "1002": {"descricao": "O processo explorer.exe travou (Hang) e parou de responder.", "acao": "Verificar se há falhas de leitura no disco ou HD em 100%."},
    "1500": {"descricao": "Falha geral de logon.", "acao": "Verificar integridade do perfil de usuário no registro."},
    "1502": {"descricao": "O perfil de usuário falhou ao carregar.", "acao": "O Windows pode ter criado um perfil temporário. Corrigir o NTUSER.DAT."},
    "1508": {"descricao": "Falha no carregamento do registro do usuário.", "acao": "Reiniciar a máquina e rodar chkdsk /f /r."},
    "4101": {"descricao": "O driver de vídeo parou de responder e se recuperou.", "acao": "Atualizar ou reinstalar o driver da placa de vídeo usando DDU."},
    "4005": {"descricao": "Tempo limite de logon excedido.", "acao": "O sistema está muito lento para carregar o perfil. Verificar saúde do SSD/HD."},
    "41":   {"descricao": "Sistema foi reiniciado inesperadamente (Desligamento forçado).", "acao": "Nenhuma ação corretiva direta, mas confirma que o PC travou."}
}

def buscar_logs_tela_preta(horas=24):
    ids_alvo = ", ".join(DICIONARIO_ERROS.keys())
    
    ps_command = (
        f"Get-WinEvent -FilterHashTable @{{LogName='System', 'Application'; "
        f"ID={ids_alvo}; StartTime=(Get-Date).AddHours(-{horas})}} "
        f"-ErrorAction SilentlyContinue | "
        f"Select-Object Id, ProviderName, "
        f"@{{Name='Data';Expression={{$_.TimeCreated.ToString('dd/MM/yyyy HH:mm:ss')}}}}, "
        f"Message | ConvertTo-Json -Compress"
    )
    
    print("[*] Buscando erros de tela preta no Windows...")
    
    try:
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True, check=True)
        saida_json = result.stdout.strip()
        
        if not saida_json:
            print("[✓] Nenhum erro de tela preta encontrado.")
            gerar_relatorio([])
            return

        eventos = json.loads(saida_json)
        if isinstance(eventos, dict):
            eventos = [eventos]
            
        gerar_relatorio(eventos)
        print("[!] Relatório gerado com sucesso na pasta atual!")
            
    except Exception as e:
        print(f"[x] Ocorreu um erro ao buscar os logs: {e}")

def gerar_relatorio(eventos):
    nome_pc = platform.node()
    data_atual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_arquivo = f"Relatorio_TelaPreta_{nome_pc}_{data_atual}.txt"
    
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write(f"  RELATORIO DE DIAGNOSTICO - TELA PRETA (WINDOWS)\n")
        f.write("=" * 60 + "\n")
        f.write(f"Nome da Maquina: {nome_pc}\n")
        f.write(f"Data da Varredura: {datetime.now().strftime('%d/%M/%Y %H:%M:%S')}\n")
        f.write("-" * 60 + "\n\n")
        
        if not eventos:
            f.write("Status: NENHUM ERRO CRITICO ENCONTRADO nas ultimas 24 horas.\n")
            f.write("O sistema parece estar estavel.\n")
            return
            
        f.write(f"Foram encontrados {len(eventos)} eventos suspeitos:\n\n")
        
        for ev in eventos:
            ev_id = str(ev.get('Id'))
            info_erro = DICIONARIO_ERROS.get(ev_id, {"descricao": "Erro desconhecido", "acao": "Investigar."})
            
            f.write(f"ID do Erro : {ev_id}\n")
            f.write(f"Data/Hora  : {ev.get('Data')}\n")
            f.write(f"Fonte      : {ev.get('ProviderName')}\n")
            f.write(f"Descricao  : {info_erro['descricao']}\n")
            f.write(f"Resolucao  : {info_erro['acao']}\n")
            
            # Pega só as primeiras palavras da mensagem técnica do Windows para não poluir
            msg_bruta = " ".join(str(ev.get('Message')).split())
            f.write(f"Log do Win : {msg_bruta[:150]}...\n")
            f.write("-" * 60 + "\n")

if __name__ == "__main__":
    buscar_logs_tela_preta(24)
