import subprocess
import json
import platform
from datetime import datetime

def buscar_todos_erros(horas=24):
    """
    Busca todos os eventos de nível Crítico (1) e Erro (2) 
    nos logs de Sistema e Aplicativo.
    """
    # Level=1 (Crítico), Level=2 (Erro)
    ps_command = (
        f"Get-WinEvent -FilterHashTable @{{LogName='System', 'Application'; "
        f"Level=1, 2; StartTime=(Get-Date).AddHours(-{horas})}} "
        f"-ErrorAction SilentlyContinue | "
        f"Select-Object Id, ProviderName, "
        f"@{{Name='Data';Expression={{$_.TimeCreated.ToString('dd/MM/yyyy HH:mm:ss')}}}}, "
        f"Message | ConvertTo-Json -Compress"
    )
    
    print(f"[*] Coletando TODOS os erros e eventos críticos das últimas {horas} horas...")
    print("[*] Isso pode levar alguns segundos dependendo da quantidade de erros...")
    
    try:
        # Executa o PowerShell silenciosamente
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True, check=True)
        saida_json = result.stdout.strip()
        
        if not saida_json:
            print("[✓] Nenhum erro ou evento crítico encontrado no período!")
            gerar_relatorio([], horas)
            return

        eventos = json.loads(saida_json)
        
        # Se retornar apenas 1 erro, o JSON vem como dicionário, precisamos converter para lista
        if isinstance(eventos, dict):
            eventos = [eventos]
            
        gerar_relatorio(eventos, horas)
        print(f"[!] Sucesso! Relatório gerado com {len(eventos)} erro(s) encontrado(s).")
            
    except subprocess.CalledProcessError as e:
        print(f"[x] Falha ao executar o comando no PowerShell: {e}")
    except json.JSONDecodeError:
        print("[x] Erro ao processar os dados recebidos do Windows.")
    except Exception as e:
        print(f"[x] Erro inesperado: {e}")

def gerar_relatorio(eventos, horas):
    nome_pc = platform.node()
    data_atual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_arquivo = f"Relatorio_Geral_Erros_{nome_pc}.txt"
    
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        f.write("=" * 80 + "\n")
        f.write(" " * 20 + "RELATORIO GERAL DE ERROS DO WINDOWS\n")
        f.write("=" * 80 + "\n")
        f.write(f"Maquina Analisada: {nome_pc}\n")
        f.write(f"Periodo de Busca : Ultimas {horas} horas\n")
        f.write(f"Data da Geracao  : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("-" * 80 + "\n\n")
        
        if not eventos:
            f.write("Status: O sistema esta incrivelmente saudavel!\n")
            f.write("NENHUM evento Critico ou de Erro foi registrado nos logs de Sistema/Aplicativo.\n")
            return
            
        f.write(f"Total de problemas encontrados: {len(eventos)}\n\n")
        
        # Inverte a lista para mostrar do mais antigo para o mais recente (opcional, mas ajuda a entender a cronologia)
        eventos_cronologicos = reversed(eventos)
        
        for ev in eventos_cronologicos:
            ev_id = str(ev.get('Id'))
            fonte = str(ev.get('ProviderName'))
            data_hora = str(ev.get('Data'))
            
            # Limpa as quebras de linha gigantes que o Windows coloca nas mensagens de erro
            msg_bruta = str(ev.get('Message', 'Nenhuma descrição detalhada disponível para este erro.'))
            msg_limpa = " ".join(msg_bruta.split())
            
            f.write(f"[{data_hora}] ID: {ev_id} | Fonte: {fonte}\n")
            f.write(f"Comentario : {msg_limpa}\n")
            f.write("-" * 80 + "\n")

if __name__ == "__main__":
    # Ajustado para buscar erros das últimas 24 horas. 
    # Você pode alterar esse número para 48, 72, etc.
    buscar_todos_erros(24)
