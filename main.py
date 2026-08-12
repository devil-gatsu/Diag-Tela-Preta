import subprocess
import json
import platform
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def buscar_erros_hoje():
    # StartTime=(Get-Date).Date pega os logs a partir de 00:00:00 do dia de hoje
    ps_command = (
        f"Get-WinEvent -FilterHashTable @{{LogName='System', 'Application'; "
        f"Level=1, 2; StartTime=(Get-Date).Date}} "
        f"-ErrorAction SilentlyContinue | "
        f"Select-Object Id, ProviderName, LevelDisplayName, "
        f"@{{Name='Data';Expression={{$_.TimeCreated.ToString('dd/MM/yyyy HH:mm:ss')}}}}, "
        f"Message | ConvertTo-Json -Compress"
    )
    
    print("[*] Coletando erros criticos e avisos do DIA DE HOJE...")
    
    try:
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True, check=True, encoding='cp850')
        saida_json = result.stdout.strip()
        
        if not saida_json:
            print("[✓] Nenhum erro registrado no dia de hoje!")
            gerar_planilha([])
            return

        eventos = json.loads(saida_json)
        if isinstance(eventos, dict):
            eventos = [eventos]
            
        gerar_planilha(eventos)
        print(f"[!] Sucesso! Planilha gerada com {len(eventos)} erro(s).")
            
    except subprocess.CalledProcessError as e:
        print(f"[x] Falha ao executar o comando no PowerShell.")
    except json.JSONDecodeError:
        print("[x] Erro ao processar os dados. O Windows retornou um formato inesperado.")
    except Exception as e:
        print(f"[x] Erro inesperado: {e}")

def gerar_planilha(eventos):
    nome_pc = platform.node()
    data_atual = datetime.now().strftime("%Y-%m-%d")
    nome_arquivo = f"Relatorio_Erros_{nome_pc}_{data_atual}.xlsx"
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Erros Windows"
    
    # Cabeçalhos
    headers = ["Data/Hora", "Nível", "ID", "Fonte", "Detalhes do Erro"]
    ws.append(headers)
    
    if not eventos:
        ws.append(["-", "-", "-", "-", "Nenhum erro registrado hoje. Sistema saudável."])
    else:
        # Inverte para mostrar do mais antigo para o mais novo
        for ev in reversed(eventos):
            ev_id = str(ev.get('Id'))
            fonte = str(ev.get('ProviderName'))
            nivel = str(ev.get('LevelDisplayName', 'Erro'))
            data_hora = str(ev.get('Data'))
            
            # Limpa quebras de linha excessivas
            msg_bruta = str(ev.get('Message', 'Sem descrição.'))
            msg_limpa = " ".join(msg_bruta.split())
            
            ws.append([data_hora, nivel, ev_id, fonte, msg_limpa])

    # --- INÍCIO DA FORMATAÇÃO DA PLANILHA ---
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'),
                         top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD'))
    zebra_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")

    # Formatando o cabeçalho
    for col in range(1, 6):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = thin_border

    # Formatando as linhas de dados (com efeito Zebra)
    for row in range(2, ws.max_row + 1):
        for col in range(1, 6):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            if col in [1, 2, 3]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
            
            if row % 2 == 0:
                cell.fill = zebra_fill

    # Ajustando a largura das colunas
    widths = [20, 12, 10, 25, 90]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Adicionando Filtros e Congelando o cabeçalho
    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    ws.freeze_panes = "A2"
    
    wb.save(nome_arquivo)

if __name__ == "__main__":
    buscar_erros_hoje()
