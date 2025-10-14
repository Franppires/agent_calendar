import datetime
import os.path
import pickle
import re
from dateutil import parser
import pytz 
import webbrowser

# --- Imports para a Interface Gráfica ---
import customtkinter as ctk 
from tkinter import messagebox, simpledialog 

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# --- Configurações Globais ---
SCOPES = ['https://www.googleapis.com/auth/calendar']
TIME_ZONE_NAME = 'America/Sao_Paulo'
TIMEZONE = pytz.timezone(TIME_ZONE_NAME)
SERVICE = None 

# Configurações do CustomTkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue") 

# --- FUNÇÕES NOVAS E MODIFICADAS ---

def get_calendar_service():
    """Autentica o usuário e constrói o objeto de serviço da API."""
    global SERVICE
    if SERVICE:
        return SERVICE
    # ... (Resto da autenticação) ...
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    SERVICE = build('calendar', 'v3', credentials=creds)
    return SERVICE

def get_calendar_list(service):
    """NOVO: Busca todos os calendários disponíveis e os armazena em um dicionário."""
    calendar_list_result = service.calendarList().list().execute()
    items = calendar_list_result.get('items', [])
    
    # Mapeia Nome do Calendário -> ID do Calendário
    calendars = {}
    
    # Usamos get('summary') para o nome visível, e get('id') para o ID real
    for item in items:
        # Se for um calendário principal, usamos o email como nome
        name = item.get('summaryOverride') or item.get('summary')
        
        # O ID é usado para agendar, verificar conflito e buscar horários livres.
        calendars[name] = item.get('id') 
        
    return calendars

def check_event_conflict(service, start_time_iso, end_time_iso, calendar_id):
    """MODIFICADO: Recebe o calendar_id para verificar o calendário correto."""
    events_result = service.events().list(
        calendarId=calendar_id, 
        timeMin=start_time_iso,
        timeMax=end_time_iso,
        singleEvents=True,
        orderBy='startTime',
        maxResults=1
    ).execute()
    # ... (Resto da lógica de checagem mantida inalterada) ...
    events = events_result.get('items', [])
    
    if events:
        first_conflict = events[0]
        summary = first_conflict.get('summary', 'Sem Título')
        
        start_time_data = first_conflict.get('start', {})
        start_time_iso_conf = start_time_data.get('dateTime', start_time_data.get('date'))
        end_time_data = first_conflict.get('end', {})
        end_time_iso_conf = end_time_data.get('dateTime', end_time_data.get('date'))
        
        try:
            start_dt = parser.parse(start_time_iso_conf)
            end_dt = parser.parse(end_time_iso_conf)
            time_info = f"{start_dt.strftime('%d/%m %H:%M')} - {end_dt.strftime('%d/%m %H:%M')}"
        except Exception:
            time_info = "Data/Hora Indisponível"
            
        return True, f"O horário se sobrepõe ao evento: '{summary}' ({time_info})"
    else:
        return False, None

def find_first_free_slot(service, duration_minutes, calendar_id):
    """MODIFICADO: Recebe o calendar_id para buscar no calendário correto."""
    now = datetime.datetime.now(TIMEZONE)
    search_end = now + datetime.timedelta(days=14)
    
    if now.minute > 30:
        search_start = now.replace(minute=0, second=0, microsecond=0) + datetime.timedelta(hours=1)
    else:
        search_start = now.replace(minute=30, second=0, microsecond=0)

    try:
        events_result = service.events().list(
            calendarId=calendar_id, # Usando o ID da agenda
            timeMin=search_start.isoformat(),
            timeMax=search_end.isoformat(),
            singleEvents=True,
            orderBy='startTime',
        ).execute()
        
        # ... (Resto da lógica de busca mantida inalterada) ...
        events = events_result.get('items', [])

        current_time = search_start
        
        for event in events:
            event_start_iso = event['start'].get('dateTime', event['start'].get('date'))
            event_end_iso = event['end'].get('dateTime', event['end'].get('date'))
            
            event_start = parser.parse(event_start_iso).astimezone(TIMEZONE)
            event_end = parser.parse(event_end_iso).astimezone(TIMEZONE)
            
            if event_start > current_time:
                gap = event_start - current_time
                if gap.total_seconds() >= duration_minutes * 60:
                    return current_time 
            
            current_time = event_end

        return current_time

    except Exception as e:
        messagebox.showerror("Erro de Busca", f"Ocorreu um erro ao buscar horários: {e}")
        return None

def open_event_link(link):
    """Abre o link do evento no navegador padrão do sistema."""
    webbrowser.open(link)

def create_calendar_event(service, summary, location, description, start_time, end_time, calendar_id, attendees=None, time_zone='America/Sao_Paulo'):
    """MODIFICADO: Recebe o calendar_id para criar o evento na agenda correta."""
    event = {
        'summary': summary,
        'location': location,
        'description': description,
        'start': {'dateTime': start_time, 'timeZone': time_zone},
        'end': {'dateTime': end_time, 'timeZone': time_zone},
        'reminders': {'useDefault': False, 'overrides': [{'method': 'email', 'minutes': 24 * 60}, {'method': 'popup', 'minutes': 10}]},
    }
    
    if attendees:
        event['attendees'] = [{'email': email.strip()} for email in attendees if email.strip()]

    try:
        event = service.events().insert(calendarId=calendar_id, body=event).execute() # Usando o ID da agenda
        event_link = event.get('htmlLink')
        
        # --- Janela de Sucesso Personalizada ---
        success_window = ctk.CTkToplevel()
        success_window.title("Sucesso!")
        success_window.geometry("350x180")
        success_window.transient(success_window.master) 
        
        ctk.CTkLabel(success_window, text="Evento criado com sucesso!", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=15)
        ctk.CTkLabel(success_window, text=f"Horário: {parser.parse(start_time).strftime('%d/%m/%Y %H:%M')}").pack()
        
        ctk.CTkButton(
            success_window,
            text="Abrir Evento no Google Calendar",
            command=lambda: open_event_link(event_link),
            fg_color="green" 
        ).pack(pady=15, padx=20, fill='x')

        return True
    except Exception as e:
        messagebox.showerror("Erro na API", f"Ocorreu um erro ao criar o evento: {e}")
        return False


# --- LÓGICA DE TRATAMENTO DA INTERFACE ---

def parse_time_input(user_input, duration_minutes):
    """Função auxiliar para processar a string de data/hora."""
    # ... (Lógica de parsing mantida inalterada) ...
    user_input = user_input.lower().strip()
    start_datetime = None

    if 'amanha' in user_input or 'amanhã' in user_input or 'hoje' in user_input:
        base_date = datetime.datetime.now()
        if 'amanha' in user_input or 'amanhã' in user_input:
            base_date += datetime.timedelta(days=1)
        
        time_match = re.search(r'(\d{1,2}):(\d{2})', user_input) or re.search(r'(\d{1,2})(\d{2})(?!\d)', user_input) or re.search(r'(\d{1,2})(?!\d)', user_input)

        hour, minute = 9, 0
        if time_match:
            hour = int(time_match.group(1))
            minute = int(time_match.group(2)) if len(time_match.groups()) > 1 else 0

        start_datetime = base_date.replace(hour=hour, minute=minute, second=0, microsecond=0)

    else:
        replacements = {'terca': 'tuesday', 'terça': 'tuesday', 'quarta': 'wednesday', 'quinta': 'thursday', 'sexta': 'friday', 'sabado': 'saturday', 'sábado': 'saturday', 'domingo': 'sunday', 'segunda': 'monday'}
        for pt, en in replacements.items():
            user_input = user_input.replace(pt, en)
        
        start_datetime = parser.parse(user_input, dayfirst=True)
    
    if start_datetime is None:
        raise ValueError("Não foi possível interpretar a data/hora.")

    start_datetime = TIMEZONE.localize(start_datetime, is_dst=None)
    end_datetime = start_datetime + datetime.timedelta(minutes=duration_minutes)
    
    return start_datetime, end_datetime


def handle_agendamento(summary_var, location_var, time_var, duration_var, calendar_var, attendees_var, calendar_map):
    """MODIFICADO: Recebe a variável de seleção de calendário e o mapa de calendários."""
    summary = summary_var.get()
    location = location_var.get()
    time_input = time_var.get()
    duration_input = duration_var.get()
    calendar_name = calendar_var.get() # NOVO: Pega o nome da agenda selecionada
    attendees_str = attendees_var.get()

    # Obtém o ID real do calendário a partir do nome
    calendar_id = calendar_map.get(calendar_name) 

    if not summary or not time_input or not duration_input or not calendar_id:
        messagebox.showerror("Erro", "Título, Data/Hora, Duração e Agenda são campos obrigatórios.")
        return

    # 1. Validação da Duração
    try:
        duration_minutes = int(duration_input)
        if duration_minutes <= 0:
            raise ValueError("A duração deve ser um número positivo.")
    except ValueError:
        messagebox.showerror("Erro de Formato", "Duração deve ser um número inteiro em minutos.")
        return
    
    attendees = [email.strip() for email in attendees_str.split(',') if email.strip()]

    try:
        # 2. Parsing da hora
        start_datetime, end_datetime = parse_time_input(time_input, duration_minutes)
        start_time_str = start_datetime.isoformat()
        end_time_str = end_datetime.isoformat()
        
        # 3. Verificação de conflito (passando o calendar_id)
        is_conflict, conflict_message = check_event_conflict(SERVICE, start_time_str, end_time_str, calendar_id)
        
        if is_conflict:
            response = messagebox.askyesnocancel(
                "Conflito de Horário", 
                f"ATENÇÃO: Conflito de horário!\n{conflict_message}\n\nDeseja agendar mesmo assim? (Não = Buscar próximo horário; Cancelar = Abortar)"
            )
            
            if response is True: 
                create_calendar_event(SERVICE, summary, location, f"Agendado via Agente PC. Local: {location}", start_time_str, end_time_str, calendar_id, attendees, TIME_ZONE_NAME)
            
            elif response is False: 
                free_slot_start = find_first_free_slot(SERVICE, duration_minutes, calendar_id)
                
                if free_slot_start:
                    free_slot_end = free_slot_start + datetime.timedelta(minutes=duration_minutes)
                    
                    new_time_str = free_slot_start.strftime('%d/%m/%Y %H:%M')
                    confirm = messagebox.askyesno(
                        "Horário Livre Encontrado",
                        f"Próximo horário livre: {new_time_str} ({duration_minutes} min).\n\nDeseja agendar neste horário?"
                    )
                    
                    if confirm:
                        create_calendar_event(SERVICE, summary, location, f"Agendado via Agente PC. Local: {location}", free_slot_start.isoformat(), free_slot_end.isoformat(), calendar_id, attendees, TIME_ZONE_NAME)
                    else:
                        messagebox.showinfo("Ação Cancelada", "Você pode inserir um novo horário na janela principal.")
                else:
                    messagebox.showinfo("Busca Falhou", "Não foi possível encontrar um horário livre nas próximas 2 semanas.")

        else:
            # Sem conflito: Agenda diretamente (passando o calendar_id)
            create_calendar_event(SERVICE, summary, location, f"Agendado via Agente PC. Local: {location}", start_time_str, end_time_str, calendar_id, attendees, TIME_ZONE_NAME)
            
    except ValueError as e:
        messagebox.showerror("Erro de Formato", f"Erro no formato da Data/Hora: {e}.\nUse formatos como 'hoje 15:30' ou '2025-01-15 15:00'.")
    except Exception as e:
        messagebox.showerror("Erro Desconhecido", f"Ocorreu um erro inesperado: {e}")


def create_gui(calendar_map):
    """MODIFICADO: Recebe o mapa de calendários e adiciona o ComboBox."""
    
    root = ctk.CTk()
    root.title("Agente de Calendário Proativo")
    root.geometry("450x500") # Aumenta a janela para o novo campo
    
    main_frame = ctk.CTkFrame(root, corner_radius=10)
    main_frame.pack(padx=20, pady=20, fill="both", expand=True)

    # Variáveis de controle
    summary_var = ctk.StringVar(root)
    location_var = ctk.StringVar(root)
    time_var = ctk.StringVar(root)
    duration_var = ctk.StringVar(root, value="30")
    calendar_var = ctk.StringVar(root) # NOVO: Variável para a agenda selecionada
    attendees_var = ctk.StringVar(root)
    
    # Lista de nomes para o ComboBox
    calendar_names = list(calendar_map.keys())
    # Define o primeiro calendário como padrão, se houver
    if calendar_names:
        calendar_var.set(calendar_names[0]) 
    
    ctk.CTkLabel(main_frame, text="Novo Agendamento Rápido", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(10, 15))

    # 1. Título e Local (Inalterados)
    ctk.CTkLabel(main_frame, text="Título do Evento:", anchor="w").pack(fill='x', padx=10, pady=(0, 0))
    ctk.CTkEntry(main_frame, textvariable=summary_var, placeholder_text="Ex: Reunião com Cliente").pack(fill='x', padx=10, pady=(0, 5))

    ctk.CTkLabel(main_frame, text="Local (Sala / Online):", anchor="w").pack(fill='x', padx=10, pady=(0, 0))
    ctk.CTkEntry(main_frame, textvariable=location_var, placeholder_text="Ex: Sala A ou Google Meet").pack(fill='x', padx=10, pady=(0, 5))
    
    # 2. SELEÇÃO DE AGENDA (NOVO COMPONENTE)
    ctk.CTkLabel(main_frame, text="Selecione a Agenda:", anchor="w").pack(fill='x', padx=10, pady=(0, 0))
    ctk.CTkComboBox(main_frame, values=calendar_names, variable=calendar_var).pack(fill='x', padx=10, pady=(0, 5))
    
    # 3. Data/Hora e Duração (Inalterados)
    ctk.CTkLabel(main_frame, text="Data e Hora (Ex: hoje 15:30):", anchor="w").pack(fill='x', padx=10, pady=(0, 0))
    ctk.CTkEntry(main_frame, textvariable=time_var, placeholder_text="Ex: amanhã 10:00").pack(fill='x', padx=10, pady=(0, 5))
    
    ctk.CTkLabel(main_frame, text="Duração (em minutos):", anchor="w").pack(fill='x', padx=10, pady=(0, 0))
    ctk.CTkEntry(main_frame, textvariable=duration_var, placeholder_text="Ex: 60").pack(fill='x', padx=10, pady=(0, 5))
    
    # 4. Convidados (Inalterados)
    ctk.CTkLabel(main_frame, text="Convidados (Emails, separados por vírgula):", anchor="w").pack(fill='x', padx=10, pady=(0, 0))
    ctk.CTkEntry(main_frame, textvariable=attendees_var, placeholder_text="Ex: joao@email.com, maria@email.com").pack(fill='x', padx=10, pady=(0, 15))
    
    # 5. Botão Agendar (Passando o mapa de calendários)
    ctk.CTkButton(
        main_frame, 
        text="Agendar Evento Proativamente", 
        command=lambda: handle_agendamento(summary_var, location_var, time_var, duration_var, calendar_var, attendees_var, calendar_map),
        font=ctk.CTkFont(size=12, weight="bold")
    ).pack(fill='x', padx=10, pady=(10, 10))
    
    root.mainloop()


# --- Execução Principal (MUDANÇA) ---
if __name__ == '__main__':
    try:
        SERVICE = get_calendar_service()
        # NOVO: Busca a lista de calendários antes de iniciar a GUI
        CALENDAR_MAP = get_calendar_list(SERVICE)
        if not CALENDAR_MAP:
            messagebox.showerror("Erro de Calendário", "Nenhuma agenda foi encontrada na sua conta Google.")
        else:
            create_gui(CALENDAR_MAP)
    except Exception as e:
        messagebox.showerror("Erro de Inicialização", f"Falha ao iniciar o serviço de calendário. Verifique o arquivo credentials.json e o token.pickle. Erro: {e}")