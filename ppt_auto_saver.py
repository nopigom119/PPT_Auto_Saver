import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import win32com.client
import pythoncom
import json
import os
import threading # For initial COM init, actual periodic tasks use root.after

# --- Configuration ---
CONFIG_FILE = "autosave_ppt_config.json"
DEFAULT_SAVE_INTERVAL_SECONDS = 60
ACTIVE_LIST_REFRESH_MS = 5000  # 5 seconds

# --- COM Interaction Functions ---
def get_open_presentations_info():
    """
    현재 열려있는 모든 파워포인트 프레젠테이션의 이름과 전체 경로를 가져옵니다.
    Returns:
        list: 각 프레젠테이션 정보를 담은 딕셔너리 리스트.
              예: [{'name': 'MyPres.pptx', 'fullName': 'C:\\...\\MyPres.pptx'}, ...]
              저장되지 않은 경우 fullName은 None이 됩니다.
    """
    # 메인 스레드에서 COM이 초기화되었다고 가정합니다.
    # threading.current_thread().name == "MainThread" 일 때만 CoInitialize/CoUninitialize
    # 여기서는 App 클래스 생성자에서 한 번만 호출합니다.
    presentations_details = []
    try:
        powerpoint_app = win32com.client.DispatchEx("PowerPoint.Application")
        if powerpoint_app.Presentations.Count == 0:
            return presentations_details

        for presentation_obj in powerpoint_app.Presentations:
            name = "Unknown Presentation"
            full_name = None
            try:
                name = presentation_obj.Name
                full_name = presentation_obj.FullName
                presentations_details.append({'name': name, 'fullName': full_name})
            except pythoncom.com_error:
                # 주로 아직 저장되지 않은 프레젠테이션 (FullName 없음)
                presentations_details.append({'name': f"{name} (저장 안됨)", 'fullName': None})
            except Exception:
                # 기타 예외 발생 시
                if name != "Unknown Presentation":
                     presentations_details.append({'name': f"{name} (정보 오류)", 'fullName': None})
                # else: skip
    except pythoncom.com_error:
        # PowerPoint가 실행 중이 아니거나 접근 불가
        pass # UI에 메시지 표시 가능
    except Exception:
        pass # UI에 메시지 표시 가능
    return presentations_details

def save_ppt_by_fullname(fullname_to_save):
    """
    주어진 전체 경로를 가진 파워포인트 프레젠테이션을 저장합니다.
    Args:
        fullname_to_save (str): 저장할 프레젠테이션의 전체 경로.
    Returns:
        bool: 저장 성공 여부.
    """
    if not fullname_to_save:
        return False
    
    # 메인 스레드에서 COM이 초기화되었다고 가정합니다.
    try:
        powerpoint_app = win32com.client.DispatchEx("PowerPoint.Application")
        if powerpoint_app.Presentations.Count == 0:
            return False

        for presentation_obj in powerpoint_app.Presentations:
            try:
                if presentation_obj.FullName == fullname_to_save:
                    presentation_obj.Save()
                    print(f"성공적으로 저장됨: {fullname_to_save}")
                    return True
            except pythoncom.com_error:
                continue 
            except Exception:
                continue
        print(f"저장할 프레젠테이션을 찾지 못했거나 저장할 수 없음: {fullname_to_save}")
        return False
    except pythoncom.com_error:
        print(f"PowerPoint에 연결하여 {fullname_to_save} 저장 실패")
        return False
    except Exception:
        print(f"{fullname_to_save} 저장 중 예외 발생")
        return False

# --- Application Class ---
class PPTAutoSaverApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("PPT 자동 저장 도우미")
        self.root.geometry("800x600")

        # COM 초기화 (애플리케이션의 메인 스레드에서 한 번)
        try:
            pythoncom.CoInitializeEx(0) # 0 for COINIT_APARTMENTTHREADED
            self.com_initialized = True
        except pythoncom.com_error:
            messagebox.showerror("오류", "COM 라이브러리 초기화에 실패했습니다. 프로그램이 정상적으로 동작하지 않을 수 있습니다.")
            self.com_initialized = False


        self.active_presentations_data = []  # 현재 열린 PPT 목록 (UI용 아님, 실제 데이터)
        self.auto_save_targets = []      # 자동 저장 대상 PPT 목록 (딕셔너리 리스트)
        self.save_interval_seconds = DEFAULT_SAVE_INTERVAL_SECONDS
        
        self._auto_save_job = None # To store root.after job ID

        self.load_config()
        self.init_ui()
        
        if self.com_initialized:
            self.refresh_active_ppt_list_ui() # 초기 호출
            self.schedule_auto_save_cycle() # 초기 호출
        else:
            self.status_var.set("PowerPoint 연동 불가. COM 초기화 실패.")

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def init_ui(self):
        # --- Styles ---
        style = ttk.Style()
        style.configure("TLabel", padding=5, font=('Helvetica', 10))
        style.configure("TButton", padding=5, font=('Helvetica', 10))
        style.configure("TEntry", padding=5, font=('Helvetica', 10))
        style.configure("TFrame", padding=10)

        # --- Main Panes ---
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)

        left_pane = ttk.Frame(main_pane, padding=10)
        right_pane = ttk.Frame(main_pane, padding=10)
        main_pane.add(left_pane, weight=1)
        main_pane.add(right_pane, weight=1)

        # --- Left Pane: Active PPTs ---
        ttk.Label(left_pane, text="현재 열려있는 PPT 목록", font=('Helvetica', 12, 'bold')).pack(pady=(0,10))
        
        self.active_ppt_listbox = tk.Listbox(left_pane, height=15, exportselection=False, selectmode=tk.SINGLE)
        self.active_ppt_listbox.pack(fill=tk.BOTH, expand=True, pady=(0,10))
        self.active_ppt_listbox.bind('<<ListboxSelect>>', self.on_active_ppt_select)

        self.add_to_save_button = ttk.Button(left_pane, text="자동 저장 목록에 추가 >>", command=self.add_to_auto_save_list, state=tk.DISABLED)
        self.add_to_save_button.pack(fill=tk.X, pady=(0,5))
        
        refresh_button = ttk.Button(left_pane, text="목록 새로고침", command=self.refresh_active_ppt_list_ui)
        refresh_button.pack(fill=tk.X)

        # --- Right Pane: Auto-Save Targets & Settings ---
        ttk.Label(right_pane, text="자동 저장 대상 PPT", font=('Helvetica', 12, 'bold')).pack(pady=(0,10))
        self.auto_save_listbox = tk.Listbox(right_pane, height=10, exportselection=False, selectmode=tk.SINGLE)
        self.auto_save_listbox.pack(fill=tk.BOTH, expand=True, pady=(0,10))
        self.auto_save_listbox.bind('<<ListboxSelect>>', self.on_auto_save_ppt_select)

        self.remove_from_save_button = ttk.Button(right_pane, text="<< 목록에서 제거", command=self.remove_from_auto_save_list, state=tk.DISABLED)
        self.remove_from_save_button.pack(fill=tk.X, pady=(0,20))

        ttk.Label(right_pane, text="자동 저장 간격 (초):", font=('Helvetica', 10, 'bold')).pack(pady=(5,0))
        self.interval_var = tk.StringVar(value=str(self.save_interval_seconds))
        self.interval_entry = ttk.Entry(right_pane, textvariable=self.interval_var, width=10)
        self.interval_entry.pack(pady=(0,5))
        
        set_interval_button = ttk.Button(right_pane, text="간격 설정 및 저장 시작/재시작", command=self.set_save_interval_and_restart_autosave)
        set_interval_button.pack(fill=tk.X)

        # --- Status Bar ---
        self.status_var = tk.StringVar(value="준비 완료.")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.populate_auto_save_listbox_ui()


    def on_active_ppt_select(self, event=None):
        selection_indices = self.active_ppt_listbox.curselection()
        if not selection_indices:
            self.add_to_save_button.config(state=tk.DISABLED)
            return

        selected_index = selection_indices[0]
        if 0 <= selected_index < len(self.active_presentations_data):
            ppt_info = self.active_presentations_data[selected_index]
            # 저장 안된 파일(fullName is None)이거나 이미 자동 저장 목록에 있는 경우 추가 버튼 비활성화
            is_in_auto_save = any(target['fullName'] == ppt_info['fullName'] for target in self.auto_save_targets if target['fullName'])
            
            if ppt_info['fullName'] and not is_in_auto_save:
                self.add_to_save_button.config(state=tk.NORMAL)
            else:
                self.add_to_save_button.config(state=tk.DISABLED)
        else:
            self.add_to_save_button.config(state=tk.DISABLED)


    def on_auto_save_ppt_select(self, event=None):
        if self.auto_save_listbox.curselection():
            self.remove_from_save_button.config(state=tk.NORMAL)
        else:
            self.remove_from_save_button.config(state=tk.DISABLED)

    def refresh_active_ppt_list_ui(self):
        if not self.com_initialized:
            self.status_var.set("PowerPoint 연동 불가 (COM 초기화 실패).")
            if hasattr(self, '_active_list_refresh_job'): # 이전 작업 취소
                 if self._active_list_refresh_job:
                    self.root.after_cancel(self._active_list_refresh_job)
            self._active_list_refresh_job = self.root.after(ACTIVE_LIST_REFRESH_MS, self.refresh_active_ppt_list_ui)
            return

        self.status_var.set("활성 PPT 목록 새로고침 중...")
        self.active_presentations_data = get_open_presentations_info()
        self.active_ppt_listbox.delete(0, tk.END)

        current_auto_save_fullnames = {target['fullName'] for target in self.auto_save_targets if target['fullName']}

        if not self.active_presentations_data:
            self.active_ppt_listbox.insert(tk.END, "열려있는 PowerPoint 없음")
            self.active_ppt_listbox.itemconfig(0, {'fg': 'grey'})
        else:
            for index, ppt in enumerate(self.active_presentations_data):
                display_text = f"{ppt['name']}"
                if ppt['fullName']:
                    display_text += f" ({os.path.basename(ppt['fullName'])})" # 경로 대신 파일명만
                
                self.active_ppt_listbox.insert(tk.END, display_text)
                if ppt['fullName'] and ppt['fullName'] in current_auto_save_fullnames:
                    self.active_ppt_listbox.itemconfig(index, {'bg': 'lightgrey', 'fg': 'dim gray'}) # 이미 추가된 항목
                elif not ppt['fullName']:
                     self.active_ppt_listbox.itemconfig(index, {'fg': 'slate gray'}) # 저장 안된 항목

        self.on_active_ppt_select() # 선택 상태에 따라 버튼 상태 업데이트
        self.status_var.set("활성 PPT 목록 업데이트 완료.")
        
        if hasattr(self, '_active_list_refresh_job'): # 이전 작업 취소
             if self._active_list_refresh_job:
                self.root.after_cancel(self._active_list_refresh_job)
        self._active_list_refresh_job = self.root.after(ACTIVE_LIST_REFRESH_MS, self.refresh_active_ppt_list_ui)


    def add_to_auto_save_list(self):
        selection_indices = self.active_ppt_listbox.curselection()
        if not selection_indices:
            return
        
        selected_index = selection_indices[0]
        if 0 <= selected_index < len(self.active_presentations_data):
            ppt_to_add = self.active_presentations_data[selected_index]

            if not ppt_to_add['fullName']:
                messagebox.showwarning("추가 불가", f"'{ppt_to_add['name']}' 파일은 아직 저장되지 않아 자동 저장 목록에 추가할 수 없습니다. 먼저 수동으로 저장해주세요.")
                return

            # 중복 체크 (fullName 기준)
            if any(target['fullName'] == ppt_to_add['fullName'] for target in self.auto_save_targets):
                messagebox.showinfo("알림", f"'{ppt_to_add['name']}'은(는) 이미 자동 저장 목록에 있습니다.")
                return

            self.auto_save_targets.append({'name': ppt_to_add['name'], 'fullName': ppt_to_add['fullName']})
            self.populate_auto_save_listbox_ui()
            self.save_config()
            self.status_var.set(f"'{ppt_to_add['name']}' 자동 저장 목록에 추가됨.")
            self.refresh_active_ppt_list_ui() # 활성 목록 UI 갱신 (스타일 변경 위해)
        self.add_to_save_button.config(state=tk.DISABLED)


    def remove_from_auto_save_list(self):
        selection_indices = self.auto_save_listbox.curselection()
        if not selection_indices:
            return

        selected_index = selection_indices[0]
        if 0 <= selected_index < len(self.auto_save_targets):
            removed_ppt = self.auto_save_targets.pop(selected_index)
            self.populate_auto_save_listbox_ui()
            self.save_config()
            self.status_var.set(f"'{removed_ppt['name']}' 자동 저장 목록에서 제거됨.")
            self.refresh_active_ppt_list_ui() # 활성 목록 UI 갱신
        self.remove_from_save_button.config(state=tk.DISABLED)


    def populate_auto_save_listbox_ui(self):
        self.auto_save_listbox.delete(0, tk.END)
        if not self.auto_save_targets:
            self.auto_save_listbox.insert(tk.END, "자동 저장 대상 없음")
            self.auto_save_listbox.itemconfig(0, {'fg': 'grey'})
        else:
            for ppt in self.auto_save_targets:
                display_text = f"{ppt['name']}"
                if ppt['fullName']:
                    display_text += f" ({os.path.basename(ppt['fullName'])})"
                self.auto_save_listbox.insert(tk.END, display_text)
        self.on_auto_save_ppt_select()


    def set_save_interval_and_restart_autosave(self):
        try:
            new_interval = int(self.interval_var.get())
            if new_interval <= 0:
                messagebox.showerror("오류", "저장 간격은 0보다 큰 정수여야 합니다.")
                return
            self.save_interval_seconds = new_interval
            self.save_config()
            self.status_var.set(f"저장 간격이 {self.save_interval_seconds}초로 설정되었습니다. 자동 저장 재시작 중...")
            
            # 기존 자동 저장 작업이 있다면 취소하고 새로 시작
            if self._auto_save_job:
                self.root.after_cancel(self._auto_save_job)
            self.schedule_auto_save_cycle() 
            messagebox.showinfo("성공", f"저장 간격이 {self.save_interval_seconds}초로 설정되었습니다.")

        except ValueError:
            messagebox.showerror("오류", "저장 간격은 숫자로 입력해야 합니다.")


    def schedule_auto_save_cycle(self):
        if not self.com_initialized:
            self.status_var.set("자동 저장 불가 (COM 초기화 실패).")
            if self._auto_save_job:
                self.root.after_cancel(self._auto_save_job)
            self._auto_save_job = self.root.after(self.save_interval_seconds * 1000, self.schedule_auto_save_cycle)
            return

        if self.auto_save_targets:
            self.status_var.set(f"{self.save_interval_seconds}초 후 자동 저장 실행 예정...")
            # print(f"Scheduling auto-save in {self.save_interval_seconds} seconds for {len(self.auto_save_targets)} items.")
        else:
            self.status_var.set("자동 저장 대상 없음. 대기 중...")
            # print(f"No targets to auto-save. Scheduling next check in {self.save_interval_seconds} seconds.")

        if hasattr(self, '_auto_save_job'): # 이전 작업 취소 (간격 변경 시 중요)
            if self._auto_save_job:
                self.root.after_cancel(self._auto_save_job)
        
        self._auto_save_job = self.root.after(self.save_interval_seconds * 1000, self.run_auto_save_once_and_reschedule)

    def run_auto_save_once_and_reschedule(self):
        if not self.com_initialized:
            self.schedule_auto_save_cycle() # 그냥 다음 스케줄로 넘김
            return

        if not self.auto_save_targets:
            self.status_var.set("자동 저장 대상 없음. 다음 주기로...")
            self.schedule_auto_save_cycle() # 다음 주기로
            return

        self.status_var.set(f"자동 저장 실행 중 ({len(self.auto_save_targets)}개 대상)...")
        print(f"--- 자동 저장 시작 ({len(self.auto_save_targets)}개) ---")
        saved_count = 0
        failed_count = 0
        for target_ppt in self.auto_save_targets:
            if target_ppt['fullName']: # fullName이 있는 경우에만 시도
                if save_ppt_by_fullname(target_ppt['fullName']):
                    saved_count +=1
                else:
                    failed_count +=1
                    print(f"저장 실패: {target_ppt['name']} ({target_ppt['fullName']})")
            else:
                print(f"건너뜀 (경로 없음): {target_ppt['name']}")
        
        self.status_var.set(f"자동 저장 완료: 성공 {saved_count}개, 실패 {failed_count}개.")
        print(f"--- 자동 저장 완료: 성공 {saved_count}, 실패 {failed_count} ---")
        
        self.schedule_auto_save_cycle() # 다음 저장 예약


    def load_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.auto_save_targets = config.get('auto_save_targets', [])
                    # fullName이 없는 항목 필터링 (혹시 모를 오류 방지)
                    self.auto_save_targets = [item for item in self.auto_save_targets if item.get('fullName')]
                    self.save_interval_seconds = config.get('save_interval_seconds', DEFAULT_SAVE_INTERVAL_SECONDS)
            else: # 기본값 사용 및 파일 생성
                self.save_config(is_initial=True)

        except (json.JSONDecodeError, IOError) as e:
            print(f"설정 파일 로드 오류: {e}. 기본값 사용.")
            self.auto_save_targets = []
            self.save_interval_seconds = DEFAULT_SAVE_INTERVAL_SECONDS
            if os.path.exists(CONFIG_FILE): # 손상된 파일일 수 있으므로 삭제 시도
                try:
                    os.remove(CONFIG_FILE)
                except OSError:
                    pass
            self.save_config(is_initial=True) # 새 파일 생성

    def save_config(self, is_initial=False):
        config_data = {
            'auto_save_targets': self.auto_save_targets,
            'save_interval_seconds': self.save_interval_seconds
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=4, ensure_ascii=False)
            if not is_initial:
                 self.status_var.set("설정이 저장되었습니다.")
        except IOError as e:
            messagebox.showerror("오류", f"설정 파일 저장에 실패했습니다: {e}")
            self.status_var.set("설정 파일 저장 실패!")
            
    def on_closing(self):
        # 자동 저장 작업 및 새로고침 작업 중지
        if hasattr(self, '_auto_save_job') and self._auto_save_job:
            self.root.after_cancel(self._auto_save_job)
        if hasattr(self, '_active_list_refresh_job') and self._active_list_refresh_job:
            self.root.after_cancel(self._active_list_refresh_job)
            
        # COM 해제
        if self.com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                print(f"COM 해제 중 오류: {e}")
        
        self.save_config() # 마지막으로 설정 저장
        self.root.destroy()


if __name__ == '__main__':
    main_root = tk.Tk()
    app = PPTAutoSaverApp(main_root)
    main_root.mainloop()