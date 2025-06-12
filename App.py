import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog, font as tkFont
import serial
import serial.tools.list_ports
import threading
import time
import pyautogui
import json
import os
import random
from collections import deque

pyautogui.FAILSAFE = False
PROFILES_DIR = "mouse_profiles"

DEFAULT_SETTINGS = {
    "HST": 360, "NMIN": 460, "NMAX": 550, "SPT": 600, "HPT": 700,
    "JDZ": 20, "CSP": 10, "SAD": 150, "JMT": 5, "JFR": 100, "JCB": 20,
    "JRC": 0.5, "JIR": 0.3, "JPA": 50,
}

class IntegratedMouthMouseApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Mouth Mouse Tuner, Trainer & Calibrator (CTk V7.2)")
        self.root.geometry("1050x850") # Can adjust if needed

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        base_font_family = "Segoe UI"
        try: tkFont.Font(family=base_font_family, size=10).actual()
        except tk.TclError: base_font_family = "Arial"

        self.font_normal = ctk.CTkFont(family=base_font_family, size=12)
        self.font_small = ctk.CTkFont(family=base_font_family, size=10)
        self.font_bold = ctk.CTkFont(family=base_font_family, size=12, weight="bold")
        self.font_header = ctk.CTkFont(family=base_font_family, size=14, weight="bold")
        self.font_pressure = ctk.CTkFont(family=base_font_family, size=20, weight="bold")
        self.font_log = ctk.CTkFont(family="Courier New", size=10)
        self.font_labelframe_title = ctk.CTkFont(family=base_font_family, size=12, weight="bold")
        # Using a slightly larger font for threshold labels for clarity
        self.font_canvas_threshold_text = tkFont.Font(family="Arial", size=9)


        self.ser = None; self.is_connected = False; self.stop_read_thread = threading.Event()
        self.params_tkvars = {}
        for key, value in DEFAULT_SETTINGS.items():
            if key in ["JRC", "JIR"]: self.params_tkvars[key] = tk.DoubleVar(value=value)
            else: self.params_tkvars[key] = tk.IntVar(value=value)
        self.current_pressure_tkvar = tk.StringVar(value="Pressure: N/A")
        self.current_profile_name = tk.StringVar(value="<Default Settings>")

        self.osk_open = False; self.OSK_ZONE_SIZE = 30; self.OSK_POLL_INTERVAL_MS = 100
        try: self.screen_width,self.screen_height=pyautogui.size()
        except Exception: self.screen_width,self.screen_height=1920,1080
        self.pyautogui_left_button_down=False; self.pyautogui_right_button_down=False

        self.trainer_score_display_var=tk.StringVar(value="Score: 0"); self.trainer_score_value=0
        self.trainer_target_hits=0; self.trainer_target_misses=0; self.trainer_target_active=False
        self.trainer_target_coords=None; self.trainer_target_id=None; self.trainer_click_target_id=None
        self.trainer_click_target_text_id=None; self.trainer_click_target_button_type_expected=None
        self.trainer_scroll_font=ctk.CTkFont(family=base_font_family,size=12)
        self.trainer_content_host_frame=None
        self.trainer_active_tk_canvas = None
        self.is_target_hit_and_waiting_for_respawn=False
        
        self.mouse_trail_points = deque(maxlen=20) 
        self.mouse_trail_ids = []
        self.TRAIL_UPDATE_INTERVAL_MS = 16 
        self._mouse_trail_job_id = None

        self.calibrating_action_name=tk.StringVar(value=""); self.calibration_samples=[]
        self.calibration_current_value_tkvar=tk.StringVar(value="Raw Pressure: ---")
        self.is_calibrating_arduino_mode=False; self.collected_calibration_data={}; self._calibration_collect_job=None
        self.pressure_history = []
        self.pressure_canvas_min_width = 450 
        self.pressure_canvas_min_height = 450 
        
        self.pressure_label_area_width = 70 
        self._calculate_pressure_label_area_width() 

        self.max_history_points = self.pressure_canvas_min_width - self.pressure_label_area_width 
        if self.max_history_points <=0: self.max_history_points = 1 
        
        self.joystick_x_centered_tkvar = tk.IntVar(value=0); self.joystick_y_centered_tkvar = tk.IntVar(value=0)
        self.joystick_canvas_min_width = 200; self.joystick_canvas_min_height = 200
        self.joystick_indicator_id = None; self.joystick_deadzone_viz_id = None
        self.joystick_movethresh_viz_id = None

        if not os.path.exists(PROFILES_DIR):
            try: os.makedirs(PROFILES_DIR)
            except OSError as e: print(f"Error creating profiles directory {PROFILES_DIR}: {e}")

        self.create_main_layout()
        self.populate_ports(); self.populate_profiles_dropdown()
        self.set_status("Disconnected. Select port and connect.")
        self.schedule_osk_check()
        self.root.bind("<Configure>", self._on_window_resize)

    def _calculate_pressure_label_area_width(self):
        if not hasattr(self, 'font_canvas_threshold_text') or not self.font_canvas_threshold_text:
            self.font_canvas_threshold_text = tkFont.Font(family="Arial", size=9)
        
        max_measured_w = 0
        font_to_measure = self.font_canvas_threshold_text
        threshold_keys = ["HPT", "SPT", "NMAX", "NMIN", "HST"] 
        
        for key in threshold_keys:
            label_sample_text = f"{key}: 1023" 
            max_measured_w = max(max_measured_w, font_to_measure.measure(label_sample_text))

        gap_to_graph_line = 5 
        padding_left_of_text = 3 
        calculated_width = max_measured_w + gap_to_graph_line + padding_left_of_text
        self.pressure_label_area_width = max(70, int(calculated_width))

    def _on_window_resize(self, event=None):
        if hasattr(self, 'tab_view') and self.tab_view.winfo_exists():
            current_tab = self.tab_view.get()
            if current_tab == "Calibrate Sensor" and self.is_calibrating_arduino_mode:
                if hasattr(self, 'pressure_visualizer_canvas') and self.pressure_visualizer_canvas.winfo_exists():
                    w = self.pressure_visualizer_canvas.winfo_width()
                    if w > 1 : 
                        label_area_width = self.pressure_label_area_width 
                        graph_area_width = w - label_area_width
                        if graph_area_width <=0 : graph_area_width = 1 
                        if abs(self.max_history_points - graph_area_width) > 5: 
                           self.max_history_points = graph_area_width
                           if self.max_history_points <=0: self.max_history_points = 1
                    self._update_pressure_visualizer() 
            elif current_tab == "Stick Control":
                if hasattr(self, 'joystick_canvas') and self.joystick_canvas.winfo_exists():
                    self._update_joystick_visualizer()

    def _get_themed_canvas_bg(self):
        appearance_mode = ctk.get_appearance_mode()
        if appearance_mode == "Dark":
            try: return ctk.ThemeManager.theme["CTkFrame"]["fg_color"][1]
            except: return "#2B2B2B"
        else:
            try: return ctk.ThemeManager.theme["CTkFrame"]["fg_color"][0]
            except: return "#DBDBDB"

    def create_main_layout(self):
        top_bar_frame = ctk.CTkFrame(self.root, fg_color="transparent"); top_bar_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10,5))
        self.create_connection_widgets(top_bar_frame)
        self.status_var = tk.StringVar()
        self.status_bar = ctk.CTkLabel(self.root, textvariable=self.status_var, font=self.font_normal, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        self.tab_view = ctk.CTkTabview(self.root, command=self.on_tab_change)
        self.tab_view.pack(expand=True, fill='both', padx=10, pady=(5,10))
        self.tab_view.add("Tuner & Profiles")
        self.tab_view.add("Trainer")
        self.tab_view.add("Calibrate Sensor")
        self.tab_view.add("Stick Control")
        self.create_tuner_widgets(self.tab_view.tab("Tuner & Profiles"))
        self.create_trainer_widgets(self.tab_view.tab("Trainer"))
        self.create_calibration_widgets(self.tab_view.tab("Calibrate Sensor"))
        self.create_joystick_control_widgets(self.tab_view.tab("Stick Control"))
        self.tab_view.set("Tuner & Profiles")

    def _create_labeled_frame(self, parent, title_text, **kwargs):
        fg_color_main = kwargs.pop("fg_color", None) 
        outer_frame = ctk.CTkFrame(parent, border_width=1, fg_color=fg_color_main, **kwargs)
        title_label = ctk.CTkLabel(outer_frame, text=title_text, font=self.font_labelframe_title, anchor="w")
        title_label.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(5,2))
        content_frame = ctk.CTkFrame(outer_frame, fg_color="transparent")
        content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=(0,5))
        return outer_frame, content_frame

    def create_connection_widgets(self, parent_frame):
        conn_lf_outer, conn_frame = self._create_labeled_frame(parent_frame, "Serial Connection")
        conn_lf_outer.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X)
        ctk.CTkLabel(conn_frame, text="Port:", font=self.font_normal).pack(side=tk.LEFT, padx=(5,0), pady=5)
        self.port_combo = ctk.CTkComboBox(conn_frame, width=180, state="readonly", font=self.font_normal); self.port_combo.pack(side=tk.LEFT, padx=5, pady=5)
        self.connect_button = ctk.CTkButton(conn_frame, text="Connect", command=self.toggle_connect, font=self.font_bold, width=100); self.connect_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.refresh_button = ctk.CTkButton(conn_frame, text="Refresh Ports", command=self.populate_ports, font=self.font_bold, width=120); self.refresh_button.pack(side=tk.LEFT, padx=5, pady=5)
        pressure_lf_outer, pressure_display_frame = self._create_labeled_frame(parent_frame, "Live Pressure (Avg)")
        pressure_lf_outer.pack(side=tk.LEFT, padx=10, pady=5, fill=tk.X, expand=True)
        self.pressure_label = ctk.CTkLabel(pressure_display_frame, textvariable=self.current_pressure_tkvar, font=self.font_pressure); self.pressure_label.pack(padx=10, pady=(3,6))

    def create_tuner_widgets(self, parent_tab):
        main_tuner_pane = ctk.CTkFrame(parent_tab, fg_color="transparent"); main_tuner_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
        params_lf_outer, params_content_frame = self._create_labeled_frame(main_tuner_pane, "Parameters")
        params_lf_outer.pack(padx=5, pady=5, fill=tk.X, side=tk.TOP)
        col1_frame = ctk.CTkFrame(params_content_frame, fg_color="transparent")
        col1_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        col2_frame = ctk.CTkFrame(params_content_frame, fg_color="transparent")
        col2_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        params_content_frame.columnconfigure(0, weight=1); params_content_frame.columnconfigure(1, weight=1)
        row_idx_col1 = 0
        ctk.CTkLabel(col1_frame, text="Pressure Thresholds:", font=self.font_bold).grid(row=row_idx_col1, column=0, columnspan=3, pady=(5,2), sticky="w", padx=5)
        row_idx_col1=self.create_param_slider_widget(col1_frame, "Hard Sip (HST):", self.params_tkvars["HST"], "HST", 0, 1023, row_idx_col1)
        row_idx_col1=self.create_param_slider_widget(col1_frame, "Neutral Min (NMIN):", self.params_tkvars["NMIN"], "NMIN", 0, 1023, row_idx_col1)
        row_idx_col1=self.create_param_slider_widget(col1_frame, "Neutral Max (NMAX):", self.params_tkvars["NMAX"], "NMAX", 0, 1023, row_idx_col1)
        row_idx_col1=self.create_param_slider_widget(col1_frame, "Soft Puff (SPT):", self.params_tkvars["SPT"], "SPT", 0, 1023, row_idx_col1)
        row_idx_col1=self.create_param_slider_widget(col1_frame, "Hard Puff (HPT):", self.params_tkvars["HPT"], "HPT", 0, 1023, row_idx_col1)
        row_idx_col2 = 0
        ctk.CTkLabel(col2_frame, text="General Settings:", font=self.font_bold).grid(row=row_idx_col2, column=0, columnspan=3, pady=(5,2), sticky="w", padx=5)
        row_idx_col2=self.create_param_slider_widget(col2_frame, "Joystick Deadzone (JDZ%):", self.params_tkvars["JDZ"], "JDZ", 0, 100, row_idx_col2)
        row_idx_col2=self.create_param_slider_widget(col2_frame, "Cursor Speed (CSP):", self.params_tkvars["CSP"], "CSP", 1, 50, row_idx_col2)
        row_idx_col2=self.create_param_slider_widget(col2_frame, "Soft Action Delay ms (SAD):", self.params_tkvars["SAD"], "SAD", 0, 1000, row_idx_col2)
        col1_frame.columnconfigure(1, weight=1); col2_frame.columnconfigure(1, weight=1)
        self.apply_button = ctk.CTkButton(params_content_frame, text="Apply All Settings to Arduino", command=self.apply_all_settings, state=tk.DISABLED, font=self.font_bold)
        self.apply_button.grid(row=max(row_idx_col1, row_idx_col2)+1, column=0, columnspan=2, pady=20, padx=5, sticky="ew") 
        profile_lf_outer, profile_content_frame = self._create_labeled_frame(main_tuner_pane, "Profiles")
        profile_lf_outer.pack(padx=5, pady=10, fill=tk.X, side=tk.TOP)
        self.create_profile_widgets_content(profile_content_frame)

    def create_param_slider_widget(self, parent, label_text, tk_var, param_key, from_, to_, row_idx, desc_text=None):
        ctk.CTkLabel(parent, text=label_text, font=self.font_normal).grid(row=row_idx+1, column=0, padx=5, pady=(5,0), sticky="w") 
        is_float = isinstance(tk_var, tk.DoubleVar)
        slider_cmd = lambda val, v=tk_var, k=param_key: self._slider_update_wrapper(val, v, k, is_float)
        slider = ctk.CTkSlider(parent, from_=from_, to=to_, variable=tk_var, width=180, command=slider_cmd)
        slider.grid(row=row_idx+1, column=1, padx=5, pady=(5,0), sticky="ew")
        entry = ctk.CTkEntry(parent, textvariable=tk_var, width=50, font=self.font_normal)
        entry.grid(row=row_idx+1, column=2, padx=5, pady=(5,0))
        current_row = row_idx + 1 
        if desc_text:
            current_row +=1 
            try: wrap_len = parent.winfo_width() - 20 if parent.winfo_exists() and parent.winfo_width() > 20 else 200
            except: wrap_len = 200
            desc_label = ctk.CTkLabel(parent, text=desc_text, font=self.font_small, text_color=("gray30", "gray70"), wraplength=wrap_len, justify=tk.LEFT, anchor="w")
            desc_label.grid(row=current_row, column=0, columnspan=3, padx=15, pady=(0,5), sticky="w")
            
        parent.columnconfigure(1, weight=1)
        return current_row 

    def _slider_update_wrapper(self, value, tk_var_ref, param_key_ref, is_float_type):
        if is_float_type: tk_var_ref.set(round(value, 1))
        else: tk_var_ref.set(int(value))
        if param_key_ref in ["JDZ", "JMT", "HST", "NMIN", "NMAX", "SPT", "HPT"]: 
            if hasattr(self, 'tab_view') and self.tab_view.winfo_exists():
                current_tab = self.tab_view.get()
                if current_tab == "Stick Control" and param_key_ref in ["JDZ", "JMT"]:
                    self._update_joystick_visualizer()
                elif current_tab == "Calibrate Sensor" and param_key_ref in ["HST", "NMIN", "NMAX", "SPT", "HPT"]:
                    if self.is_calibrating_arduino_mode: 
                         self._update_pressure_visualizer()


    def create_profile_widgets_content(self, profile_frame):
        ctk.CTkLabel(profile_frame, text="Profile:", font=self.font_normal).grid(row=0, column=0, padx=5, pady=(5,3), sticky="w")
        self.profile_combo = ctk.CTkComboBox(profile_frame, variable=self.current_profile_name, width=250, state="readonly", font=self.font_normal)
        self.profile_combo.grid(row=0, column=1, padx=5, pady=(5,3), sticky="ew")
        profile_action_buttons_frame = ctk.CTkFrame(profile_frame, fg_color="transparent"); profile_action_buttons_frame.grid(row=0, column=2, rowspan=2, padx=(10,5), pady=3, sticky="ns")
        self.load_profile_button = ctk.CTkButton(profile_action_buttons_frame, text="Load Selected", command=self.load_selected_profile, font=self.font_bold); self.load_profile_button.pack(pady=(0,3), fill=tk.X)
        self.delete_profile_button = ctk.CTkButton(profile_action_buttons_frame, text="Delete Selected", command=self.delete_selected_profile, font=self.font_bold); self.delete_profile_button.pack(pady=3, fill=tk.X)
        save_buttons_frame = ctk.CTkFrame(profile_frame, fg_color="transparent"); save_buttons_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=(3,5), sticky="ew")
        self.save_profile_button = ctk.CTkButton(save_buttons_frame, text="Save to Selected", command=self.save_current_profile, font=self.font_bold); self.save_profile_button.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
        self.save_as_button = ctk.CTkButton(save_buttons_frame, text="Save As New...", command=self.save_profile_as, font=self.font_bold); self.save_as_button.pack(side=tk.LEFT, padx=(5,0), expand=True, fill=tk.X)
        self.load_defaults_button = ctk.CTkButton(profile_frame, text="Load Default Settings", command=self.load_default_settings, font=self.font_bold)
        self.load_defaults_button.grid(row=2, column=0, columnspan=3, padx=5, pady=(10,5), sticky="ew")
        profile_frame.columnconfigure(1, weight=1)

    def on_tab_change(self, selected_tab_name=None):
        if selected_tab_name is None and hasattr(self, 'tab_view') and self.tab_view.winfo_exists():
             selected_tab_name = self.tab_view.get()
        elif not hasattr(self, 'tab_view') or not self.tab_view.winfo_exists():
            return 
        if self._mouse_trail_job_id:
            if selected_tab_name != 'Trainer' or not (self.trainer_target_active and hasattr(self, 'trainer_active_tk_canvas') and self.trainer_active_tk_canvas):
                self.root.after_cancel(self._mouse_trail_job_id)
                self._mouse_trail_job_id = None
        if selected_tab_name != 'Calibrate Sensor' and self.is_calibrating_arduino_mode: self.stop_arduino_calibration_mode()
        if selected_tab_name == 'Tuner & Profiles': self.trainer_target_active=False; self._trainer_clear_canvas_content()
        elif selected_tab_name == 'Trainer':
            if hasattr(self,'instructions_label_trainer'):self.instructions_label_trainer.configure(text="Select a training mode.")
            if self.trainer_target_active and hasattr(self, 'trainer_active_tk_canvas') and self.trainer_active_tk_canvas: 
                if self._mouse_trail_job_id: self.root.after_cancel(self._mouse_trail_job_id) 
                self._mouse_trail_job_id = self.root.after(self.TRAIL_UPDATE_INTERVAL_MS, self._update_mouse_trail_loop) 
        elif selected_tab_name == 'Calibrate Sensor':
             if hasattr(self,'calibration_instructions_label'):
                 initial_calib_text = "Click 'Start Sensor Stream' then select an action." if not self.is_calibrating_arduino_mode else "Sensor stream active. Select an action or Stop Stream."
                 self.calibration_instructions_label.configure(text=initial_calib_text)
                 if self.is_calibrating_arduino_mode: self._update_pressure_visualizer()
        elif selected_tab_name == 'Stick Control':
            if hasattr(self, 'joystick_canvas') and self.joystick_canvas.winfo_exists(): self._update_joystick_visualizer()

    def create_trainer_widgets(self, parent_tab):
        trainer_controls_frame=ctk.CTkFrame(parent_tab, fg_color="transparent"); trainer_controls_frame.pack(pady=10,fill=tk.X, padx=5)
        ctk.CTkButton(trainer_controls_frame,text="Target Practice",command=self.start_target_practice, font=self.font_bold).pack(side=tk.LEFT,padx=5)
        ctk.CTkButton(trainer_controls_frame,text="Click Accuracy",command=self.start_click_accuracy, font=self.font_bold).pack(side=tk.LEFT,padx=5)
        ctk.CTkButton(trainer_controls_frame,text="Scroll Practice",command=self.start_scroll_practice, font=self.font_bold).pack(side=tk.LEFT,padx=5)
        self.trainer_score_label=ctk.CTkLabel(trainer_controls_frame,textvariable=self.trainer_score_display_var,font=self.font_header)
        self.trainer_score_display_var.set("Score: 0"); self.trainer_score_label.pack(side=tk.LEFT,padx=20)
        self.instructions_label_trainer=ctk.CTkLabel(parent_tab,text="Select a training mode.",justify=tk.CENTER,font=self.font_bold)
        self.instructions_label_trainer.pack(fill=tk.X,pady=5, padx=10)
        self.trainer_canvas_area_host=ctk.CTkFrame(parent_tab, border_width=1)
        self.trainer_canvas_area_host.pack(fill=tk.BOTH,expand=True,padx=10,pady=5)

    def create_calibration_widgets(self, parent_tab):
        calib_main_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        calib_main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        calib_main_frame.columnconfigure(0, weight=1)
        calib_main_frame.rowconfigure(0, weight=0) 
        calib_main_frame.rowconfigure(1, weight=3)  
        calib_main_frame.rowconfigure(2, weight=1)  
        calib_main_frame.rowconfigure(3, weight=0)  

        top_section_frame = ctk.CTkFrame(calib_main_frame, fg_color="transparent")
        top_section_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.calibration_instructions_label=ctk.CTkLabel(top_section_frame,text="Click 'Start Sensor Stream' then select an action.",justify=tk.CENTER,font=self.font_bold)
        self.calibration_instructions_label.pack(pady=5, fill=tk.X, padx=5)
        
        live_pressure_label_calib=ctk.CTkLabel(top_section_frame,textvariable=self.calibration_current_value_tkvar,font=self.font_pressure)
        live_pressure_label_calib.pack(pady=10)
        
        arduino_mode_frame=ctk.CTkFrame(top_section_frame, fg_color="transparent")
        arduino_mode_frame.pack(pady=5)
        self.start_arduino_calib_button=ctk.CTkButton(arduino_mode_frame,text="Start Sensor Stream",command=self.start_arduino_calibration_mode,width=180, font=self.font_bold)
        self.start_arduino_calib_button.pack(side=tk.LEFT,padx=5)
        self.stop_arduino_calib_button=ctk.CTkButton(arduino_mode_frame,text="Stop Sensor Stream",command=self.stop_arduino_calibration_mode,state=tk.DISABLED,width=180, font=self.font_bold)
        self.stop_arduino_calib_button.pack(side=tk.LEFT,padx=5)

        viz_lf_outer, viz_content_frame = self._create_labeled_frame(calib_main_frame, "Live Pressure Visualizer")
        viz_lf_outer.grid(row=1, column=0, sticky="nsew", pady=(0, 10), padx=5)
        canvas_bg = self._get_themed_canvas_bg()
        self.pressure_visualizer_canvas = tk.Canvas(viz_content_frame, width=self.pressure_canvas_min_width,
                                                    height=self.pressure_canvas_min_height, bg=canvas_bg,
                                                    highlightthickness=0)
        self.pressure_visualizer_canvas.pack(pady=5, padx=5, expand=True, fill=tk.BOTH, anchor=tk.CENTER)

        actions_and_log_frame = ctk.CTkFrame(calib_main_frame, fg_color="transparent")
        actions_and_log_frame.grid(row=2, column=0, sticky="nsew", pady=(0,5), padx=0) 

        actions_lf_outer, actions_content_frame = self._create_labeled_frame(actions_and_log_frame, "Calibration Actions")
        actions_lf_outer.pack(fill=tk.X, pady=5, padx=5) 
        self.calibration_actions = ["Neutral", "Soft Sip", "Hard Sip", "Soft Puff", "Hard Puff"]
        self.action_buttons = {}
        for i, action_name in enumerate(self.calibration_actions):
            btn = ctk.CTkButton(actions_content_frame, text=f"Record {action_name}",
                                command=lambda name=action_name: self.start_collecting_samples(name),
                                state=tk.DISABLED, font=self.font_bold)
            btn.grid(row=i // 3, column=i % 3, padx=5, pady=5, sticky="ew")
            self.action_buttons[action_name] = btn
        actions_content_frame.columnconfigure((0, 1, 2), weight=1)

        log_lf_outer, log_content_frame = self._create_labeled_frame(actions_and_log_frame, "Calibration Log & Results")
        log_lf_outer.pack(fill=tk.X, pady=5, padx=5) 
        self.calib_log_text = ctk.CTkTextbox(log_content_frame, height=100, wrap=tk.WORD, font=self.font_log,
                                             activate_scrollbars=True)
        self.calib_log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.calib_log_text.insert(tk.END, "Calibration Log:\n");
        self.calib_log_text.configure(state=tk.DISABLED)
        
        self.analyze_button = ctk.CTkButton(calib_main_frame, text="Analyze Data & Suggest Thresholds",
                                            command=self.analyze_calibration_data, state=tk.DISABLED,
                                            font=self.font_bold)
        self.analyze_button.grid(row=3, column=0, pady=(5,10), padx=5, sticky="ew")


    def create_joystick_control_widgets(self, parent_tab):
        main_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        main_frame.columnconfigure(0, weight=3); main_frame.columnconfigure(1, weight=1) 
        main_frame.rowconfigure(0, weight=1)    
        canvas_lf_outer, canvas_content_frame = self._create_labeled_frame(main_frame, "Joystick Position")
        canvas_lf_outer.grid(row=0, column=0, padx=(5,10), pady=5, sticky="nsew")
        canvas_bg = self._get_themed_canvas_bg()
        self.joystick_canvas = tk.Canvas(canvas_content_frame, width=self.joystick_canvas_min_width, height=self.joystick_canvas_min_height, bg=canvas_bg, relief=tk.FLAT, borderwidth=0, highlightthickness=0)
        self.joystick_canvas.pack(padx=10, pady=10, expand=True, fill=tk.BOTH, anchor=tk.CENTER)
        params_lf_outer, params_content_frame = self._create_labeled_frame(main_frame, "Stick Parameters")
        params_lf_outer.grid(row=0, column=1, padx=(0,5), pady=5, sticky="ns") 
        row_idx = 0 
        desc_texts = {
            "JMT": "Min stick movement (%) to trigger cursor motion.",
            "JFR": "Stick deflection (%) considered as full speed.",
            "JCB": "Stick must be within this % of center for clicks.",
            "JRC": "Time (s) for stick to auto-recenter (0=disabled).",
            "JIR": "Time (s) to hold in inner band for repeat action.",
            "JPA": "Pointer acceleration (0-100). 50 is linear."
        }
        row_idx = self.create_param_slider_widget(params_content_frame, "Move Threshold (JMT%):", self.params_tkvars["JMT"], "JMT", 0, 50, row_idx, desc_texts["JMT"])
        row_idx = self.create_param_slider_widget(params_content_frame, "Full Range (JFR%):", self.params_tkvars["JFR"], "JFR", 20, 100, row_idx, desc_texts["JFR"])
        row_idx = self.create_param_slider_widget(params_content_frame, "Click Deadband (JCB%):", self.params_tkvars["JCB"], "JCB", 0, 50, row_idx, desc_texts["JCB"])
        row_idx = self.create_param_slider_widget(params_content_frame, "Recenter Time (JRC s):", self.params_tkvars["JRC"], "JRC", 0.0, 5.0, row_idx, desc_texts["JRC"])
        row_idx = self.create_param_slider_widget(params_content_frame, "Inner Repeat (JIR s):", self.params_tkvars["JIR"], "JIR", 0.0, 2.0, row_idx, desc_texts["JIR"])
        row_idx = self.create_param_slider_widget(params_content_frame, "Acceleration (JPA%):", self.params_tkvars["JPA"], "JPA", 0, 100, row_idx, desc_texts["JPA"])
        
        row_idx += 1 
        ctk.CTkLabel(params_content_frame, text="Note: JDZ (Reset Zone) and CSP (Pointer Speed)\nare set in Tuner tab.",
                  font=self.font_small, justify=tk.LEFT, text_color=("gray30", "gray70")).grid(row=row_idx, column=0, columnspan=3, pady=(15,0), padx=5, sticky="w")

    def _update_joystick_visualizer(self):
        if not hasattr(self, 'joystick_canvas') or not self.joystick_canvas.winfo_exists(): return
        canvas = self.joystick_canvas; canvas_bg = self._get_themed_canvas_bg(); canvas.configure(bg=canvas_bg)
        w = canvas.winfo_width(); h = canvas.winfo_height()
        if w <=1 or h <=1: self.root.after(50, self._update_joystick_visualizer); return
        canvas.delete("all"); cx, cy = w / 2, h / 2; max_stick_deflection = 512.0 
        canvas_radius = (min(w, h) / 2) - 15; 
        if canvas_radius < 10: canvas_radius = 10
        canvas.create_oval(cx - canvas_radius, cy - canvas_radius, cx + canvas_radius, cy + canvas_radius, outline="gray40", width=1, dash=(4,4), tags="outer_boundary")
        joy_x = self.joystick_x_centered_tkvar.get(); joy_y = self.joystick_y_centered_tkvar.get()
        joy_x_clamped = max(-max_stick_deflection, min(max_stick_deflection, joy_x)); joy_y_clamped = max(-max_stick_deflection, min(max_stick_deflection, joy_y))
        indicator_x = cx + (joy_x_clamped / max_stick_deflection) * canvas_radius; indicator_y = cy - (joy_y_clamped / max_stick_deflection) * canvas_radius 
        jmt_percent = self.params_tkvars["JMT"].get() / 100.0; jmt_radius_pixels = jmt_percent * canvas_radius
        canvas.create_oval(cx - jmt_radius_pixels, cy - jmt_radius_pixels, cx + jmt_radius_pixels, cy + jmt_radius_pixels, outline="gray60", dash=(1,3), tags="jmt_circle")
        deadzone_percent = self.params_tkvars["JDZ"].get() / 100.0; deadzone_radius_pixels = deadzone_percent * canvas_radius
        canvas.create_oval(cx - deadzone_radius_pixels, cy - deadzone_radius_pixels, cx + deadzone_radius_pixels, cy + deadzone_radius_pixels, outline="skyblue", dash=(3,3), tags="jdz_circle")
        ind_r = 4; canvas.create_oval(indicator_x - ind_r, indicator_y - ind_r, indicator_x + ind_r, indicator_y + ind_r, fill="red", outline="white", tags="indicator")

    # --- Serial and Core Logic Methods ---
    def populate_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo.configure(values=ports)
        if ports:
            self.port_combo.set(ports[0])
        else:
            self.port_combo.set("")

    def toggle_connect(self):
        if not self.is_connected:
            port = self.port_combo.get()
            if not port:
                messagebox.showerror("Error", "No serial port selected.", parent=self.root)
                return
            try:
                self.ser = serial.Serial(port, 115200, timeout=0.1)
                self.is_connected = True
                self.connect_button.configure(text="Disconnect")
                if hasattr(self, 'apply_button'):
                    self.apply_button.configure(state=tk.NORMAL)
                self.set_status(f"Connected to {port}")
                self.stop_read_thread.clear()
                self.read_thread = threading.Thread(target=self.read_from_arduino, daemon=True)
                self.read_thread.start()
                self.send_command("H\n")
                self.apply_all_settings()
            except serial.SerialException as e:
                messagebox.showerror("Connection Error", str(e), parent=self.root)
                self.ser = None
                self.is_connected = False
        else:
            if self.is_calibrating_arduino_mode:
                self.stop_arduino_calibration_mode(silent=True)
            self.is_connected = False
            self.stop_read_thread.set()
            if hasattr(self, 'read_thread') and self.read_thread.is_alive():
                self.read_thread.join(timeout=0.5)
            if self.ser and self.ser.is_open:
                self.ser.close()
            self.ser = None
            self.connect_button.configure(text="Connect")
            if hasattr(self, 'apply_button'):
                self.apply_button.configure(state=tk.DISABLED)
            self.set_status("Disconnected")
            self.current_pressure_tkvar.set("Pressure: N/A")
            self.joystick_x_centered_tkvar.set(0)
            self.joystick_y_centered_tkvar.set(0)
            if hasattr(self, 'joystick_canvas') and self.joystick_canvas.winfo_exists():
                self._update_joystick_visualizer()

    def send_command(self, command):
        if self.ser and self.ser.is_open:
            try:
                if not command.endswith('\n'): command += '\n'
                self.ser.write(command.encode('utf-8'))
            except serial.SerialException as e:
                self.set_status(f"Send Error: {e}")
                self.handle_serial_error_disconnect()

    def apply_all_settings(self):
        if not self.is_connected:
            messagebox.showwarning("Not Connected", "Connect to Arduino first to apply settings.", parent=self.root)
            return
        self.set_status("Applying all settings to Arduino...")
        for key, tk_var in self.params_tkvars.items():
            value_to_send = tk_var.get()
            value_to_send_int = 0

            if isinstance(tk_var, tk.DoubleVar): 
                if key in ["JRC", "JIR"]: 
                    value_to_send_int = int(value_to_send * 10) 
                else: 
                    value_to_send_int = int(value_to_send)
            else: 
                 value_to_send_int = int(value_to_send)

            self.send_param_update(key, value_to_send_int)
            time.sleep(0.02) 
        self.set_status("All settings applied to Arduino.")
        if self.is_calibrating_arduino_mode and hasattr(self, 'pressure_visualizer_canvas'):
            self._update_pressure_visualizer()
        if hasattr(self, 'joystick_control_tab') and self.tab_view.winfo_exists() and self.tab_view.get() == "Stick Control":
            self._update_joystick_visualizer()

    def send_param_update(self, param_key, value):
        self.send_command(f"SET_{param_key}:{value}\n")


    def read_from_arduino(self):
        while not self.stop_read_thread.is_set():
            if not self.ser or not self.ser.is_open: break
            try:
                if self.ser.in_waiting > 0:
                    line = self.ser.readline().decode('utf-8', errors='ignore').strip()
                    if line:
                        if self.is_calibrating_arduino_mode and line.startswith("CALIB_P:"):
                            try:
                                value = int(line.split(":")[1])
                                self.calibration_current_value_tkvar.set(f"Raw Pressure: {value}")
                                self.pressure_history.append(value)
                                if len(self.pressure_history) > self.max_history_points:
                                    self.pressure_history.pop(0)
                                self.root.after(0, self._update_pressure_visualizer)
                                if self.calibrating_action_name.get():
                                    self.calibration_samples.append(value)
                            except (ValueError, IndexError): pass 
                        elif line.startswith("P:"):
                            if not self.is_calibrating_arduino_mode:
                                try: self.current_pressure_tkvar.set(f"Pressure: {line.split(':')[1]}")
                                except IndexError: pass
                        elif line.startswith("JOY:"):
                            try:
                                parts = line.split(':')[1].split(',')
                                joy_x = int(parts[0])
                                joy_y = int(parts[1])
                                self.joystick_x_centered_tkvar.set(joy_x)
                                self.joystick_y_centered_tkvar.set(joy_y)
                                if hasattr(self, 'tab_view') and self.tab_view.winfo_exists() and self.tab_view.get() == "Stick Control":
                                    self.root.after(0, self._update_joystick_visualizer)
                            except (IndexError, ValueError, tk.TclError): pass 
                        elif line.startswith("ACK:") or line.startswith("ERR:") or line.startswith("CMD_RECV:"):
                            self.set_status(f"Arduino: {line}")
                        else:
                            if not self.is_calibrating_arduino_mode:
                                self.handle_mouse_command_from_arduino(line)
            except serial.SerialException:
                self.root.after(0, self.handle_serial_error_disconnect)
                break
            except Exception as e: 
                if not self.stop_read_thread.is_set():
                    print(f"Read thread error: {e}")
            time.sleep(0.005) 

    def handle_mouse_command_from_arduino(self,line):
        parts=line.split(',');cmd=parts[0]
        try:
            if cmd=='MOVE' and len(parts)==3:pyautogui.moveRel(int(parts[1]),int(parts[2]),duration=0)
            elif cmd=='LEFT_CLICK_DOWN':pyautogui.mouseDown(button='left');self.pyautogui_left_button_down=True
            elif cmd=='LEFT_CLICK_UP':pyautogui.mouseUp(button='left');self.pyautogui_left_button_down=False
            elif cmd=='RIGHT_CLICK_DOWN':pyautogui.mouseDown(button='right');self.pyautogui_right_button_down=True
            elif cmd=='RIGHT_CLICK_UP':pyautogui.mouseUp(button='right');self.pyautogui_right_button_down=False
            elif cmd=='SCROLL_UP':pyautogui.scroll(20)
            elif cmd=='SCROLL_DOWN':pyautogui.scroll(-20)
        except Exception as e:self.set_status(f"PyAutoGUI Err: {str(e)[:50]}")

    def handle_serial_error_disconnect(self):
        if self.is_connected:self.toggle_connect()

    # --- Profile Methods ---
    def get_current_settings_dict(self): 
        return {key:tk_var.get() for key,tk_var in self.params_tkvars.items()}

    def apply_settings_from_dict(self,settings_dict,profile_name="<Loaded Profile>"):
        actual_settings=settings_dict.get("settings",settings_dict)
        for key,value in actual_settings.items():
            if key in self.params_tkvars:
                if key in ["JRC", "JIR"] and isinstance(self.params_tkvars[key], tk.DoubleVar):
                    self.params_tkvars[key].set(float(value))
                elif key in self.params_tkvars: 
                    self.params_tkvars[key].set(int(value))
        self.current_profile_name.set(profile_name)
        self.set_status(f"Settings from '{profile_name}' loaded into GUI.")
        if self.is_connected:self.apply_all_settings()

    def populate_profiles_dropdown(self):
        try:
            profile_files=[f for f in os.listdir(PROFILES_DIR) if f.endswith(".json")]
            profile_names=[os.path.splitext(f)[0] for f in profile_files]
            current_selection_is_valid_profile = self.current_profile_name.get() in profile_names
            if hasattr(self,'profile_combo'):
                self.profile_combo.configure(values=profile_names)
                if profile_names:
                    if current_selection_is_valid_profile and self.current_profile_name.get() != "<Default Settings Applied>":
                        self.profile_combo.set(self.current_profile_name.get())
                    else: self.profile_combo.set(profile_names[0]); self.current_profile_name.set(profile_names[0])
                else: self.profile_combo.set(""); self.current_profile_name.set("<Default Settings>")
        except Exception as e:self.set_status(f"Err listing profiles: {e}")

    def load_selected_profile(self):
        profile_name=self.profile_combo.get()
        if not profile_name or profile_name=="<Default Settings>": messagebox.showwarning("Load Profile","No profile selected.",parent=self.root);return
        self._load_profile_by_name(profile_name)

    def _load_profile_by_name(self,profile_name_to_load):
        filepath=os.path.join(PROFILES_DIR,f"{profile_name_to_load}.json")
        try:
            with open(filepath,'r') as f:settings_data=json.load(f)
            self.apply_settings_from_dict(settings_data,profile_name_to_load); 
            self.current_profile_name.set(profile_name_to_load)
        except FileNotFoundError:messagebox.showerror("Load Error",f"Profile '{profile_name_to_load}' not found.",parent=self.root)
        except Exception as e:messagebox.showerror("Load Error",f"Failed to load '{profile_name_to_load}': {e}",parent=self.root)

    def load_default_settings(self):
        if messagebox.askyesno("Load Defaults","Reset sliders to factory defaults?",parent=self.root):
            for key, value in DEFAULT_SETTINGS.items():
                if key in self.params_tkvars:
                    if isinstance(self.params_tkvars[key], tk.DoubleVar):
                        self.params_tkvars[key].set(float(value))
                    else:
                        self.params_tkvars[key].set(int(value))
            self.current_profile_name.set("<Default Settings Applied>")
            self.set_status(f"Default settings loaded into GUI.")
            if self.is_connected:self.apply_all_settings()

    def _save_profile_to_file(self,profile_name,settings_dict): 
        if not profile_name.strip() or profile_name=="<Default Settings>": messagebox.showerror("Save Profile","Invalid profile name.",parent=self.root);return False
        native_settings_dict = {}
        for key, tk_var in self.params_tkvars.items():
            native_settings_dict[key] = tk_var.get()
        data_to_save={"profile_name_meta":profile_name,"settings":native_settings_dict}
        filepath=os.path.join(PROFILES_DIR,f"{profile_name}.json")
        try:
            with open(filepath,'w') as f:json.dump(data_to_save,f,indent=2)
            self.set_status(f"Profile '{profile_name}' saved.");self.populate_profiles_dropdown()
            self.profile_combo.set(profile_name);self.current_profile_name.set(profile_name); return True
        except Exception as e:messagebox.showerror("Save Error",f"Could not save profile: {e}",parent=self.root);return False

    def save_current_profile(self):
        name=self.profile_combo.get()
        if not name or name=="<Default Settings>" or name=="<Default Settings Applied>":self.save_profile_as();return
        if messagebox.askyesno("Overwrite Profile",f"Overwrite existing profile '{name}'?",parent=self.root): self._save_profile_to_file(name, self.get_current_settings_dict())

    def save_profile_as(self):
        settings = self.get_current_settings_dict()
        dialog = ctk.CTkInputDialog(text="Enter new profile name:", title="Save As New Profile")
        name = dialog.get_input() 
        if name: self._save_profile_to_file(name, settings)

    def delete_selected_profile(self):
        name=self.profile_combo.get()
        if not name or name=="<Default Settings>" or name=="<Default Settings Applied>": messagebox.showwarning("Delete Profile","No saved profile selected.",parent=self.root);return
        if messagebox.askyesno("Confirm Delete",f"Delete profile '{name}'? This cannot be undone.",parent=self.root):
            try:
                os.remove(os.path.join(PROFILES_DIR,f"{name}.json")); self.set_status(f"Profile '{name}' deleted.")
                self.current_profile_name.set("<Default Settings>"); self.populate_profiles_dropdown()
            except Exception as e:messagebox.showerror("Delete Error",f"Failed to delete '{name}': {e}",parent=self.root)

    # --- Trainer Methods ---
    def _trainer_clear_canvas_content(self):
        self.trainer_target_active = False
        if self._mouse_trail_job_id:
            self.root.after_cancel(self._mouse_trail_job_id)
            self._mouse_trail_job_id = None
        if hasattr(self, 'trainer_content_host_frame') and self.trainer_content_host_frame and self.trainer_content_host_frame.winfo_exists():
            for widget in self.trainer_content_host_frame.winfo_children(): widget.destroy()
        self.trainer_active_tk_canvas = None
        self.mouse_trail_points.clear(); self.mouse_trail_ids = []
        if hasattr(self, 'instructions_label_trainer'): self.instructions_label_trainer.configure(text="Select a training mode.")
        if hasattr(self, 'trainer_score_display_var'): self.trainer_score_display_var.set("Score: 0")
        self.trainer_target_id, self.trainer_target_coords, self.is_target_hit_and_waiting_for_respawn = None, None, False

    def _setup_trainer_content_host(self):
        if hasattr(self, 'trainer_content_host_frame') and self.trainer_content_host_frame and self.trainer_content_host_frame.winfo_exists():
            for widget in self.trainer_content_host_frame.winfo_children(): widget.destroy()
        else: 
            if hasattr(self, 'trainer_canvas_area_host') and self.trainer_canvas_area_host.winfo_exists():
                self.trainer_content_host_frame = ctk.CTkFrame(self.trainer_canvas_area_host, fg_color="transparent")
                self.trainer_content_host_frame.pack(fill=tk.BOTH, expand=True)
            else: 
                trainer_tab = self.tab_view.tab("Trainer") 
                self.trainer_canvas_area_host=ctk.CTkFrame(trainer_tab, border_width=1)
                self.trainer_canvas_area_host.pack(fill=tk.BOTH,expand=True,padx=10,pady=5)
                self.trainer_content_host_frame = ctk.CTkFrame(self.trainer_canvas_area_host, fg_color="transparent")
                self.trainer_content_host_frame.pack(fill=tk.BOTH, expand=True)
        return self.trainer_content_host_frame

    def start_target_practice(self):
        self._trainer_clear_canvas_content(); self.trainer_target_active=True
        self.instructions_label_trainer.configure(text="Move cursor over the red target!")
        self.trainer_score_value=0; self.trainer_score_display_var.set(f"Score: {self.trainer_score_value}")
        host = self._setup_trainer_content_host()
        canvas_bg = self._get_themed_canvas_bg()
        self.trainer_active_tk_canvas = tk.Canvas(host, bg=canvas_bg, highlightthickness=0)
        self.trainer_active_tk_canvas.pack(fill=tk.BOTH, expand=True)
        self.trainer_target_size=30; self.trainer_target_id=None; self.trainer_target_coords=None; self.is_target_hit_and_waiting_for_respawn=False
        self.mouse_trail_points.clear(); self.mouse_trail_ids = []
        self.trainer_active_tk_canvas.after(100, self._trainer_spawn_hover_target)
        if self._mouse_trail_job_id: self.root.after_cancel(self._mouse_trail_job_id)
        self._mouse_trail_job_id = self.root.after(self.TRAIL_UPDATE_INTERVAL_MS, self._update_mouse_trail_loop)
        self._trainer_check_hover_loop()

    def _update_mouse_trail_loop(self):
        if not self.trainer_target_active or not hasattr(self, 'trainer_active_tk_canvas') or not self.trainer_active_tk_canvas or not self.trainer_active_tk_canvas.winfo_exists():
            self._mouse_trail_job_id = None 
            return
        canvas = self.trainer_active_tk_canvas
        if canvas.winfo_exists():
            mx_g, my_g = pyautogui.position()
            try:
                crx, cry = canvas.winfo_rootx(), canvas.winfo_rooty()
                rel_mx, rel_my = mx_g - crx, my_g - cry
                if 0 <= rel_mx <= canvas.winfo_width() and 0 <= rel_my <= canvas.winfo_height():
                    self.mouse_trail_points.append((rel_mx, rel_my))
                self._draw_mouse_trail()
            except tk.TclError: 
                self._mouse_trail_job_id = None
                return
        self._mouse_trail_job_id = self.root.after(self.TRAIL_UPDATE_INTERVAL_MS, self._update_mouse_trail_loop)

    def _trainer_spawn_hover_target(self):
        if not self.trainer_target_active or not hasattr(self,'trainer_active_tk_canvas') or not self.trainer_active_tk_canvas or not self.trainer_active_tk_canvas.winfo_exists():return
        canvas = self.trainer_active_tk_canvas
        if self.trainer_target_id:
            try: canvas.delete(self.trainer_target_id)
            except tk.TclError:pass
        canvas.update_idletasks() 
        w,h = canvas.winfo_width(), canvas.winfo_height()
        if w <= self.trainer_target_size or h <= self.trainer_target_size :
            if self.trainer_target_active: canvas.after(100,self._trainer_spawn_hover_target);return
        x1,y1=random.randint(0,max(0, w-self.trainer_target_size)),random.randint(0,max(0,h-self.trainer_target_size))
        self.trainer_target_coords=(x1,y1,x1+self.trainer_target_size,y1+self.trainer_target_size)
        self.trainer_target_id=canvas.create_oval(x1,y1,x1+self.trainer_target_size,y1+self.trainer_target_size,fill="red",outline="black", tags="target")
        self.is_target_hit_and_waiting_for_respawn=False

    def _draw_mouse_trail(self):
        if not hasattr(self, 'trainer_active_tk_canvas') or not self.trainer_active_tk_canvas or not self.trainer_active_tk_canvas.winfo_exists(): return
        canvas = self.trainer_active_tk_canvas
        for trail_id in self.mouse_trail_ids:
            try: canvas.delete(trail_id)
            except tk.TclError: pass
        self.mouse_trail_ids.clear()
        if len(self.mouse_trail_points) > 1:
            trail_color = "lightgrey" if ctk.get_appearance_mode() == "Dark" else "darkgrey"
            line_coords = []
            for point in self.mouse_trail_points: line_coords.extend(point)
            if line_coords:
                line_id = canvas.create_line(line_coords, fill=trail_color, width=2, tags="trail", smooth=True, splinesteps=5)
                self.mouse_trail_ids.append(line_id)

    def _trainer_check_hover_loop(self): 
        if not self.trainer_target_active or not hasattr(self,'trainer_active_tk_canvas') or not self.trainer_active_tk_canvas or not self.trainer_active_tk_canvas.winfo_exists():return
        canvas = self.trainer_active_tk_canvas
        rel_mx, rel_my = -1, -1 
        if canvas.winfo_exists():
            mx_g, my_g = pyautogui.position()
            try:
                crx, cry = canvas.winfo_rootx(), canvas.winfo_rooty()
                rel_mx, rel_my = mx_g - crx, my_g - cry
            except tk.TclError: 
                if self.trainer_target_active: self.root.after(50, self._trainer_check_hover_loop)
                return
        if self.is_target_hit_and_waiting_for_respawn:
            if self.trainer_target_active:self.root.after(50,self._trainer_check_hover_loop);return
        if self.trainer_target_id and self.trainer_target_coords and rel_mx != -1:
            try:
                if not canvas.coords(self.trainer_target_id): 
                    self.trainer_target_id,self.trainer_target_coords=None,None
                    if self.trainer_target_active:self.root.after(50,self._trainer_check_hover_loop);return
                if self.trainer_target_coords[0]<rel_mx<self.trainer_target_coords[2] and self.trainer_target_coords[1]<rel_my<self.trainer_target_coords[3]:
                    self.trainer_score_value+=1;self.trainer_score_display_var.set(f"Score: {self.trainer_score_value}")
                    canvas.itemconfig(self.trainer_target_id,fill="lightgreen");self.is_target_hit_and_waiting_for_respawn=True
                    hit_target_id_closure=self.trainer_target_id
                    self.trainer_target_id,self.trainer_target_coords=None,None 
                    canvas_ref = canvas 
                    def delayed_actions_after_hit():
                        if canvas_ref.winfo_exists():
                            try:
                                if hit_target_id_closure in canvas_ref.find_all(): canvas_ref.delete(hit_target_id_closure)
                            except tk.TclError:pass
                        if self.trainer_target_active:self._trainer_spawn_hover_target()
                    self.root.after(400,delayed_actions_after_hit)
            except tk.TclError:self.trainer_target_id,self.trainer_target_coords=None,None
            except Exception:pass 
        if self.trainer_target_active:self.root.after(50,self._trainer_check_hover_loop) 

    def start_click_accuracy(self):
        self._trainer_clear_canvas_content();self.trainer_target_active=True
        self.instructions_label_trainer.configure(text="Click target with correct button (LC/RC)!")
        self.trainer_target_hits,self.trainer_target_misses=0,0
        self.trainer_score_display_var.set(f"Hits: {self.trainer_target_hits} Misses: {self.trainer_target_misses}")
        host = self._setup_trainer_content_host()
        canvas_bg = self._get_themed_canvas_bg()
        self.trainer_active_tk_canvas = tk.Canvas(host, bg=canvas_bg, highlightthickness=0)
        self.trainer_active_tk_canvas.pack(fill=tk.BOTH, expand=True)
        self.trainer_click_target_id,self.trainer_click_target_text_id=None,None
        self.trainer_active_tk_canvas.after(50,self._trainer_spawn_click_target)
        self.trainer_active_tk_canvas.bind("<Button-1>",lambda e:self._trainer_on_canvas_click(e,"left"))
        self.trainer_active_tk_canvas.bind("<Button-3>",lambda e:self._trainer_on_canvas_click(e,"right"))

    def _trainer_spawn_click_target(self):
        if not self.trainer_target_active or not hasattr(self,'trainer_active_tk_canvas') or \
           not self.trainer_active_tk_canvas or not self.trainer_active_tk_canvas.winfo_exists():
            return
        canvas = self.trainer_active_tk_canvas
        if self.trainer_click_target_id: 
            try: canvas.delete(self.trainer_click_target_id)
            except tk.TclError: pass
        if self.trainer_click_target_text_id: 
            try: canvas.delete(self.trainer_click_target_text_id)
            except tk.TclError: pass
        canvas.update_idletasks();w,h=canvas.winfo_width(),canvas.winfo_height()
        if w<=60 or h<=60:
            if self.trainer_target_active: canvas.after(100,self._trainer_spawn_click_target);return
        rad=30;x,y=random.randint(rad,max(rad, w-rad)),random.randint(rad,max(rad, h-rad))
        self.trainer_click_target_button_type_expected=random.choice(["left","right"])
        clr="dodgerblue" if self.trainer_click_target_button_type_expected=="left" else "mediumpurple"
        txt="LC" if self.trainer_click_target_button_type_expected=="left" else "RC"
        self.trainer_click_target_id=canvas.create_oval(x-rad,y-rad,x+rad,y+rad,fill=clr,outline="black", tags="target")
        self.trainer_click_target_text_id=canvas.create_text(x,y,text=txt,fill="white",font=("Arial",16,"bold"), tags="target_text")

    def _trainer_on_canvas_click(self,event,clicked_button_type):
        if not self.trainer_target_active or not self.trainer_click_target_id or \
           not hasattr(self,'trainer_active_tk_canvas') or not self.trainer_active_tk_canvas or \
           not self.trainer_active_tk_canvas.winfo_exists():
            return
        canvas = self.trainer_active_tk_canvas
        try:
            coords=canvas.coords(self.trainer_click_target_id)
            if not coords:return 
        except tk.TclError: return 

        if coords[0]<event.x<coords[2] and coords[1]<event.y<coords[3]: 
            upd_id_closure,upd_txt_id_closure=self.trainer_click_target_id,self.trainer_click_target_text_id
            self.trainer_click_target_id,self.trainer_click_target_text_id=None,None 
            
            if clicked_button_type==self.trainer_click_target_button_type_expected:
                self.trainer_target_hits+=1; 
                if canvas.winfo_exists() and upd_id_closure in canvas.find_all(): canvas.itemconfig(upd_id_closure,fill="lightgreen")
            else:
                self.trainer_target_misses+=1; 
                if canvas.winfo_exists() and upd_id_closure in canvas.find_all(): canvas.itemconfig(upd_id_closure,fill="orangered")
            
            self.trainer_score_display_var.set(f"Hits: {self.trainer_target_hits} Misses: {self.trainer_target_misses}")
            canvas_ref = canvas 
            def d_respawn_click():
                if canvas_ref.winfo_exists():
                    try:
                        if upd_id_closure in canvas_ref.find_all():canvas_ref.delete(upd_id_closure)
                        if upd_txt_id_closure and upd_txt_id_closure in canvas_ref.find_all():canvas_ref.delete(upd_txt_id_closure)
                    except tk.TclError:pass
                if self.trainer_target_active:self._trainer_spawn_click_target()
            self.root.after(600,d_respawn_click)

    def start_scroll_practice(self):
        self._trainer_clear_canvas_content();self.trainer_target_active=True
        self.instructions_label_trainer.configure(text="Scroll text using soft sips/puffs.")
        self.trainer_score_display_var.set("Scroll Test Active")
        host = self._setup_trainer_content_host()
        scroll_text_widget=ctk.CTkTextbox(host, wrap=tk.WORD, height=200,
                                       font=self.trainer_scroll_font, spacing1=5,spacing2=2,spacing3=10,
                                       border_width=1, activate_scrollbars=True)
        scroll_text_widget.pack(side=tk.LEFT,fill=tk.BOTH,expand=True, padx=2, pady=2)
        content="Scroll Practice Area\n\n"+"Scroll down to read more.\n\n"+"\n".join([f"Section {i}:\nThis is a sample paragraph for CustomTkinter scroll test...\nIt should be long enough to demonstrate scrolling capabilities effectively.\n" for i in range(1,31)])
        scroll_text_widget.insert(tk.END,content);scroll_text_widget.configure(state=tk.DISABLED)
        self.trainer_active_tk_canvas = None 

    # --- OSK Methods ---
    def schedule_osk_check(self): 
        self.check_osk_trigger();self.root.after(self.OSK_POLL_INTERVAL_MS,self.schedule_osk_check)
    def toggle_osk(self):
        try:pyautogui.hotkey('win','ctrl','o');self.osk_open=not self.osk_open; state='Opened' if self.osk_open else 'Closed'; self.set_status(f"{state} On-Screen Keyboard")
        except Exception as e:self.set_status(f"OSK err: {e}")
    def check_osk_trigger(self):
        try:
            x,y=pyautogui.position()
            if x<self.OSK_ZONE_SIZE and y>self.screen_height-self.OSK_ZONE_SIZE:
                if not self.osk_open:self.toggle_osk()
            elif x>self.screen_width-self.OSK_ZONE_SIZE and y<self.OSK_ZONE_SIZE:
                if self.osk_open:self.toggle_osk()
        except Exception:pass
        
    # --- Calibration Methods ---
    def _add_to_calib_log(self, message):
        if hasattr(self, 'calib_log_text') and self.calib_log_text.winfo_exists():
            self.calib_log_text.configure(state=tk.NORMAL)
            self.calib_log_text.insert(tk.END, message + "\n")
            self.calib_log_text.see(tk.END)
            self.calib_log_text.configure(state=tk.DISABLED)

    def _update_pressure_visualizer(self):
        if not hasattr(self, 'pressure_visualizer_canvas') or \
           not self.pressure_visualizer_canvas.winfo_exists() or \
           not self.is_calibrating_arduino_mode:
            return
        canvas = self.pressure_visualizer_canvas
        canvas_bg = self._get_themed_canvas_bg()
        canvas.configure(bg=canvas_bg)
        w = canvas.winfo_width(); h = canvas.winfo_height()
        if w <=1 or h <=1: self.root.after(50, self._update_pressure_visualizer); return

        label_area_width = self.pressure_label_area_width 
        
        graph_area_x_start = label_area_width 
        graph_area_width = w - label_area_width
        if graph_area_width <=0 : graph_area_width = 1 
        
        if abs(self.max_history_points - graph_area_width) > 5 or self.max_history_points <= 1 :
           self.max_history_points = graph_area_width
           if self.max_history_points <=0: self.max_history_points = 1


        canvas.delete("all") 
        y_padding = 10 
        def pressure_to_y(pressure_val):
            graph_height = h - (2 * y_padding)
            if graph_height <= 0: graph_height = 1 
            scaled_pressure = max(0, min(1023, pressure_val))
            return (h - y_padding) - ((scaled_pressure / 1023.0) * graph_height)
        
        threshold_colors = {"HPT": "orangered", "SPT": "gold", "NMAX": "lightgreen", "NMIN": "lightgreen", "HST": "deepskyblue"}
        text_color = "white" if ctk.get_appearance_mode() == "Dark" else "black"
        current_font = self.font_canvas_threshold_text
        sorted_threshold_keys = ["HPT", "SPT", "NMAX", "NMIN", "HST"]
        
        gap_between_text_and_graph = 5 

        for key in sorted_threshold_keys:
            if key in self.params_tkvars:
                value = self.params_tkvars[key].get(); y_coord = pressure_to_y(value)
                canvas.create_line(graph_area_x_start, y_coord, w, y_coord, fill=threshold_colors.get(key, "gray50"), width=1, dash=(4, 2))
                label_text = f"{key}: {value}" 
                canvas.create_text(graph_area_x_start - gap_between_text_and_graph, y_coord, text=label_text, fill=text_color, anchor="e", font=current_font)
        
        if self.pressure_history:
            points = []; 
            current_max_history = max(1, int(self.max_history_points)) 
            
            step_x = 0
            if current_max_history > 0 and graph_area_width > 0 :
                step_x = graph_area_width / current_max_history
            
            drawable_history = self.pressure_history[-current_max_history:]
            for i, pressure_val in enumerate(drawable_history):
                x = graph_area_x_start + (i * step_x)
                y = pressure_to_y(pressure_val)
                points.extend([x, y])
            if len(points) >= 4: canvas.create_line(points, fill="cyan", width=2, tags="pressure_line")


    def start_arduino_calibration_mode(self):
        if not self.is_connected: messagebox.showwarning("Not Connected", "Connect to Arduino first.", parent=self.root); return
        self.send_command("START_CALIBRATION\n"); self.is_calibrating_arduino_mode = True
        self.start_arduino_calib_button.configure(state=tk.DISABLED); self.stop_arduino_calib_button.configure(state=tk.NORMAL)
        for btn_key in self.action_buttons: self.action_buttons[btn_key].configure(state=tk.NORMAL)
        self._add_to_calib_log("Arduino calibration stream started."); self.calibration_instructions_label.configure(text="Sensor stream active. Select an action.")
        self.pressure_history = [] 
        self._update_pressure_visualizer() 

    def stop_arduino_calibration_mode(self, silent=False):
        if not self.is_connected and not silent : return 
        if self.is_calibrating_arduino_mode or silent: self.send_command("STOP_CALIBRATION\n")
        self.is_calibrating_arduino_mode = False
        if hasattr(self, 'start_arduino_calib_button'): 
            self.start_arduino_calib_button.configure(state=tk.NORMAL); self.stop_arduino_calib_button.configure(state=tk.DISABLED)
            if hasattr(self, 'action_buttons'): 
                for btn_key in self.action_buttons: self.action_buttons[btn_key].configure(state=tk.DISABLED)
            self.calibration_current_value_tkvar.set("Raw Pressure: ---")
            if not silent: self._add_to_calib_log("Arduino calibration stream stopped.")
            if hasattr(self,'calibration_instructions_label'): self.calibration_instructions_label.configure(text="Click 'Start Sensor Stream'.")
        if hasattr(self, 'pressure_visualizer_canvas') and self.pressure_visualizer_canvas.winfo_exists():
             self.pressure_visualizer_canvas.delete("all") 
        self.pressure_history = []

    def start_collecting_samples(self, action_name):
        if not self.is_calibrating_arduino_mode: messagebox.showinfo("Info", "Start Arduino stream first.", parent=self.root); return
        self.calibrating_action_name.set(action_name); self.calibration_samples = []
        self.calibration_instructions_label.configure(text=f"PERFORMING: {action_name.upper()}. Hold pressure for ~3s..."); self._add_to_calib_log(f"--- Recording for {action_name} ---")
        for btn in self.action_buttons.values(): btn.configure(state=tk.DISABLED)
        self.stop_arduino_calib_button.configure(state=tk.DISABLED)
        COLLECTION_DURATION_MS = 3000
        if self._calibration_collect_job: self.root.after_cancel(self._calibration_collect_job)
        self._calibration_collect_job = self.root.after(COLLECTION_DURATION_MS, self.finish_collecting_samples)

    def finish_collecting_samples(self):
        self._calibration_collect_job = None 
        action_name = self.calibrating_action_name.get()
        if action_name:
            self.collected_calibration_data[action_name] = list(self.calibration_samples)
            self._add_to_calib_log(f"Collected {len(self.calibration_samples)} samples for {action_name}.")
            self.calibrating_action_name.set("")
        if self.is_calibrating_arduino_mode: 
            for btn_key in self.action_buttons: self.action_buttons[btn_key].configure(state=tk.NORMAL)
            self.stop_arduino_calib_button.configure(state=tk.NORMAL)
            self.calibration_instructions_label.configure(text="Sensor stream active. Select an action or Stop Stream.")
        else: 
            if hasattr(self, 'calibration_instructions_label'): self.calibration_instructions_label.configure(text="Stream stopped. Start stream to record.")
        if self.collected_calibration_data and hasattr(self, 'analyze_button'): self.analyze_button.configure(state=tk.NORMAL)

    def analyze_calibration_data(self):
        if not self.collected_calibration_data:
            self._add_to_calib_log("No data collected.")
            return
        self._add_to_calib_log("\n--- Analysis & Suggested Thresholds ---")
        stats = {}
        for action, samples in self.collected_calibration_data.items():
            if samples:
                avg = sum(samples) // len(samples)
                m_min = min(samples)
                m_max = max(samples)
            else:
                avg, m_min, m_max = 0, 0, 0
            stats[action] = {"avg": avg, "min": m_min, "max": m_max, "count": len(samples)}
            self._add_to_calib_log(f"{action}: Avg={avg}, Min={m_min}, Max={m_max} (Cnt:{len(samples)})")

        sug = {k: self.params_tkvars[k].get() for k in DEFAULT_SETTINGS}
        
        MIN_ZONE_SEPARATION = 25
        
        # --- User-defined stable neutral range ---
        KNOWN_NEUTRAL_LOW = 425  # MODIFIED: User's typical low neutral
        KNOWN_NEUTRAL_HIGH = 520 # MODIFIED: User's typical high neutral
        
        if KNOWN_NEUTRAL_LOW >= KNOWN_NEUTRAL_HIGH: 
            self._add_to_calib_log(f"Warning: KNOWN_NEUTRAL_LOW ({KNOWN_NEUTRAL_LOW}) is not less than KNOWN_NEUTRAL_HIGH ({KNOWN_NEUTRAL_HIGH}). Using defaults for known range.")
            known_neutral_low_eff = DEFAULT_SETTINGS["NMIN"]
            known_neutral_high_eff = DEFAULT_SETTINGS["NMAX"]
        else:
            known_neutral_low_eff = KNOWN_NEUTRAL_LOW
            known_neutral_high_eff = KNOWN_NEUTRAL_HIGH

        sug["NMIN"] = known_neutral_low_eff
        sug["NMAX"] = known_neutral_high_eff
        
        neutral_data = stats.get("Neutral")
        if neutral_data and neutral_data["count"] > 0:
            self._add_to_calib_log(f"Observed 'Neutral' during calibration: Min={neutral_data['min']}, Avg={neutral_data['avg']}, Max={neutral_data['max']}")
            self._add_to_calib_log(f"Prioritizing known stable neutral range: {known_neutral_low_eff}-{known_neutral_high_eff} for NMIN/NMAX base.")
            if sug["NMIN"] >= sug["NMAX"]:
                 sug["NMIN"] = known_neutral_low_eff 
                 sug["NMAX"] = known_neutral_high_eff 
                 if sug["NMIN"] >= sug["NMAX"]: 
                     sug["NMIN"] = neutral_data["avg"] - 10 
                     sug["NMAX"] = neutral_data["avg"] + 10
        else: 
            self._add_to_calib_log(f"Warning: No 'Neutral' data collected. Using known stable range: {known_neutral_low_eff}-{known_neutral_high_eff} for NMIN/NMAX.")

        soft_sip_data = stats.get("Soft Sip")
        hard_sip_data = stats.get("Hard Sip")

        if soft_sip_data and soft_sip_data["count"] > 0:
            required_nmin_from_softsip = soft_sip_data["max"] + MIN_ZONE_SEPARATION
            sug["NMIN"] = max(sug["NMIN"], required_nmin_from_softsip) 
        
        if hard_sip_data and hard_sip_data["count"] > 0:
            sug["HST"] = hard_sip_data["avg"] 
            if soft_sip_data and soft_sip_data["count"] > 0:
                 sug["HST"] = min(sug["HST"], soft_sip_data["min"] - MIN_ZONE_SEPARATION) 
            sug["HST"] = min(sug["HST"], sug["NMIN"] - MIN_ZONE_SEPARATION) 
        elif soft_sip_data and soft_sip_data["count"] > 0: 
            sug["HST"] = soft_sip_data["min"] - MIN_ZONE_SEPARATION
        
        soft_puff_data = stats.get("Soft Puff")
        hard_puff_data = stats.get("Hard Puff")

        if soft_puff_data and soft_puff_data["count"] > 0:
            required_nmax_from_softpuff = soft_puff_data["min"] - MIN_ZONE_SEPARATION
            sug["NMAX"] = min(sug["NMAX"], required_nmax_from_softpuff) 
            sug["SPT"] = soft_puff_data["avg"]
        
        if hard_puff_data and hard_puff_data["count"] > 0:
            sug["HPT"] = hard_puff_data["avg"]
            current_spt_base = sug.get("SPT")
            if current_spt_base is None: 
                current_spt_base = sug.get("NMAX", DEFAULT_SETTINGS["NMAX"]) + MIN_ZONE_SEPARATION 
                sug["SPT"] = current_spt_base 
            sug["HPT"] = max(sug["HPT"], current_spt_base + MIN_ZONE_SEPARATION)
        elif soft_puff_data and soft_puff_data["count"] > 0:
            sug["HPT"] = sug.get("SPT", DEFAULT_SETTINGS["SPT"]) + MIN_ZONE_SEPARATION * 2

        for k_val in ["HST", "NMIN", "NMAX", "SPT", "HPT"]:
            sug[k_val] = max(0, min(1023, sug.get(k_val, DEFAULT_SETTINGS[k_val])))

        if sug["NMIN"] >= sug["NMAX"]:
            self._add_to_calib_log(f"Warning: NMIN ({sug['NMIN']}) >= NMAX ({sug['NMAX']}) after sip/puff adjustment. Re-centering neutral band.")
            neutral_center_target = (known_neutral_low_eff + known_neutral_high_eff) // 2
            if neutral_data and neutral_data["count"] > 0: 
                if known_neutral_low_eff >= known_neutral_high_eff : neutral_center_target = neutral_data["avg"]
            
            min_neutral_width = max(10, MIN_ZONE_SEPARATION // 2) 
            sug["NMIN"] = neutral_center_target - (min_neutral_width // 2)
            sug["NMAX"] = sug["NMIN"] + min_neutral_width 

            sug["HST"] = min(sug.get("HST", DEFAULT_SETTINGS["HST"]), sug["NMIN"] - MIN_ZONE_SEPARATION)
            sug["SPT"] = max(sug.get("SPT", DEFAULT_SETTINGS["SPT"]), sug["NMAX"] + MIN_ZONE_SEPARATION)
            sug["HPT"] = max(sug.get("HPT", DEFAULT_SETTINGS["HPT"]), sug.get("SPT",DEFAULT_SETTINGS["SPT"]) + MIN_ZONE_SEPARATION)

        sug["HST"] = max(0, min(sug.get("HST", DEFAULT_SETTINGS["HST"]), sug["NMIN"] - MIN_ZONE_SEPARATION, 1023 - 4 * MIN_ZONE_SEPARATION))
        sug["NMIN"] = max(sug["HST"] + MIN_ZONE_SEPARATION, min(sug.get("NMIN", DEFAULT_SETTINGS["NMIN"]), sug["NMAX"] - MIN_ZONE_SEPARATION, 1023 - 3 * MIN_ZONE_SEPARATION))
        sug["NMAX"] = max(sug["NMIN"] + MIN_ZONE_SEPARATION, min(sug.get("NMAX", DEFAULT_SETTINGS["NMAX"]), sug.get("SPT", DEFAULT_SETTINGS["SPT"]) - MIN_ZONE_SEPARATION, 1023 - 2 * MIN_ZONE_SEPARATION))
        sug["SPT"] = max(sug["NMAX"] + MIN_ZONE_SEPARATION, min(sug.get("SPT", DEFAULT_SETTINGS["SPT"]), sug.get("HPT", DEFAULT_SETTINGS["HPT"]) - MIN_ZONE_SEPARATION, 1023 - 1 * MIN_ZONE_SEPARATION))
        sug["HPT"] = max(sug["SPT"] + MIN_ZONE_SEPARATION, min(sug.get("HPT", DEFAULT_SETTINGS["HPT"]), 1023))
        
        if sug["NMIN"] >= sug["NMAX"]:
            sug["NMAX"] = sug["NMIN"] + MIN_ZONE_SEPARATION
            sug["SPT"] = max(sug["NMAX"] + MIN_ZONE_SEPARATION, sug.get("SPT", DEFAULT_SETTINGS["SPT"]))
            sug["HPT"] = max(sug["SPT"] + MIN_ZONE_SEPARATION, sug.get("HPT", DEFAULT_SETTINGS["HPT"]))

        for k_val in ["HST", "NMIN", "NMAX", "SPT", "HPT"]:
            sug[k_val] = int(max(0, min(1023, sug.get(k_val, DEFAULT_SETTINGS[k_val]))))

        self._add_to_calib_log("\nSuggested values for Tuner tab (Review & Apply Manually or via prompt):")
        for k, v_val in sug.items():
            if k in ["HST", "NMIN", "NMAX", "SPT", "HPT"]:
                self._add_to_calib_log(f"  {k}: {v_val}")

        if messagebox.askyesno("Apply Suggestions?", "Apply these suggested pressure thresholds to Tuner sliders?", parent=self.root):
            for k, v_val in sug.items():
                if k in self.params_tkvars and k in ["HST", "NMIN", "NMAX", "SPT", "HPT"]:
                    self.params_tkvars[k].set(v_val)
            self._add_to_calib_log("Pressure threshold suggestions applied to Tuner sliders.")
            if self.is_connected:
                self.apply_all_settings()

    # --- Utility Methods ---
    def set_status(self, message):
        if hasattr(self, 'status_var'): self.status_var.set(message)

    def on_closing(self):
        if self._mouse_trail_job_id:
            self.root.after_cancel(self._mouse_trail_job_id)
            self._mouse_trail_job_id = None
        if self._calibration_collect_job: 
            self.root.after_cancel(self._calibration_collect_job)
            self._calibration_collect_job = None
        if self.is_calibrating_arduino_mode: self.stop_arduino_calibration_mode(silent=True)
        self.trainer_target_active = False
        self.stop_read_thread.set()
        if self.is_connected: self.toggle_connect() 
        self.root.destroy()

if __name__ == "__main__":
    root = ctk.CTk()
    app = IntegratedMouthMouseApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()