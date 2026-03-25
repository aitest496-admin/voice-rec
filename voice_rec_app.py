import sys
import os
import threading
import numpy as np
import tkinter as tk
from tkinter import messagebox, font as tkfont
from datetime import datetime
import xml.etree.ElementTree as ET
import sounddevice as sd
import soundfile as sf
import subprocess
import urllib.request
import json
import uuid

import configparser

# ============================================================
# 設定ファイルの読み込み
# ============================================================
# 実行ファイル（またはスクリプト）の格納ディレクトリを取得
if getattr(sys, 'frozen', False):
    # PyInstallerでビルドされたEXEの場合
    app_dir = os.path.dirname(sys.executable)
else:
    # Pythonスクリプトとして実行された場合
    app_dir = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(app_dir, "config.ini")
config = configparser.ConfigParser()

# 設定ファイルが存在しない場合はデフォルトで作成
if not os.path.exists(CONFIG_FILE):
    config["Settings"] = {
        "ROOT_SAVE_DIR": r"C:\I-EXP",
        "API_KEY": "D5nVbJKZ3FiegUR2_Ck2jY_p5cxQBULp_wZfdJ8wghUTnqy8xtPLr_Bb4FxZXHFp"
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)
else:
    config.read(CONFIG_FILE, encoding="utf-8")
    
    # 既存のconfig.iniにAPI_KEYがない場合（互換性対応）追記する
    if not config.has_option("Settings", "API_KEY"):
        config.set("Settings", "API_KEY", "D5nVbJKZ3FiegUR2_Ck2jY_p5cxQBULp_wZfdJ8wghUTnqy8xtPLr_Bb4FxZXHFp")
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)

ROOT_SAVE_DIR = config.get("Settings", "ROOT_SAVE_DIR", fallback=r"C:\I-EXP")
API_KEY = config.get("Settings", "API_KEY", fallback="D5nVbJKZ3FiegUR2_Ck2jY_p5cxQBULp_wZfdJ8wghUTnqy8xtPLr_Bb4FxZXHFp")

# ============================================================
# その他の定数
# ============================================================
RELAY_PATH = r"C:\Actiongate\Relay.xml"
SAMPLE_RATE = 44100
CHANNELS = 1  # モノラル

SUMMARY_URL = "https://api.white-ai.com/v1/ai/f264672e-0210-01e2-0c25-a532dcaeda3a/predict"
SOAP_URL = "https://api.white-ai.com/v1/ai/8a3f0146-88a4-c515-05fb-b3fdf1966594/predict"


# ============================================================
# ユーティリティ関数
# ============================================================
def check_relay_file():
    """Relay.xml の存在チェック。なければ警告して終了。"""
    if not os.path.exists(RELAY_PATH):
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.update_idletasks()
        messagebox.showwarning("確認", "患者を呼び出してください", parent=root)
        root.destroy()
        sys.exit()


def get_patient_info():
    """Relay.xml から患者番号と氏名を取得。失敗時はコマンドライン引数にフォールバック。"""
    try:
        tree = ET.parse(RELAY_PATH)
        root_elem = tree.getroot()
        karute = root_elem.get("Karute")
        name = root_elem.get("Name", "")
        # KaruteString属性があればそれを使い、なければKaruteを代用する
        karute_string = root_elem.get("KaruteString") or karute
        
        if karute:
            return karute, name, karute_string
    except Exception:
        pass

    # フォールバック: コマンドライン引数
    if len(sys.argv) > 1:
        return sys.argv[1], "", sys.argv[1]

    return "TEST_USER", "", "TEST_USER"


def ensure_patient_folder(patient_id):
    """患者番号フォルダを作成し、パスを返す。"""
    folder = os.path.join(ROOT_SAVE_DIR, patient_id)
    os.makedirs(folder, exist_ok=True)
    return folder




# ============================================================
# メインアプリケーション
# ============================================================
class DentalApp:
    def __init__(self, root, patient_id, patient_name, karute_string):
        self.root = root
        self.patient_id = patient_id
        self.patient_name = patient_name
        self.karute_string = karute_string
        
        # 保存先フォルダは「KaruteString」を使用する
        self.save_folder = ensure_patient_folder(karute_string)

        # 録音用データ
        self.audio_frames = []
        self.is_recording = False
        self.is_paused = False
        self.stream = None

        self._setup_window()
        self._setup_ui()
        self._start_recording()

        # ×ボタンで閉じるときの処理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 起動直後にWindowsの音声入力(Win+H)を自動で開く
        self.root.after(500, self._trigger_windows_dictation)

    def _trigger_windows_dictation(self):
        """Windows標準の音声認識（Win+H）を起動する"""
        import ctypes
        import time
        try:
            # VK_LWIN = 0x5B, 'H' = 0x48
            ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x48, 0, 0, 0)
            time.sleep(0.05)
            ctypes.windll.user32.keybd_event(0x48, 0, 0x0002, 0)
            ctypes.windll.user32.keybd_event(0x5B, 0, 0x0002, 0)
        except Exception as e:
            print("Windows音声入力の自動起動に失敗しました:", e)

    # -------------------------------------------------------
    # ウィンドウ設定
    # -------------------------------------------------------
    def _setup_window(self):
        title_text = f"音声入力メモ - 患者番号: {self.patient_id}"
        if self.patient_name:
            title_text += f" ({self.patient_name})"
        self.root.title(title_text)
        self.root.geometry("1000x700")
        self.root.attributes('-topmost', True)
        self.root.configure(bg="#f0f4f8")
        self.root.resizable(True, True)

    # -------------------------------------------------------
    # UI 構築
    # -------------------------------------------------------
    def _setup_ui(self):
        # --- 上部フレーム: 患者番号 + 録音状態 ---
        top_frame = tk.Frame(self.root, bg="#2c3e50", padx=10, pady=8)
        top_frame.pack(fill=tk.X)

        label_font = tkfont.Font(family="Meiryo UI", size=13, weight="bold")
        display_text = f"患者番号: {self.patient_id}"
        if self.patient_name:
            display_text += f"　氏名: {self.patient_name}"
            
        patient_label = tk.Label(
            top_frame,
            text=display_text,
            font=label_font,
            fg="white",
            bg="#2c3e50",
        )
        patient_label.pack(side=tk.LEFT)

        self.rec_label = tk.Label(
            top_frame,
            text="🔴 録音中",
            font=tkfont.Font(family="Meiryo UI", size=11),
            fg="#e74c3c",
            bg="#2c3e50",
        )
        self.rec_label.pack(side=tk.RIGHT)

        # --- 説明ラベル ---
        info_frame = tk.Frame(self.root, bg="#f0f4f8", pady=4)
        info_frame.pack(fill=tk.X, padx=10)

        info_label = tk.Label(
            info_frame,
            text="Windows音声入力（Win + H）でメモを入力してください",
            font=tkfont.Font(family="Meiryo UI", size=9),
            fg="#7f8c8d",
            bg="#f0f4f8",
        )
        info_label.pack(anchor=tk.W)

        # --- 中央: メインコンテナ (左右2段組み PanedWindow) ---
        main_paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg="#bdc3c7", sashwidth=6)
        main_paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ====== 左側カラム (Frame) ======
        left_frame = tk.Frame(main_paned, bg="#f0f4f8")
        main_paned.add(left_frame, minsize=300)

        text_font = tkfont.Font(family="Meiryo UI", size=12)

        def create_text_panel(parent, title):
            frame = tk.Frame(parent, bg="#f0f4f8")
            lbl = tk.Label(frame, text=title, font=tkfont.Font(family="Meiryo UI", size=10, weight="bold"), bg="#f0f4f8", fg="#2c3e50")
            lbl.pack(anchor=tk.W, pady=(0, 2))
            
            text_widget = tk.Text(
                frame, font=text_font, wrap=tk.WORD, bg="white", fg="#2c3e50",
                insertbackground="#2c3e50", relief=tk.FLAT, bd=2,
                highlightthickness=1, highlightcolor="#3498db", highlightbackground="#bdc3c7",
                padx=8, pady=8, width=1, height=1
            )
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar = tk.Scrollbar(frame, command=text_widget.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            text_widget.config(yscrollcommand=scrollbar.set)
            
            return frame, text_widget

        # 左1. 音声テキスト
        frame1, self.text_area = create_text_panel(left_frame, "音声テキスト")
        frame1.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 5))

        # 左2. ボタン群 (固定高さ)
        btn_frame = tk.Frame(left_frame, bg="#f0f4f8")
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

        btn_font = tkfont.Font(family="Meiryo UI", size=12, weight="bold")
        
        self.pause_btn = tk.Button(
            btn_frame,
            text="⏸ 一時停止",
            font=btn_font,
            bg="#f39c12",
            fg="white",
            activebackground="#f1c40f",
            activeforeground="white",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2",
            command=self.toggle_pause,
        )
        self.pause_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.summarize_btn = tk.Button(
            btn_frame,
            text="📝 要約",
            font=btn_font,
            bg="#27ae60",
            fg="white",
            activebackground="#2ecc71",
            activeforeground="white",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2",
            command=self.trigger_summary,
        )
        self.summarize_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.soap_btn = tk.Button(
            btn_frame,
            text="📑 SOAP",
            font=btn_font,
            bg="#8e44ad",
            fg="white",
            activebackground="#9b59b6",
            activeforeground="white",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2",
            command=self.trigger_soap,
        )
        self.soap_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # 左3. 要約欄
        frame2, self.summary_area = create_text_panel(left_frame, "要約欄")
        frame2.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(5, 0))

        # ====== 右側カラム ======
        frame3, self.soap_area = create_text_panel(main_paned, "SOAP欄")
        main_paned.add(frame3, minsize=250)

        # SOAP用カラータグの設定
        self.soap_area.tag_config('s_tag', foreground='#b33939') # 赤系
        self.soap_area.tag_config('o_tag', foreground='#227093') # 青系
        self.soap_area.tag_config('a_tag', foreground='#218c74') # 緑系
        self.soap_area.tag_config('p_tag', foreground='#8c7ae6') # 紫系
        self.soap_area.tag_config('n_tag', foreground='#e67e22') # オレンジ系（Note等）

        # 自動フォーカス
        self.text_area.focus_set()

    # -------------------------------------------------------
    # 録音機能
    # -------------------------------------------------------
    def _audio_callback(self, indata, frames, time_info, status):
        """sounddevice のコールバック。録音データを蓄積する。"""
        if self.is_recording and not self.is_paused:
            self.audio_frames.append(indata.copy())

    # -------------------------------------------------------
    # API 連携処理
    # -------------------------------------------------------
    def _post_api(self, url, text):
        data = {
            "params": {"text": text},
            "userID": str(uuid.uuid4())
        }
        req = urllib.request.Request(
            url, 
            data=json.dumps(data).encode("utf-8"), 
            headers={"x-api-key": API_KEY, "Content-Type": "application/json"},
            method="POST"
        )
        try:
            with urllib.request.urlopen(req) as response:
                res_body = response.read().decode("utf-8")
                res_json = json.loads(res_body)
                return res_json.get("result", "")
        except Exception as e:
            return f"API連携エラー: {e}"

    def _save_api_text(self, folder_name, text_data):
        """指定したフォルダにAPIのレスポンスを保存する"""
        if not text_data or text_data.startswith("API連携エラー"):
            return
        
        target_dir = os.path.join(self.save_folder, folder_name)
        os.makedirs(target_dir, exist_ok=True)
        
        # 西暦月日秒 (YYYYMMDDHHMMSS)
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        file_path = os.path.join(target_dir, f"{timestamp}.txt")
        
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(text_data)
                
            # 上書き保存用にパスを保持
            if folder_name == "summary":
                self.last_summary_path = file_path
            elif folder_name == "SOAP":
                self.last_soap_path = file_path
        except Exception:
            pass

    def trigger_summary(self):
        self._save_session()
        text_content = self.text_area.get("1.0", tk.END).strip()
        if not text_content:
            messagebox.showwarning("警告", "テキストがありません", parent=self.root)
            return

        self.summary_area.delete("1.0", tk.END)
        self.summary_area.insert(tk.END, "要約を取得中...\n")
        self.summarize_btn.config(state=tk.DISABLED)
        self.soap_btn.config(state=tk.DISABLED)

        threading.Thread(target=self._call_summary_api, args=(text_content,), daemon=True).start()

    def _call_summary_api(self, text_content):
        summary = self._post_api(SUMMARY_URL, text_content)
        self.root.after(0, self._update_summary_ui, summary)

    def _update_summary_ui(self, summary):
        self.summary_area.delete("1.0", tk.END)
        self.summary_area.insert(tk.END, summary)
        self.summarize_btn.config(state=tk.NORMAL)
        self.soap_btn.config(state=tk.NORMAL)
        self._save_api_text("summary", summary)

    def trigger_soap(self):
        self._save_session()
        text_content = self.text_area.get("1.0", tk.END).strip()
        if not text_content:
            messagebox.showwarning("警告", "テキストがありません", parent=self.root)
            return

        self.summary_area.delete("1.0", tk.END)
        self.summary_area.insert(tk.END, "要約を取得中...\n")
        self.soap_area.delete("1.0", tk.END)
        self.soap_area.insert(tk.END, "待機中...\n")
        self.summarize_btn.config(state=tk.DISABLED)
        self.soap_btn.config(state=tk.DISABLED)

        threading.Thread(target=self._call_soap_api_sequence, args=(text_content,), daemon=True).start()

    def _call_soap_api_sequence(self, text_content):
        summary = self._post_api(SUMMARY_URL, text_content)
        self.root.after(0, self._update_summary_intermediate, summary)
        
        if summary.startswith("API連携エラー"):
            self.root.after(0, self._update_soap_ui, "要約の取得に失敗したため、SOAP分類を中断しました。")
            return
            
        soap_result = self._post_api(SOAP_URL, summary)
        self.root.after(0, self._update_soap_ui, soap_result)

    def _update_summary_intermediate(self, summary):
        self.summary_area.delete("1.0", tk.END)
        self.summary_area.insert(tk.END, summary)
        self.soap_area.delete("1.0", tk.END)
        self.soap_area.insert(tk.END, "SOAP分類を取得中...\n")
        self._save_api_text("summary", summary)

    def _update_soap_ui(self, soap_result):
        self.soap_area.delete("1.0", tk.END)
        
        current_tag = None
        import re
        for line in soap_result.split("\n"):
            line_upper = line.strip().upper()
            
            # 正規表現で「S:」「S@」「【S】」「[S]」等に幅広く対応
            m = re.match(r'^([SOAPN])[：:@＠\]】]', line_upper)
            if m:
                tag_char = m.group(1).lower()
                current_tag = f"{tag_char}_tag"
            elif line_upper in ["S", "O", "A", "P", "N"]:
                current_tag = f"{line_upper.lower()}_tag"
            
            if current_tag:
                self.soap_area.insert(tk.END, line + "\n", current_tag)
            else:
                self.soap_area.insert(tk.END, line + "\n")

        if self.soap_area.get("end-1c", tk.END) == "\n":
            self.soap_area.delete("end-1c", tk.END)
            
        self.summarize_btn.config(state=tk.NORMAL)
        self.soap_btn.config(state=tk.NORMAL)
        self._save_api_text("SOAP", soap_result)

    def toggle_pause(self):
        """録音の一時停止・再開を切り替える。"""
        if not self.is_recording:
            return
            
        self.is_paused = not self.is_paused
        
        if self.is_paused:
            self.pause_btn.config(text="▶ 録音再開", bg="#3498db", activebackground="#2980b9")
            self.rec_label.config(text="⏸ 一時停止中", fg="#f39c12")
        else:
            self.pause_btn.config(text="⏸ 一時停止", bg="#f39c12", activebackground="#f1c40f")
            self.rec_label.config(text="🔴 録音中", fg="#e74c3c")

    def _start_recording(self):
        """バックグラウンド録音を開始する。"""
        self.is_recording = True
        self.audio_frames = []
        try:
            self.stream = sd.InputStream(
                samplerate=SAMPLE_RATE,
                channels=CHANNELS,
                callback=self._audio_callback,
            )
            self.stream.start()
        except Exception as e:
            self.rec_label.config(text="⚠ 録音エラー", fg="#e67e22")
            messagebox.showerror(
                "録音エラー",
                f"マイクの初期化に失敗しました:\n{e}",
                parent=self.root,
            )

    def _stop_recording(self):
        """録音を停止する。"""
        self.is_recording = False
        if self.stream:
            try:
                self.stream.stop()
                self.stream.close()
            except Exception:
                pass
            self.stream = None

    # -------------------------------------------------------
    # 保存・終了
    # -------------------------------------------------------
    def _save_session(self):
        """録音を停止し、テキストと音声を保存する（アプリは終了しない）。"""
        if getattr(self, '_session_saved', False):
            return
            
        self._stop_recording()
        self._session_saved = True

        # 要約テキストと同じ命名規則 (YYYYMMDDHHMMSS)
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        messages = []

        # --- 音声保存 (RECフォルダ) ---
        if self.audio_frames:
            rec_folder = os.path.join(self.save_folder, "REC")
            os.makedirs(rec_folder, exist_ok=True)

            audio_data = np.concatenate(self.audio_frames, axis=0)
            wav_path = os.path.join(rec_folder, f"{timestamp}.wav")
            m4a_path = os.path.join(rec_folder, f"{timestamp}.m4a")
            
            # 1. 一時WAVとして保存
            sf.write(wav_path, audio_data, SAMPLE_RATE, format='WAV', subtype='PCM_16')
            
            # 2. FFmpegでM4Aに変換
            try:
                # Windows特有の設定（コマンドプロンプト画面を隠す）
                CREATE_NO_WINDOW = 0x08000000
                subprocess.run(
                    ["ffmpeg", "-y", "-i", wav_path, "-c:a", "aac", "-b:a", "128k", m4a_path],
                    check=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=CREATE_NO_WINDOW
                )
                # 変換成功時は一時WAVを削除
                if os.path.exists(m4a_path):
                    os.remove(wav_path)
                    messages.append(f"音声: {m4a_path}")
                else:
                    raise Exception("M4A file not created")
            except Exception:
                # 変換失敗時はWAVを残す
                messages.append(f"音声: {wav_path} (M4A変換失敗)")

        # --- ステータス更新 ---
        if messages:
            self.rec_label.config(text="💾 保存完了", fg="#27ae60")
        else:
            self.rec_label.config(text="保存データなし", fg="#7f8c8d")

    # -------------------------------------------------------
    # ウィンドウ閉じる処理
    # -------------------------------------------------------
    def _on_close(self):
        """ウィンドウを閉じるときに、要約とSOAPの内容を既存ファイルに上書き保存して終了する。"""
        self._stop_recording()
        
        # 要約テキストの上書き保存
        if getattr(self, "last_summary_path", None):
            summary_text = self.summary_area.get("1.0", tk.END).strip()
            if summary_text:
                try:
                    with open(self.last_summary_path, "w", encoding="utf-8") as f:
                        f.write(summary_text)
                except Exception:
                    pass

        # SOAPテキストの上書き保存
        if getattr(self, "last_soap_path", None):
            soap_text = self.soap_area.get("1.0", tk.END).strip()
            if soap_text:
                try:
                    with open(self.last_soap_path, "w", encoding="utf-8") as f:
                        f.write(soap_text)
                except Exception:
                    pass

        self.root.destroy()


# ============================================================
# エントリーポイント
# ============================================================
if __name__ == "__main__":
    # 1. Relay.xml の存在チェック
    check_relay_file()

    # 2. 患者情報を取得
    p_id, p_name, k_str = get_patient_info()

    # 3. メイン画面を起動
    root = tk.Tk()
    app = DentalApp(root, p_id, p_name, k_str)
    root.mainloop()