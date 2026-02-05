# 出退勤管理システム - ユーティリティモジュール

import os
import json
import winsound

class ConfigManager:
    """設定管理クラス"""
    
    def __init__(self, config_path):
        self.config_path = config_path
    
    def load_config(self):
        """設定ファイルの読み込み"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if config.get('configured', False):
                        return {
                            'class_reader': config.get('class_reader', ''),
                            'meeting_reader': config.get('meeting_reader', ''),
                            'configured': True
                        }
            return None
        except Exception as e:
            print(f"設定読み込みエラー: {e}")
            return None
    
    def save_config(self, class_reader_name, meeting_reader_name):
        """設定ファイルの保存"""
        try:
            config = {
                'class_reader': class_reader_name,
                'meeting_reader': meeting_reader_name,
                'configured': True
            }
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"設定保存エラー: {e}")
            return False


class SoundManager:
    """音声管理クラス"""
    
    def __init__(self):
        self.sound_enabled = True
    
    def toggle_sound(self):
        """音声ON/OFF切り替え"""
        self.sound_enabled = not self.sound_enabled
        return self.sound_enabled
    
    def play_beep(self, beep_type="success"):
        """ビープ音再生"""
        if not self.sound_enabled:
            return
        
        try:
            if beep_type == "success":
                winsound.Beep(1000, 200)
                import time
                time.sleep(0.1)
                winsound.Beep(1000, 200)
            elif beep_type == "error":
                winsound.Beep(400, 500)
            elif beep_type == "card_detected":
                winsound.Beep(800, 150)
        except Exception as e:
            print(f"音声再生エラー: {e}")