# -*- coding: utf-8 -*-
"""
多機能ファイル一括リネームツール「連番スーパー」
対話形式ですべての操作を実行できる統合版スクリプト
"""
import os
import re
import datetime
from pathlib import Path

# 将来的に必要になるライブラリ（現時点ではAI機能はAPIキーなしで動作しない）
try:
    from dotenv import load_dotenv
    import google.generativeai as genai
    from docx import Document
    from PIL import Image
except ImportError:
    print("警告: 必要なライブラリが不足しています。'pip install python-dotenv google-generativeai python-docx Pillow' を実行してください。")
    genai = None # AI機能を無効化

# --- 1. コアロジックを定義するクラス ---

class RenbanSuper:
    """ファイルリネーム処理を行うメインクラス"""

    def __init__(self, directory, dry_run=False, recursive=False, extensions=None, prefix="", suffix=""):
        self.directory = Path(directory)
        self.dry_run = dry_run
        self.recursive = recursive
        self.extensions = extensions
        self.prefix = prefix
        self.suffix = suffix

        if not self.directory.is_dir():
            raise FileNotFoundError(f"エラー: ディレクトリ '{directory}' が見つかりません。")

        # AIのセットアップ
        self.ai_model = self._setup_ai()

    def _setup_ai(self):
        """AIモデルの初期設定を行う"""
        if not genai: return None # ライブラリがなければNone
        
        # .envファイルから環境変数を読み込む（ファイルがなくてもエラーにならない）
        load_dotenv()
        api_key = os.getenv("GOOGLE_API_KEY")
        
        if not api_key:
            # AI機能が選択された時に警告を出す
            return None
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash-latest')
            return model
        except Exception as e:
            print(f"AIモデルの初期化に失敗しました: {e}")
            return None

    def _sanitize_filename(self, name):
        """ファイル名として不適切な文字を除去・置換する"""
        invalid_chars = r'[\\/:*?"<>|]'
        sanitized = re.sub(invalid_chars, '_', name)
        if name.startswith('.') and not sanitized.startswith('.'):
             sanitized = '.' + sanitized
        return sanitized

    def _get_files(self):
        """対象となるファイルリストを取得する"""
        files = []
        iterator = self.directory.rglob('*') if self.recursive else self.directory.glob('*')

        for path in iterator:
            if path.is_file():
                if self.extensions and path.suffix.lower() not in self.extensions:
                    continue
                files.append(path)
        
        return sorted(files)

    def _rename_file(self, old_path, new_name_base, extension):
        """実際にファイル名を変更する処理"""
        new_name = f"{self.prefix}{new_name_base}{self.suffix}{extension}"
        new_name = self._sanitize_filename(new_name)
        new_path = old_path.with_name(new_name)

        counter = 1
        while new_path.exists() and new_path != old_path:
            new_name = f"{self.prefix}{new_name_base}{self.suffix}_{counter}{extension}"
            new_name = self._sanitize_filename(new_name)
            new_path = old_path.with_name(new_name)
            counter += 1

        print(f"'{old_path.name}' -> '{new_path.name}'")

        if not self.dry_run:
            try:
                old_path.rename(new_path)
            except OSError as e:
                print(f"  エラー: '{old_path.name}' のリネームに失敗。理由: {e}")

    # --- 各種情報取得メソッド ---
    def get_file_author(self, filepath):
        ext = filepath.suffix.lower()
        try:
            if ext == '.docx':
                return Document(filepath).core_properties.author or "UnknownAuthor"
            elif ext in ['.jpg', '.jpeg', '.tiff']:
                with Image.open(filepath) as img:
                    exif_data = img._getexif()
                    if exif_data: return exif_data.get(315, "UnknownArtist")
        except Exception:
            return "AuthorNotAvailable"
        return "AuthorNotSupported"

    def get_ai_summary(self, filepath):
        if not self.ai_model:
            print("警告: AI機能はAPIキーが設定されていないため利用できません。")
            return "AI_Not_Available"
        
        try:
            if filepath.suffix.lower() in ['.txt', '.md', '.py', '.html', '.css', '.js']:
                content = filepath.read_text(encoding='utf-8', errors='ignore')[:10000]
                if not content.strip(): return "EmptyFile"
                prompt = f"以下のファイル内容を、ファイル名として使えるように3〜5単語程度の非常に短い英語で要約してください:\n\n---\n{content}\n---"
                response = self.ai_model.generate_content(prompt)
                return response.text.strip().replace(" ", "_")
            else:
                return "SummaryNotSupported"
        except Exception as e:
            return "AI_Error"

    def sort_files_by_ai(self, files):
        if not self.ai_model:
            print("警告: AI機能はAPIキーが設定されていないため利用できません。")
            return files

        print("AIによるファイルの並べ替えを実行中...")
        file_info = [f"[{i}] {f.name}" for i, f in enumerate(files)]
        
        try:
            prompt = f"以下のファイルリストを論理的な順序に並べ替えてください。回答はインデックス番号をカンマ区切りで出力してください(例: 3,0,1,2)。\n\n{chr(10).join(file_info)}"
            response = self.ai_model.generate_content(prompt)
            sorted_indices = [int(i.strip()) for i in response.text.strip().split(',')]
            
            if len(sorted_indices) == len(files):
                print("AIによる並べ替えが完了しました。")
                return [files[i] for i in sorted_indices]
            else:
                print("警告: AIが不正な順序を返しました。元の順序を使用します。")
                return files
        except Exception as e:
            print(f"AI並べ替えエラー: {e}。元の順序を使用します。")
            return files

    # --- 各リネームモードの実行関数 ---
    def run_simple(self, start=1, digits=3):
        files = self._get_files()
        for i, old_path in enumerate(files):
            new_name_base = f"{start + i:0{digits}d}"
            self._rename_file(old_path, new_name_base, old_path.suffix)

    def run_date(self, date_type='modified', date_format='%Y-%m-%d_%H%M%S'):
        files = self._get_files()
        for old_path in files:
            ts = old_path.stat().st_ctime if date_type == 'created' else old_path.stat().st_mtime
            new_name_base = datetime.datetime.fromtimestamp(ts).strftime(date_format)
            self._rename_file(old_path, new_name_base, old_path.suffix)

    def run_size(self):
        files = self._get_files()
        for old_path in files:
            new_name_base = f"size_{old_path.stat().st_size}B"
            self._rename_file(old_path, new_name_base, old_path.suffix)

    def run_author(self):
        files = self._get_files()
        for old_path in files:
            author = self.get_file_author(old_path)
            new_name_base = f"{author}_{old_path.stem}"
            self._rename_file(old_path, new_name_base, old_path.suffix)

    def run_ai_summary(self):
        files = self._get_files()
        for old_path in files:
            summary = self.get_ai_summary(old_path)
            self._rename_file(old_path, summary, old_path.suffix)
    
    def run_ai_sort(self, start=1, digits=3):
        files = self._get_files()
        sorted_files = self.sort_files_by_ai(files)
        for i, old_path in enumerate(sorted_files):
            new_name_base = f"{start + i:0{digits}d}"
            self._rename_file(old_path, new_name_base, old_path.suffix)

# --- 2. ユーザー対話のための補助関数 ---

def get_folder_path():
    """ユーザーにフォルダパスの入力を求め、有効なパスが入力されるまで繰り返す"""
    while True:
        path_str = input("変更したいフォルダのパスを入力してください: ")
        path = Path(path_str.strip().strip('"'))
        if path.is_dir():
            return path
        else:
            print("エラー: そのようなフォルダは見つかりません。")

def get_rename_mode():
    """リネームモードを選択させる"""
    print("\n希望するリネームのモードを選択してください:")
    modes = ['simple', 'date', 'size', 'author', 'ai-summary', 'ai-sort']
    for i, mode in enumerate(modes, 1):
        print(f"  {i}: {mode}")
    
    while True:
        try:
            choice = int(input("番号で選択してください: "))
            if 1 <= choice <= len(modes):
                return modes[choice - 1]
            else:
                print(f"エラー: 1から{len(modes)}の間の番号を選択してください。")
        except ValueError:
            print("エラー: 番号で入力してください。")

def get_simple_options():
    """simpleモード用の追加オプションを取得する"""
    prefix = input("接頭辞（ファイル名の前につける文字）を入力（不要ならEnter）: ")
    suffix = input("接尾辞（拡張子の前につける文字）を入力（不要ならEnter）: ")
    start = int(input("連番の開始番号を入力（デフォルト: 1）: ") or "1")
    digits = int(input("連番の桁数を入力（例: 3 -> 001）（デフォルト: 3）: ") or "3")
    return {"prefix": prefix, "suffix": suffix, "start": start, "digits": digits}

def get_date_options():
    """dateモード用の追加オプションを取得する"""
    prefix = input("接頭辞（ファイル名の前につける文字）を入力（不要ならEnter）: ")
    suffix = input("接尾辞（拡張子の前につける文字）を入力（不要ならEnter）: ")
    date_type = input("日付の種類を選択 ('modified' or 'created')（デフォルト: modified）: ") or "modified"
    date_format = input("日付のフォーマットを入力 (デフォルト: %%Y-%%m-%%d_%%H%%M%%S): ") or "%Y-%m-%d_%H%M%S"
    return {"prefix": prefix, "suffix": suffix, "date_type": date_type, "date_format": date_format}


# --- 3. すべてを動かすメイン関数 ---

def main():
    """ユーザーとの対話を通じてリネーム処理を実行する"""
    try:
        target_path = get_folder_path()
        mode = get_rename_mode()
        
        options = {}
        if mode == 'simple':
            options = get_simple_options()
        elif mode == 'date':
            options = get_date_options()
        else: # 他のモード用の共通オプション
            options['prefix'] = input("接頭辞を入力（不要ならEnter）: ")
            options['suffix'] = input("接尾辞を入力（不要ならEnter）: ")
            if mode == 'ai-sort':
                 options['start'] = int(input("連番の開始番号を入力（デフォルト: 1）: ") or "1")
                 options['digits'] = int(input("連番の桁数を入力（デフォルト: 3）: ") or "3")

        # ドライランの実行
        print("\n--- ドライラン実行結果 ---")
        renamer_dry = RenbanSuper(str(target_path), dry_run=True, **options)
        getattr(renamer_dry, f"run_{mode}")(**options)
        print("--- ドライラン終了 ---")
        
        # 最終確認
        confirm = input("\nこの変更を適用しますか？ (Yes / No) [Y/n]: ").lower().strip()
        if confirm in ['y', 'yes', '']:
            print("\n--- 変更を適用します ---")
            renamer_final = RenbanSuper(str(target_path), dry_run=False, **options)
            getattr(renamer_final, f"run_{mode}")(**options)
            print("--- 処理完了 ---")
        else:
            print("処理を中断しました。")
            
    except KeyboardInterrupt:
        print("\n処理が中断されました。")
    except Exception as e:
        print(f"\n予期せぬエラーが発生しました: {e}")

if __name__ == '__main__':
    main()