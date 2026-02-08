"""
PPTファイルの一括PDF化スクリプト
- 指定パス内および下位フォルダのPPTファイルを再帰的に検索
- デスクトップの透かし.pngをスライドマスターの一枚目に追加
- ファイル名を基にPDF化し、各PPTが含まれるフォルダに保存
"""

import os
import sys
from pathlib import Path
import win32com.client
from win32com.client import constants

def get_desktop_path():
    """デスクトップのパスを取得"""
    return os.path.join(os.path.expanduser("~"), "Desktop")

def find_ppt_files(target_path):
    """指定パス内および下位フォルダのPPT/PPTXファイルを再帰的に検索"""
    target = Path(target_path)
    ppt_files = set()
    for ext in ["*.ppt", "*.pptx", "*.PPT", "*.PPTX"]:
        ppt_files.update(target.rglob(ext))
    return sorted([str(p) for p in ppt_files])

def get_ppt_title(presentation, ppt_path):
    """PPTファイルのタイトルを取得"""
    # 方法1: ドキュメントプロパティからタイトルを取得
    try:
        title = presentation.BuiltInDocumentProperties("Title").Value
        if title and str(title).strip():
            return str(title).strip()
    except:
        pass
    
    # 方法2: カスタムプロパティからタイトルを取得
    try:
        custom_props = presentation.BuiltInDocumentProperties
        for i in range(1, custom_props.Count + 1):
            try:
                prop = custom_props.Item(i)
                if prop.Name == "Title" and prop.Value:
                    title = str(prop.Value).strip()
                    if title:
                        return title
            except:
                continue
    except:
        pass
    
    # 方法3: ファイル名（拡張子なし）を使用（必ず返す）
    base_name = os.path.splitext(os.path.basename(ppt_path))[0]
    return base_name

# 透かしの縮小率（0.7 = 70%）
WATERMARK_SCALE = 0.7

def add_watermark_to_slide_master(presentation, watermark_path):
    """スライドマスターの一枚目に透かしを追加（70%程度に縮小して中央に配置）"""
    try:
        # スライドマスターを取得
        slide_master = presentation.SlideMaster
        
        # スライドのサイズを取得
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        
        # 透かしを70%に縮小し、中央に配置
        width = slide_width * WATERMARK_SCALE
        height = slide_height * WATERMARK_SCALE
        left = (slide_width - width) / 2
        top = (slide_height - height) / 2
        
        # スライドマスターのShapesコレクションに画像を追加
        shape = slide_master.Shapes.AddPicture(
            FileName=watermark_path,
            LinkToFile=False,  # 埋め込み
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=width,
            Height=height
        )
        
        # 画像を最背面に移動（msoSendToBack = 4）
        try:
            # 定数が利用可能な場合は使用
            shape.ZOrder(constants.msoSendToBack)
        except (AttributeError, TypeError):
            # 定数が利用できない場合は数値を使用（msoSendToBack = 4）
            try:
                shape.ZOrder(4)
            except:
                # ZOrderが失敗しても画像は追加されているので続行
                pass
        
        return True
    except Exception as e:
        print(f"  エラー: スライドマスターへの透かし追加に失敗しました: {e}")
        import traceback
        traceback.print_exc()
        return False

def convert_ppt_to_pdf(ppt_path, pdf_path, watermark_path=None, powerpoint=None):
    """PPTファイルをPDFに変換"""
    # PDFパスを先に決定し、既に同名PDFがある場合は何もせずスキップ
    # （PPTを開く・透かしを入れるなどの重い処理の前に判断）
    title = os.path.splitext(os.path.basename(ppt_path))[0]
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in invalid_chars:
        title = title.replace(char, '_')
    pdf_dir = os.path.dirname(ppt_path)
    pdf_filename = f"{title}.pdf"
    final_pdf_path = os.path.join(pdf_dir, pdf_filename)
    
    if os.path.exists(final_pdf_path):
        print(f"  同名のPDFが既に存在するためスキップします: {pdf_filename}")
        return final_pdf_path
    
    presentation = None
    try:
        # PowerPointアプリケーションが渡されていない場合は新規作成
        if powerpoint is None:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # Visibleプロパティの設定を試みる（エラーが発生しても続行）
            try:
                powerpoint.Visible = False  # バックグラウンドで実行
            except:
                pass  # 非表示にできない場合は表示のまま続行
            close_powerpoint = True
        else:
            close_powerpoint = False
        
        # PPTファイルを開く
        presentation = powerpoint.Presentations.Open(ppt_path, ReadOnly=False, WithWindow=False)
        
        # 透かしを追加
        if watermark_path and os.path.exists(watermark_path):
            print(f"  透かしを追加中...")
            if not add_watermark_to_slide_master(presentation, watermark_path):
                print(f"  警告: 透かしの追加に失敗しましたが、処理を続行します")
        
        # PDFとして保存
        print(f"  PDF化中: {pdf_filename}")
        presentation.SaveAs(final_pdf_path, FileFormat=32)  # 32 = ppSaveAsPDF
        
        print(f"  ✓ 完了: {pdf_filename}")
        return final_pdf_path
        
    except Exception as e:
        print(f"  ✗ エラー: {e}")
        import traceback
        traceback.print_exc()
        return None
        
    finally:
        # プレゼンテーションを閉じる
        if presentation:
            try:
                presentation.Close()
            except:
                pass
        
        # PowerPointアプリケーションを終了（この関数内で作成した場合のみ）
        if close_powerpoint and powerpoint:
            try:
                powerpoint.Quit()
            except:
                pass

def main():
    """メイン処理"""
    print("=" * 60)
    print("PPTファイル一括PDF化ツール")
    print("=" * 60)
    
    # 指定パスの入力
    if len(sys.argv) > 1:
        target_path = sys.argv[1]
    else:
        target_path = input("PPTファイルが入っているフォルダのパスを入力してください: ").strip().strip('"')
    
    # パスの検証
    if not os.path.exists(target_path):
        print(f"エラー: 指定されたパスが存在しません: {target_path}")
        return
    
    if not os.path.isdir(target_path):
        print(f"エラー: 指定されたパスはフォルダではありません: {target_path}")
        return
    
    # 透かし画像のパス
    desktop_path = get_desktop_path()
    watermark_path = os.path.join(desktop_path, "透かし.png")
    
    if not os.path.exists(watermark_path):
        print(f"警告: デスクトップに「透かし.png」が見つかりません: {watermark_path}")
        response = input("透かしなしで続行しますか？ (y/n): ").strip().lower()
        if response != 'y':
            print("処理を中止しました")
            return
        watermark_path = None
    
    # PPTファイルを検索
    print(f"\n指定パス: {target_path}")
    print(f"透かし画像: {watermark_path if watermark_path else 'なし'}")
    print("\nPPTファイルを検索中...")
    
    ppt_files = find_ppt_files(target_path)
    
    if not ppt_files:
        print("PPTファイルが見つかりませんでした")
        return
    
    print(f"\n{len(ppt_files)}個のPPTファイルが見つかりました:\n")
    for i, ppt_file in enumerate(ppt_files, 1):
        print(f"  {i}. {os.path.basename(ppt_file)}")
    
    # 確認
    print(f"\n上記のファイルをPDF化しますか？")
    response = input("続行しますか？ (y/n): ").strip().lower()
    if response != 'y':
        print("処理を中止しました")
        return
    
    # 各ファイルを処理
    print("\n" + "=" * 60)
    print("処理開始")
    print("=" * 60)
    
    success_count = 0
    error_count = 0
    
    # PowerPointアプリケーションを一度だけ起動
    powerpoint = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # Visibleプロパティの設定を試みる（エラーが発生しても続行）
        try:
            powerpoint.Visible = False  # バックグラウンドで実行
        except:
            pass  # 非表示にできない場合は表示のまま続行
        
        for i, ppt_file in enumerate(ppt_files, 1):
            print(f"\n[{i}/{len(ppt_files)}] {os.path.basename(ppt_file)}")
            
            result = convert_ppt_to_pdf(ppt_file, None, watermark_path, powerpoint)
            
            if result:
                success_count += 1
            else:
                error_count += 1
    finally:
        # PowerPointアプリケーションを終了
        if powerpoint:
            try:
                powerpoint.Quit()
            except:
                pass
    
    # 結果表示
    print("\n" + "=" * 60)
    print("処理完了")
    print("=" * 60)
    print(f"成功: {success_count}件")
    print(f"失敗: {error_count}件")
    print(f"\n出力先: {target_path}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n処理が中断されました")
    except Exception as e:
        print(f"\n予期しないエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
