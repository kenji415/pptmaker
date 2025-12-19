"""
PowerPointファイル（PPTX）を自動変換するスクリプト

入力PPT: 表紙 → 問題ページ群 → 解答1枚
出力PPT: 表紙 → 問題ページ群 → 解答を埋め込んだ問題ページ群 → 解答1枚
"""

import os
import re
import sys
import shutil
import io
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn


def extract_answers(answer_slide):
    """
    解答ページから大問番号と解答文字列の辞書を生成
    
    Args:
        answer_slide: 解答ページのスライドオブジェクト
        
    Returns:
        dict: {大問番号: 解答文字列} の辞書
        例: {1: "38.608㎠", 2: "52.685㎠", 3: "36", ...}
    """
    # スライド内のすべてのテキストを連結
    all_text = ""
    for shape in answer_slide.shapes:
        if hasattr(shape, "text"):
            all_text += shape.text + "\n"
    
    # 正規表現で解答を抽出
    # パターン: 大問\s*([0-9]+)\s*([^\n\r]+)
    pattern = r"大問\s*([0-9]+)\s*([^\n\r]+)"
    matches = re.findall(pattern, all_text)
    
    answers = {}
    for match in matches:
        question_num = int(match[0])
        answer_text = match[1].strip()
        answers[question_num] = answer_text
    
    return answers


def extract_text_from_shape(shape):
    """
    シェイプからテキストを取得（GROUP内の子シェイプも再帰的に探索）
    
    Args:
        shape: シェイプオブジェクト
        
    Returns:
        str: テキスト（見つからない場合は空文字列）
    """
    text = ""
    try:
        # 直接テキストを取得
        if hasattr(shape, "text") and shape.text:
            text = shape.text.strip()
        elif hasattr(shape, "text_frame") and shape.text_frame:
            text = shape.text_frame.text.strip()
        
        # GROUPシェイプの場合は子シェイプも探索
        if not text and hasattr(shape, "shapes"):
            for child_shape in shape.shapes:
                child_text = extract_text_from_shape(child_shape)
                if child_text:
                    text = child_text
                    break
    except:
        pass
    
    return text


def extract_question_number_candidates(slide, prs, debug=False):
    """
    問題ページから大問番号の候補を抽出
    スライドを左右に分割し、それぞれの領域の左上から大問番号を探す
    GROUPシェイプ内の子シェイプも再帰的に探索
    
    Args:
        slide: 問題ページのスライドオブジェクト
        prs: Presentationオブジェクト（スライドサイズ取得用）
        debug: デバッグ情報を出力するかどうか
        
    Returns:
        list: 大問番号の候補リスト（(優先度, 距離, 大問番号, テキスト, 領域)のタプルのリスト）
        領域は 'left' または 'right'
    """
    # スライドのサイズを取得
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    
    # スライドを左右に分割
    left_half_end = slide_width / 2
    right_half_start = slide_width / 2
    
    # 各領域の左上の範囲を定義（各領域の左上15%以内に絞る）
    area_width = slide_width / 2
    left_threshold = area_width * 0.15
    top_threshold = slide_height * 0.15
    
    candidates = []
    debug_shapes = []  # デバッグ用
    
    def process_shape(shape, parent_left=0, parent_top=0):
        """シェイプを再帰的に処理"""
        try:
            # シェイプの位置を取得（GROUP内の場合は親からの相対位置、直接の場合は絶対位置）
            if hasattr(shape, 'left') and hasattr(shape, 'top'):
                # python-pptxでは、GROUP内の子シェイプの位置は親からの相対位置として取得される
                # 親の位置を加算する必要がある
                shape_left = parent_left + shape.left
                shape_top = parent_top + shape.top
            else:
                shape_left = parent_left
                shape_top = parent_top
            
            # 左半分の領域か右半分の領域かを判定
            area = None
            relative_left = None
            
            if shape_left < left_half_end:
                # 左半分
                area = 'left'
                relative_left = shape_left
            elif shape_left >= right_half_start:
                # 右半分
                area = 'right'
                relative_left = shape_left - right_half_start
            
            if area is None:
                # GROUP内の子シェイプを探索
                if hasattr(shape, "shapes"):
                    for child_shape in shape.shapes:
                        process_shape(child_shape, shape_left, shape_top)
                return
            
            # 各領域の左上15%以内か確認
            if relative_left <= left_threshold and shape_top <= top_threshold:
                # テキストを取得（GROUP内も再帰的に探索）
                text = extract_text_from_shape(shape)
                
                # デバッグ情報を記録（短いテキストのみ）
                if debug and len(text) <= 10:
                    shape_type = getattr(shape, 'shape_type', 'unknown')
                    debug_shapes.append({
                        'area': area,
                        'text': text,
                        'left': shape_left,
                        'top': shape_top,
                        'relative_left': relative_left,
                        'shape_type': shape_type
                    })
                
                # 大問番号は通常、短いテキスト（1-3文字程度）で数字のみ
                # 長いテキスト（問題文など）は除外
                if len(text) > 5:
                    # GROUP内の子シェイプを探索
                    if hasattr(shape, "shapes"):
                        for child_shape in shape.shapes:
                            process_shape(child_shape, shape_left, shape_top)
                    return
                
                # テキストが数字のみの場合のみを対象
                if not text.isdigit():
                    # GROUP内の子シェイプを探索
                    if hasattr(shape, "shapes"):
                        for child_shape in shape.shapes:
                            process_shape(child_shape, shape_left, shape_top)
                    return
                
                # 大問番号を取得
                question_num = int(text)
                
                # 大問番号は通常1-20程度なので、それ以外は除外
                if question_num < 1 or question_num > 20:
                    return
                
                # 位置でソートするため、距離を計算（各領域の左上からの距離）
                distance = (relative_left ** 2 + shape_top ** 2) ** 0.5
                
                # テキストボックスかどうかを判定
                is_textbox = False
                try:
                    if hasattr(shape, 'shape_type'):
                        if shape.shape_type in [1, 17]:  # MSO_SHAPE_TYPE.AUTO_SHAPE または TEXT_BOX
                            is_textbox = True
                except:
                    pass
                
                # 優先度: テキストボックス > 数字のみのテキスト > その他
                priority = 2 if is_textbox else (1 if text.isdigit() else 0)
                
                candidates.append((priority, distance, question_num, text, area))
            else:
                # GROUP内の子シェイプを探索
                if hasattr(shape, "shapes"):
                    for child_shape in shape.shapes:
                        process_shape(child_shape, shape_left, shape_top)
        except Exception as e:
            if debug:
                print(f"    シェイプ処理エラー: {e}")
    
    # すべてのシェイプを処理
    for shape in slide.shapes:
        process_shape(shape)
    
    # デバッグ情報を出力
    if debug and debug_shapes:
        print(f"    デバッグ: 左上領域内の短いテキストシェイプ {len(debug_shapes)}個")
        for ds in debug_shapes[:10]:  # 最初の10個を表示
            print(f"      {ds['area']}: '{ds['text']}' (長さ: {len(ds['text'])}, 位置: {ds['left']:.0f}, {ds['top']:.0f}, タイプ: {ds['shape_type']})")
    
    # 優先度と距離でソート（優先度が高い順、距離が小さい順）
    candidates.sort(key=lambda x: (-x[0], x[1]))
    return candidates


def match_question_numbers(question_slides, prs, answer_numbers):
    """
    各問題ページから抽出した大問番号候補と、解答ページの大問番号を照合
    各スライドには左右2つの問題があり、それぞれに大問番号がある
    
    Args:
        question_slides: 問題ページのスライドリスト
        prs: Presentationオブジェクト
        answer_numbers: 解答ページから抽出した大問番号のセット
        
    Returns:
        list: 各スライドの(左側の大問番号, 右側の大問番号)のタプルのリスト
    """
    # 各問題ページから大問番号の候補を取得（左右両方）
    all_candidates = []
    for idx, slide in enumerate(question_slides):
        candidates = extract_question_number_candidates(slide, prs, debug=True)
        all_candidates.append(candidates)
        if not candidates:
            print(f"  スライド{idx + 1}: 候補が見つかりませんでした")
    
    # 各スライドから左右の大問番号を抽出
    slide_question_numbers = []
    all_question_numbers = []  # 検証用
    
    for idx, candidates in enumerate(all_candidates):
        # 左右の領域からそれぞれ最適な候補を選択
        left_candidates = [c for c in candidates if c[4] == 'left']  # 領域が'left'のもの
        right_candidates = [c for c in candidates if c[4] == 'right']  # 領域が'right'のもの
        
        # 左側の大問番号を選択
        left_num = None
        if left_candidates:
            for priority, distance, question_num, text, area in left_candidates:
                if question_num in answer_numbers:
                    left_num = question_num
                    print(f"  スライド{idx + 1}（左側）: 大問番号 {question_num} を抽出 (テキスト: '{text}')")
                    break
            if left_num is None:
                left_num = left_candidates[0][2]
                print(f"警告: スライド{idx + 1}（左側）の大問番号候補 {left_num} は解答ページに存在しません。")
        else:
            print(f"警告: スライド{idx + 1}（左側）から大問番号の候補が見つかりませんでした。")
        
        # 右側の大問番号を選択
        right_num = None
        if right_candidates:
            for priority, distance, question_num, text, area in right_candidates:
                if question_num in answer_numbers:
                    right_num = question_num
                    print(f"  スライド{idx + 1}（右側）: 大問番号 {question_num} を抽出 (テキスト: '{text}')")
                    break
            if right_num is None:
                right_num = right_candidates[0][2]
                print(f"警告: スライド{idx + 1}（右側）の大問番号候補 {right_num} は解答ページに存在しません。")
        else:
            print(f"警告: スライド{idx + 1}（右側）から大問番号の候補が見つかりませんでした。")
        
        # 左右の大問番号のタプルを追加
        slide_question_numbers.append((left_num, right_num))
        if left_num is not None:
            all_question_numbers.append(left_num)
        if right_num is not None:
            all_question_numbers.append(right_num)
    
    # 検証: 1から順番に並んでいるか、合計が一致するか確認
    sorted_numbers = sorted(all_question_numbers)
    expected_sum = sum(answer_numbers)
    actual_sum = sum(all_question_numbers)
    
    print(f"\n大問番号の検証:")
    print(f"  抽出された大問番号: {all_question_numbers}")
    print(f"  解答ページの大問番号: {sorted(answer_numbers)}")
    print(f"  合計: 抽出={actual_sum}, 解答={expected_sum}")
    
    if sorted_numbers == list(range(1, len(sorted_numbers) + 1)) and len(sorted_numbers) == len(answer_numbers):
        print(f"  ✓ 1から順番に並んでいます")
    else:
        print(f"  警告: 1から順番に並んでいません")
    
    if actual_sum == expected_sum:
        print(f"  ✓ 合計が一致しています")
    else:
        print(f"  警告: 合計が一致しません")
    
    return slide_question_numbers


def duplicate_slide_complete(prs, source_slide):
    """
    スライドを完全複製する（画像やコメントも保持）
    Ctrl+C/Ctrl+Vのようにスライド全体を完全にコピー
    GROUPシェイプ内のすべての要素（テキスト、線、図形、画像など）も含めてコピー
    
    Args:
        prs: Presentationオブジェクト
        source_slide: 複製元のスライド
        
    Returns:
        複製されたスライドオブジェクト
    """
    import io
    
    # 新しいスライドを作成（同じレイアウトを使用）
    dest_slide = prs.slides.add_slide(source_slide.slide_layout)
    
    # デフォルトのプレースホルダーを削除（新規作成時に自動で入るテキストボックスを削除）
    for shape in list(dest_slide.shapes):
        try:
            # プレースホルダーかどうかを確認
            if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                sp = shape.element
                sp.getparent().remove(sp)
            else:
                # プレースホルダーでない場合も削除（レイアウトから来る要素）
                sp = shape.element
                sp.getparent().remove(sp)
        except:
            # 削除できない場合はスキップ
            continue
    
    # 元のスライドのすべてのシェイプをコピー
    for shape in source_slide.shapes:
        try:
            # 画像の場合は特別な処理
            if hasattr(shape, 'image') and shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                # 画像データを取得
                image_stream = io.BytesIO(shape.image.blob)
                # 新しいスライドに画像を追加
                dest_slide.shapes.add_picture(
                    image_stream,
                    shape.left,
                    shape.top,
                    shape.width,
                    shape.height
                )
            else:
                # その他のシェイプ（GROUPシェイプを含む）はXMLをコピー
                # GROUPシェイプのXMLには子シェイプも含まれているため、これでコピーされる
                shape_xml = shape.element
                new_shape_xml = deepcopy(shape_xml)
                
                # GROUPシェイプの場合、XML内の画像リレーションシップのrIdを更新
                if hasattr(shape, 'shape_type') and shape.shape_type == 6:  # GROUP
                    # XML内のすべての画像リレーションシップを探してコピー
                    try:
                        # XML要素内のすべてのrId参照を探す（名前空間を考慮）
                        def find_and_update_rids(element, dest_slide, source_slide):
                            """XML要素内のrId参照を探して更新"""
                            rId_map = {}  # 元のrId -> 新しいrIdのマッピング
                            
                            # ソーススライドのリレーションシップを取得
                            source_rels = source_slide.part.rels
                            
                            # まず、XML内で使用されているすべてのrIdを収集
                            used_rIds = set()
                            ns_embed = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
                            
                            for elem in element.iter():
                                # r:embed属性を確認
                                if ns_embed in elem.attrib:
                                    used_rIds.add(elem.attrib[ns_embed])
                                # その他のrId属性も確認
                                for attr_name, attr_value in list(elem.attrib.items()):
                                    if isinstance(attr_value, str) and attr_value.startswith('rId'):
                                        used_rIds.add(attr_value)
                            
                            # デバッグ: 見つかったrIdを表示
                            if used_rIds:
                                print(f"  デバッグ: GROUPシェイプ内で見つかったrId: {used_rIds}")
                            
                            # 使用されているrIdの画像リレーションシップをコピー
                            for rId in used_rIds:
                                # スライドレベルのリレーションシップを確認
                                if rId in source_rels:
                                    try:
                                        rel = source_rels[rId]
                                        rel_type = rel.reltype
                                        
                                        # 画像のリレーションシップの場合（PNG/JPEG/GIFなど）
                                        if "image" in rel_type or "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" in rel_type:
                                            # 画像データ（blob）を取得
                                            try:
                                                from pptx.parts.image import ImagePart
                                                image_blob = None
                                                
                                                # target_partの型を確認
                                                target_part = rel.target_part
                                                
                                                # ImagePartの場合、blobを取得
                                                if isinstance(target_part, ImagePart):
                                                    image_blob = target_part.blob
                                                else:
                                                    # target_partがImagePartでない場合、_targetから画像パートを取得
                                                    try:
                                                        target_path = rel._target
                                                        # packageから画像パートを取得
                                                        # target_pathが相対パスの場合、package.get_partで取得できる
                                                        image_part = source_slide.part.package.get_part(target_path)
                                                        if hasattr(image_part, 'blob'):
                                                            image_blob = image_part.blob
                                                        elif hasattr(image_part, 'image'):
                                                            # imageプロパティがある場合（一部のパートタイプ）
                                                            image_blob = image_part.image.blob
                                                        else:
                                                            # パートから直接読み込む
                                                            image_blob = image_part._blob if hasattr(image_part, '_blob') else None
                                                    except Exception as e:
                                                        # _targetから取得できない場合、target_partから直接取得を試みる
                                                        if hasattr(target_part, 'blob'):
                                                            image_blob = target_part.blob
                                                        elif hasattr(target_part, 'image'):
                                                            image_blob = target_part.image.blob
                                                        else:
                                                            print(f"  デバッグ: 画像blob取得エラー (rId: {rId}): {e}")
                                                
                                                if image_blob is None:
                                                    print(f"  デバッグ: 警告: 画像blobが取得できませんでした (rId: {rId})")
                                                    continue
                                                
                                                print(f"  デバッグ: 画像リレーションシップ {rId} を処理中 (タイプ: {rel_type}, サイズ: {len(image_blob)} bytes)")
                                                
                                                # 既に同じ画像データのリレーションシップが存在するか確認
                                                # スライドレベルのリレーションシップを確認
                                                already_exists = False
                                                existing_rId = None
                                                for dest_rel in dest_slide.part.rels:
                                                    try:
                                                        if dest_slide.part.rels[dest_rel].reltype == rel_type:
                                                            dest_target_part = dest_slide.part.rels[dest_rel].target_part
                                                            if isinstance(dest_target_part, ImagePart):
                                                                dest_blob = dest_target_part.blob
                                                                if dest_blob == image_blob:
                                                                    already_exists = True
                                                                    existing_rId = dest_rel
                                                                    print(f"  デバッグ: 既存のリレーションシップ {dest_rel} を使用")
                                                                    break
                                                    except:
                                                        continue
                                                
                                                if not already_exists:
                                                    # 画像パートは既にパッケージ内に存在している（shutil.copy2でコピー済み）
                                                    # 元のスライドのリレーションシップから画像パートのパスを取得して、パッケージから直接取得
                                                    try:
                                                        # 元のスライドの画像パートのパスを取得
                                                        source_image_path = rel._target
                                                        
                                                        # パッケージから同じパスで画像パートを取得
                                                        # （パッケージはコピーされているので、同じパスで画像パートが存在するはず）
                                                        found_image_part = None
                                                        try:
                                                            found_image_part = dest_slide.part.package.get_part(source_image_path)
                                                            print(f"  デバッグ: パッケージ内で既存の画像パートを発見 (パス: {source_image_path})")
                                                        except Exception as path_error:
                                                            # パスで取得できない場合、get_or_add_image_partを使用
                                                            print(f"  デバッグ: パスで取得できなかったため、get_or_add_image_partを使用: {path_error}")
                                                            image_stream = io.BytesIO(image_blob)
                                                            image_stream.seek(0)
                                                            found_image_part = dest_slide.part.package.get_or_add_image_part(image_stream)
                                                        
                                                        # 新しいリレーションシップを作成（スライドレベル）
                                                        # relate_toメソッドを使用
                                                        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
                                                        new_rId = dest_slide.part.relate_to(found_image_part, RT.IMAGE)
                                                        rId_map[rId] = new_rId
                                                        print(f"  デバッグ: 新しいリレーションシップ {new_rId} を作成 (元のrId: {rId})")
                                                    except Exception as e:
                                                        print(f"  デバッグ: 画像パート追加エラー (rId: {rId}): {e}")
                                                        print(f"  デバッグ: エラーの型: {type(e)}")
                                                        import traceback
                                                        print(f"  デバッグ: トレースバック:\n{traceback.format_exc()}")
                                                        # エラーが発生しても処理を続行
                                                        continue
                                                else:
                                                    rId_map[rId] = existing_rId
                                                    print(f"  デバッグ: rIdマッピング {rId} -> {existing_rId}")
                                            except Exception as e:
                                                # blob取得に失敗した場合はスキップ
                                                print(f"  デバッグ: 画像blob取得エラー (rId: {rId}): {e}")
                                                pass
                                    except Exception as e:
                                        print(f"  デバッグ: リレーションシップ処理エラー (rId: {rId}): {e}")
                                        pass
                            
                            # XML内のrId参照を更新
                            updated_count = 0
                            for elem in element.iter():
                                # r:embed属性を更新
                                if ns_embed in elem.attrib:
                                    old_rId = elem.attrib[ns_embed]
                                    if old_rId in rId_map:
                                        elem.attrib[ns_embed] = rId_map[old_rId]
                                        updated_count += 1
                                        print(f"  デバッグ: r:embed属性を更新 {old_rId} -> {rId_map[old_rId]}")
                                
                                # その他のrId属性も更新
                                for attr_name, attr_value in list(elem.attrib.items()):
                                    if isinstance(attr_value, str) and attr_value.startswith('rId'):
                                        if attr_value in rId_map:
                                            elem.attrib[attr_name] = rId_map[attr_value]
                                            updated_count += 1
                                            print(f"  デバッグ: {attr_name}属性を更新 {attr_value} -> {rId_map[attr_value]}")
                            
                            if updated_count == 0 and used_rIds:
                                print(f"  デバッグ: 警告: rIdマッピングが作成されませんでした (使用されたrId: {used_rIds})")
                            elif updated_count > 0:
                                print(f"  デバッグ: {updated_count}個のrId参照を更新しました")
                        
                        find_and_update_rids(new_shape_xml, dest_slide, source_slide)
                    except Exception as e:
                        # XML更新に失敗した場合は元のXMLを使用
                        pass
                
                dest_slide.shapes._spTree.append(new_shape_xml)
        except Exception as e:
            # コピーできないシェイプはスキップ
            continue
    
    # リレーションシップをコピー（画像などのリンクを保持）
    # 注意: リレーションシップのコピーは複雑なため、必要に応じて後で処理
    # XMLをコピーした時点で、多くのリレーションシップは既に参照されている
    try:
        for rel in source_slide.part.rels:
            if "notesSlide" not in source_slide.part.rels[rel].reltype:
                try:
                    # 既存のリレーションシップを確認
                    rel_target = source_slide.part.rels[rel]._target
                    rel_type = source_slide.part.rels[rel].reltype
                    
                    # 既に同じターゲットのリレーションシップが存在するか確認
                    already_exists = False
                    for dest_rel in dest_slide.part.rels:
                        if dest_slide.part.rels[dest_rel]._target == rel_target:
                            already_exists = True
                            break
                    
                    if not already_exists:
                        # 新しいリレーションシップを追加
                        dest_slide.part.rels.add_relationship(
                            rel_type,
                            rel_target,
                            rel.rId
                        )
                except:
                    # リレーションシップの追加に失敗した場合はスキップ
                    # （既に存在する場合や、rIdの重複など）
                    pass
    except:
        # リレーションシップのコピーに失敗しても続行
        pass
    
    return dest_slide


def add_answer_textbox(slide, answer_text, prs, position='right'):
    """
    スライドの下部に解答テキストボックスを追加（左右どちらかの領域）
    
    Args:
        slide: 対象のスライドオブジェクト
        answer_text: 解答文字列（例: "38.608㎠"）
        prs: Presentationオブジェクト（スライドサイズ取得用）
        position: 'left' または 'right'（左右どちらの領域に配置するか）
    """
    # スライドのサイズを取得（プレゼンテーションから取得）
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    
    # テキストボックスのサイズと位置を設定
    textbox_width = Inches(3)
    textbox_height = Inches(0.6)
    margin = Inches(0.5)  # 端からのマージン
    top = slide_height * 0.82  # 下部（82%の位置）
    
    # 左右どちらの領域に配置するかで位置を決定
    if position == 'left':
        # 左側の領域の右端に配置
        area_width = slide_width / 2
        left = area_width - textbox_width - margin
    else:  # position == 'right'
        # 右側の領域の右端に配置
        area_start = slide_width / 2
        left = slide_width - textbox_width - margin
    
    # テキストボックスを作成
    textbox = slide.shapes.add_textbox(left, top, textbox_width, textbox_height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = False
    
    # テキストを設定（【解答】を削除）
    text_frame.text = answer_text
    
    # 段落の設定（右寄せ）
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.RIGHT
    
    # フォント設定（游明朝）
    font = paragraph.font
    font.size = Pt(9)  # 9pt
    font.bold = False  # 太字なし
    font.color.rgb = RGBColor(0, 0, 0)  # 黒色
    
    # 游明朝フォントを設定
    try:
        font.name = "游明朝"
    except:
        try:
            font.name = "Yu Mincho"
        except:
            # 游明朝が利用できない場合はデフォルトフォントを使用
            pass


def convert_pptx(input_path, output_path=None):
    """
    PPTXファイルを変換する
    
    Args:
        input_path: 入力PPTXファイルのパス
        output_path: 出力PPTXファイルのパス（Noneの場合は自動生成）
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"入力ファイルが見つかりません: {input_path}")
    
    if not input_path.lower().endswith('.pptx'):
        raise ValueError("PPTXファイルを指定してください")
    
    # 出力パスが指定されていない場合は自動生成
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_converted{ext}"
    
    # PPTXを読み込む
    prs = Presentation(input_path)
    
    if len(prs.slides) < 2:
        raise ValueError("スライドが2枚未満です。表紙と解答ページが必要です。")
    
    # スライド構成を判定
    # 0枚目 → 表紙
    # 最後 → 解答ページ
    # それ以外 → 問題ページ群
    # 注意: prs.slidesはスライスをサポートしていないため、リストに変換する
    slides_list = list(prs.slides)
    cover_slide = slides_list[0]
    answer_slide = slides_list[-1]
    question_slides = slides_list[1:-1]
    
    if len(question_slides) == 0:
        raise ValueError("問題ページが見つかりません。")
    
    print(f"スライド構成:")
    print(f"  表紙: 1枚")
    print(f"  問題ページ: {len(question_slides)}枚")
    print(f"  解答ページ: 1枚")
    
    # 解答ページから解答を抽出
    answers = extract_answers(answer_slide)
    print(f"\n抽出された解答: {len(answers)}件")
    for q_num, answer in sorted(answers.items()):
        print(f"  大問 {q_num}: {answer}")
    
    # 元のファイルをコピーしてから操作（画像やコメントも保持）
    # 出力ファイルとしてコピーを作成
    shutil.copy2(input_path, output_path)
    
    # コピーしたファイルを読み込む
    new_prs = Presentation(output_path)
    
    # スライドリストを取得
    new_slides_list = list(new_prs.slides)
    
    # 各問題ページから大問番号を抽出して照合
    answer_numbers = set(answers.keys())
    slide_question_numbers = match_question_numbers(question_slides, new_prs, answer_numbers)
    
    # 問題ページを複製して解答入りページを作成
    # 問題ページの後に挿入するため、前から処理
    new_question_slides_with_answers = []
    
    for idx in range(len(question_slides)):  # 0から始まるインデックス
        # 元のスライドのインデックスを取得（new_prs内での位置）
        # 0: 表紙, 1～len(question_slides): 問題ページ, 最後: 解答ページ
        source_slide_index = idx + 1  # 問題ページのインデックス（1から始まる）
        
        # 元のスライドを取得
        source_slide = new_slides_list[source_slide_index]
        
        # 照合された大問番号を取得（左右両方）
        left_num, right_num = slide_question_numbers[idx]
        
        # スライドを完全複製
        new_slide = duplicate_slide_complete(new_prs, source_slide)
        
        # 左側の解答を追加
        if left_num is not None:
            answer_text = answers.get(left_num, "（未設定）")
            if answer_text == "（未設定）":
                print(f"警告: 大問{left_num}の解答が見つかりませんでした。")
            else:
                add_answer_textbox(new_slide, answer_text, new_prs, position='left')
        
        # 右側の解答を追加
        if right_num is not None:
            answer_text = answers.get(right_num, "（未設定）")
            if answer_text == "（未設定）":
                print(f"警告: 大問{right_num}の解答が見つかりませんでした。")
            else:
                add_answer_textbox(new_slide, answer_text, new_prs, position='right')
        
        new_question_slides_with_answers.append(new_slide)
    
    # スライドの順序を調整: 表紙、問題ページ群、解答入り問題ページ群、解答ページ
    # python-pptxでは直接順序を変更できないため、XMLレベルで順序を調整
    # スライドIDリスト（_sldIdLst）の順序を変更
    sldIdLst = new_prs.slides._sldIdLst
    
    # 現在のスライドID要素を取得
    current_slide_ids = list(sldIdLst)
    
    # 順序: 表紙(0), 問題ページ(1～len), 解答入り問題ページ(追加分), 解答ページ(最後)
    # 新しい順序でスライドID要素を再配置
    new_order = []
    
    # 1. 表紙（最初のスライド）
    new_order.append(current_slide_ids[0])
    
    # 2. 問題ページ群（1～len(question_slides)）
    for i in range(1, len(question_slides) + 1):
        new_order.append(current_slide_ids[i])
    
    # 3. 解答入り問題ページ群（最後に追加されたスライド）
    # 追加されたスライドは最後に配置されるため、問題ページの後から解答ページの前まで
    for i in range(len(question_slides) + 1, len(current_slide_ids) - 1):
        new_order.append(current_slide_ids[i])
    
    # 4. 解答ページ（最後のスライド）
    new_order.append(current_slide_ids[-1])
    
    # スライドIDリストをクリアして再構築
    sldIdLst.clear()
    for slide_id_elem in new_order:
        sldIdLst.append(slide_id_elem)
    
    # 新しいPPTXを保存
    new_prs.save(output_path)
    print(f"\n変換完了: {output_path}")
    print(f"  総スライド数: {len(new_prs.slides)}枚")
    print(f"  (表紙: 1枚, 問題ページ: {len(question_slides)}枚, 解答入り問題ページ: {len(new_question_slides_with_answers)}枚, 解答ページ: 1枚)")


def main():
    """コマンドライン実行用のメイン関数"""
    if len(sys.argv) < 2:
        print("使用方法: python convert_pptx.py <入力PPTXファイル> [出力PPTXファイル]")
        print("例: python convert_pptx.py input.pptx output.pptx")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        convert_pptx(input_path, output_path)
    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()


