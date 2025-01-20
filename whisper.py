from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from faster_whisper import WhisperModel
from pydub import AudioSegment
from pydub.silence import detect_silence
from pykakasi import kakasi
import os
import xlsxwriter
import re
import ffmpeg  # 追加

app = Flask(__name__)
CORS(app)

# フォルダ設定
UPLOAD_FOLDER = 'uploads'
TRIMMED_FOLDER = 'static/trimmed'
RESULT_FOLDER = 'static/results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TRIMMED_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# Whisperモデルのロード
model = WhisperModel("tiny", device="cpu", compute_type="float32")

# 正解の単語リスト
correct_words_list = [
    ["みかん", "さかな", "とけい", "ぼうし", "めがね", "きりん", "でんわ", "あたま", "くるま", "りんご"],
    ["くえせ", "まごな", "らめす", "きおつ", "びれく", "まぬら", "ぼのず", "けたり", "さゆぎ", "さおか"],
    ["せんせい", "さんすう", "くつした", "ともだち", "にんじん", "くだもの", "にわとり", "てぶくろ", "ちいさい", "なわとび"],
    ["くのだり", "いぎしま", "るみかの", "こんうさ", "なびわく", "いかもの", "たずりね", "あたらび", "むきけし", "かなみに"],
    ["すべりだい", "かぶとむし", "ありがとう", "ようちえん", "おもしろい", "ゆでたまご", "おとしだま", "こんにちは", "かたつむり", "おかあさん"],
    ["きだゆなる", "しんぶりん", "からたこも", "すわれのそ", "とめうげお", "くろぶすき", "かちもちま", "うきびんと", "もろもうか", "なよさうや"]
]

# 漢字やカタカナをひらがなに変換する関数
def to_hiragana(text):
    kakasi_converter = kakasi()
    kakasi_converter.setMode("H", "H")  # ひらがなをひらがなに変換
    kakasi_converter.setMode("K", "H")  # カタカナをひらがなに変換
    kakasi_converter.setMode("J", "H")  # 漢字をひらがなに変換
    converter = kakasi_converter.getConverter()
    return converter.do(text)

# 句読点や空白を削除する関数
def clean_text(text):
    text = to_hiragana(text)
    return re.sub(r'[。、\s]', '', text)

# 音声文字起こし関数
def transcribe_audio(file_path):
    segments, _ = model.transcribe(file_path, beam_size=5, language='ja')
    transcription = ''.join([segment.text for segment in segments])
    return transcription

# 比較結果を生成する関数
def compare_transcription(transcription, correct_words):
    cleaned_transcription = clean_text(transcription)
    results = [f"{word}: {'○' if word in cleaned_transcription else '×'}" for word in correct_words]
    return results

# WebMをWAVに変換する関数
def convert_webm_to_wav(input_path, output_path):
    ffmpeg.input(input_path).output(output_path).run()

# 音声ファイルを処理する関数
def process_audio_file(audio_file, correct_words):
    file_path = os.path.join(UPLOAD_FOLDER, audio_file.filename)
    audio_file.save(file_path)

    # WebMファイルをWAVに変換
    if file_path.endswith('.webm'):
        wav_path = file_path.replace('.webm', '.wav')
        convert_webm_to_wav(file_path, wav_path)
        file_path = wav_path

    # ファイルの拡張子に応じた読み込み
    file_extension = os.path.splitext(file_path)[1].lower()
    audio = AudioSegment.from_file(file_path, format=file_extension.strip('.'))

    # 音声のトリミング
    silence_thresh = -45
    min_silence_len = 200
    start_silence = detect_silence(audio, min_silence_len=min_silence_len, silence_thresh=silence_thresh)
    start_trim = start_silence[0][1] if start_silence and start_silence[0][0] == 0 else 0
    end_silence = detect_silence(audio.reverse(), min_silence_len=min_silence_len, silence_thresh=silence_thresh)
    end_trim = end_silence[0][1] if end_silence and end_silence[0][0] == 0 else 0

    trimmed_audio = audio[start_trim:len(audio) - end_trim]
    trimmed_file_path = os.path.join(TRIMMED_FOLDER, audio_file.filename)
    trimmed_audio.export(trimmed_file_path, format="wav")

    # 文字起こし
    transcription = transcribe_audio(trimmed_file_path)
    cleaned_transcription = clean_text(transcription)
    comparison_results = compare_transcription(transcription, correct_words)

    return {
        "transcription": transcription,
        "cleaned_transcription": cleaned_transcription,
        "comparison_results": comparison_results,
        "trimmed_duration": len(trimmed_audio) / 1000,
        "trimmed_audio_file": os.path.basename(trimmed_file_path)
    }

# 結果をExcelに保存
def save_results_to_excel(results_list):
    file_path = os.path.join(RESULT_FOLDER, "results.xlsx")
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()

    # 書式設定
    merge_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    bold_format = workbook.add_format({'bold': True, 'border': 1})

    # ヘッダー設定
    worksheet.merge_range('A1:N1', '読み取りチェック採点表', merge_format)
    worksheet.merge_range('A5:A6', '課題１', merge_format)
    worksheet.merge_range('A7:A8', '課題２', merge_format)
    worksheet.merge_range('A9:A10', '課題３', merge_format)
    worksheet.merge_range('A11:A12', '課題４', merge_format)
    worksheet.merge_range('A13:A16', '課題５', merge_format)
    worksheet.merge_range('A17:A20', '課題６', merge_format)

    worksheet.write('L4', '〇数', bold_format)
    worksheet.write('M4', '計', bold_format)
    worksheet.write('N4', '時間（秒）', bold_format)
    worksheet.write('O4', '計', bold_format)
    worksheet.write('K21', '合計', bold_format)

    # 課題ごとの単語リスト
    words_list = [
        ["みかん", "さかな", "とけい", "ぼうし", "めがね", "きりん", "でんわ", "あたま", "くるま", "りんご"],
        ["くえせ", "まごな", "らめす", "きおつ", "びれく", "まぬら", "ぼのず", "けたり", "さゆぎ", "さおか"],
        ["せんせい", "さんすう", "くつした", "ともだち", "にんじん", "くだもの", "にわとり", "てぶくろ", "ちいさい", "なわとび"],
        ["くのだり", "いぎしま", "るみかの", "こんうさ", "なびわく", "いかもの", "たずりね", "あたらび", "むきけし", "かなみに"],
        ["すべりだい", "かぶとむし", "ありがとう", "ようちえん", "おもしろい", "ゆでたまご", "おとしだま", "こんにちは", "かたつむり", "おかあさん"],
        ["きだゆなる", "しんぶりん", "からたこも", "すわれのそ", "とめうげお", "くろぶすき", "かちもちま", "うきびんと", "もろもうか", "なよさうや"]
    ]

    # 課題ごとの行範囲
    row_ranges = [
        (5, 6), (7, 8), (9, 10), (11, 12), (13, 16), (17, 20)
    ]

    total_correct = 0
    total_time = 0.0

    for i, (result, words) in enumerate(zip(results_list, words_list)):
        start_row, end_row = row_ranges[i]

        comparison_marks = [res[-1] for res in result["comparison_results"]]

        # 1行ごとに2列ずつ結合して単語を表示
        if i == 4 or i == 5:  # 5と6番目の課題の処理
            # 「すべりだい」～「おもしろい」は13行目と14行目
            for j, word in enumerate(words[:5]):  # 最初の5つの単語
                col_start = chr(66 + 2 * j)  # B, D, F, H, J列
                col_end = chr(67 + 2 * j)    # C, E, G, I, K列
                worksheet.merge_range(f'{col_start}{start_row}:{col_end}{start_row}', word, merge_format)

            # 「ゆでたまご」～「おかあさん」は15行目
            for j, word in enumerate(words[5:]):  # 残りの5つの単語
                col_start = chr(66 + 2 * j)  # B, D, F, H, J列
                col_end = chr(67 + 2 * j)    # C, E, G, I, K列
                worksheet.merge_range(f'{col_start}{start_row + 2}:{col_end}{start_row + 2}', word, merge_format)

            # 課題5の比較結果を記入
            for j, mark in enumerate(comparison_marks[:5]):  # 最初の5つのマーク
                col_start = chr(66 + 2 * j)  # B, D, F, H, J列
                col_end = chr(67 + 2 * j)    # C, E, G, I, K列
                worksheet.merge_range(f'{col_start}{start_row + 1}:{col_end}{start_row + 1}', mark, merge_format)

            for j, mark in enumerate(comparison_marks[5:]):  # 残りの5つのマーク
                col_start = chr(66 + 2 * j)  # B, D, F, H, J列
                col_end = chr(67 + 2 * j)    # C, E, G, I, K列
                worksheet.merge_range(f'{col_start}{end_row}:{col_end}{end_row}', mark, merge_format)

        else:
            # それ以外の課題は元の方法で書き込み
            worksheet.write_row(f'B{start_row}', words, bold_format)
            worksheet.write_row(f'B{end_row}', comparison_marks[:10], bold_format)

        correct_count = sum(1 for mark in result["comparison_results"] if '○' in mark)
        worksheet.merge_range(f'L{start_row}:L{end_row}', correct_count, merge_format)
        total_correct += correct_count

        duration = result["trimmed_duration"]
        worksheet.merge_range(f'N{start_row}:N{end_row}', duration, merge_format)
        total_time += duration

    worksheet.merge_range(f'M6:M7', f'=L5+L7', merge_format)  # 3文字○合計
    worksheet.merge_range(f'M10:M11', f'=L9+L11', merge_format)  # 4文字○合計
    worksheet.merge_range(f'M15:M18', f'=L13+L17', merge_format)  # 5文字○合計
    worksheet.merge_range(f'O6:O7', f'=N5+N7', merge_format)  # 3文字時間合計
    worksheet.merge_range(f'O10:O11', f'=N9+N11', merge_format)  # 4文字時間合計
    worksheet.merge_range(f'O15:O18', f'=N13+N17', merge_format)  # 5文字時間合計
    worksheet.write(f'L21', f'=L5+L7+L9+L11+L13+L17', bold_format)
    worksheet.write(f'M21', f'=M6+M10+M15', bold_format)
    worksheet.write(f'N21', f'=N5+N7+N9+N11+N13+N17', bold_format)
    worksheet.write(f'O21', f'=O6+O10+O15', bold_format)

    workbook.close()
    return os.path.basename(file_path)

# APIエンドポイント
@app.route("/transcribe_six_files", methods=["POST"])
def transcribe_six_files():
    files = [request.files.get(f"file{i}") for i in range(1, 7)]
    if not all(files):
        return jsonify({"error": "6つのファイルをアップロードしてください"}), 400

    try:
        results_list = [process_audio_file(file, correct_words) for file, correct_words in zip(files, correct_words_list)]
        excel_file = save_results_to_excel(results_list)

        return jsonify({
            "results": results_list,
            "excel_file": excel_file
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/static/results/<filename>')
def download_excel(filename):
    return send_from_directory(RESULT_FOLDER, filename)

if __name__ == "__main__":
    app.run(debug=True)


# 正解単語リストとの比較の処理を変えようと思います。

# 正規化したテキスト(ex:くえすまごならめすきおつびびれくまぬらぼのずけたりさよぎさおか)と正解の単語リストで一致するものは<>で挟みます(ex:くえす<まごな><らめす><きおつ>び<びれく><まぬら><ぼのず><けたり>さよぎ<さおか>)。そして、<の左側に>がない単語(ex:す<まごな>、び<びれく>)は△と判定します。そして、<>で囲まれなかった単語は×と判定します。