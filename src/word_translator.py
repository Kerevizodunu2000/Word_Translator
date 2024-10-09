print("Kütüphaneler import ediliyor program birazdan başlayacaktır...")
import openai
from docx import Document
from deep_translator import GoogleTranslator
import pandas as pd
import time
import re
import os
import inquirer
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import excel_formatter
import time
import pygetwindow as gw

def translate_with_chatgpt(text, source_lang="en", target_lang="tr", api_key=None):
    try:
        # OpenAI API anahtarını kullan
        openai.api_key = api_key
        # Yeni OpenAI API kullanımını kullanın
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": f"Translate the following text from {source_lang} to {target_lang}."},
                {"role": "user", "content": text}
            ]
        )
        return response.choices[0].message['content'].strip()
    except Exception as e:
        print(f"ChatGPT çevirisinde hata: {str(e)}")
        return None

def translate_with_google(text, source_lang="en", target_lang="tr"):
    try:
        # Deep Translator kullanarak metni çevir
        translated_text = GoogleTranslator(source=source_lang, target=target_lang).translate(text)
        return translated_text
    except Exception as e:
        print(f"Google çevirisinde hata: {str(e)}")
        return None

def extract_text_from_docx(file_path):
    # Dosyayı aç
    doc = Document(file_path)
    text_list = []

    # Tüm paragrafları listeye ekle
    for para in doc.paragraphs:
        if para.text.strip():
            text_list.append(para.text.strip())
    
    return text_list

def split_paragraphs_to_sentences(paragraphs):
    # Cümlelere ayırma işlemi
    sentences = []
    for para in paragraphs:
        # Satır sonundaki "-" işaretlerini kaldırıp cümleleri birleştir
        para = para.replace("-\n", "").replace("- ", "")
        
        # Paragrafı "." ile ayırarak cümlelere böl
        # İstisna durumlarını ayıklama (Sayısal ifadeler, kısaltmalar, baş harfler)
        para_sentences = re.split(r'(?<!\b\w)\.(?=\s+[^a-z])', para)
        for sentence in para_sentences:
            sentence = sentence.strip()
            if sentence:
                sentences.append(sentence)  # Nokta ekleyerek cümleyi tamamla
    return sentences

def get_unique_sentence_index(sentences, text_type, file_path):
    # Kullanıcıdan cümle girişi alıp metin içindeki indeksleri bul
    while True:
        open_docx(file_path)
        user_input = input(f"Çevirilecek {text_type} cümlesini girin (tam olarak yazın): ").strip()
        indices = [i for i, s in enumerate(sentences) if user_input in s]
        
        if len(indices) > 1:
            print(f"{text_type.capitalize()} cümlesi birden fazla yerde bulundu ({len(indices)} kez). Daha spesifik bir cümle girin.")
        elif len(indices) == 0:
            print(f"{text_type.capitalize()} cümlesi bulunamadı. Lütfen tam olarak ve doğru bir cümle girin.")
        else:
            return indices[0]  # Eğer cümle sadece bir yerde bulunuyorsa doğru indeksi döndür

def translate_and_save_to_excel(file_path, output_excel, use_chatgpt=False, api_key=None):
    # Zaman ölçümünü başlat
    start_time = time.time()
    
    # Dosyadan metni çek
    paragraphs = extract_text_from_docx(file_path)
    
    # Cümlelere ayırma
    sentences = split_paragraphs_to_sentences(paragraphs)

    # Başlangıç ve bitiş cümlelerinin pozisyonunu bul
    start_index = get_unique_sentence_index(sentences, "başlangıç", file_path)
    end_index = get_unique_sentence_index(sentences, "bitiş", file_path)
    close_docx(file_path)


    # Seçilen aralıktaki cümleleri al
    sentences_to_translate = sentences[start_index:end_index + 1]

    # Çevirilecek satır sayısını yazdır
    num_sentences = len(sentences_to_translate)
    print(f"{num_sentences} adet cümle çevirilecek.")

    # Çevirileri saklamak için bir liste oluştur
    data = []

    # İlerleme yüzdesi takibi
    last_printed_progress = -1

    # Çevirileri yap ve ilerlemeyi konsola yazdır
    for index, sentence in enumerate(sentences_to_translate):
        try:
            chatgpt_translation = None
            
            # ChatGPT çevirisini kullanma isteğini kontrol et
            if use_chatgpt and api_key:
                chatgpt_translation = translate_with_chatgpt(sentence, api_key=api_key)
            
            # Google Translate çevirisi
            google_translation = translate_with_google(sentence)
            
            # Sonuçları kaydet
            if use_chatgpt:
                data.append({
                    "İngilizce Cümle": sentence,
                    "ChatGPT Çevirisi": chatgpt_translation,
                    "Google Çevirisi": google_translation
                })
            else:
                data.append({
                    "İngilizce Cümle": sentence,
                    "Google Çevirisi": google_translation
                })

            # İlerleme durumunu her %10'da bir yazdır (Sadece ilk kez yazdır)
            if num_sentences > 0:
                progress = int(((index + 1) / num_sentences) * 100)
                if progress % 10 == 0 and progress != last_printed_progress:
                    print(f"İlerleme: %{progress} tamamlandı.")
                    last_printed_progress = progress

        except Exception as e:
            print(f"Cümlenin işlenmesinde hata: {str(e)}")

    # Verileri bir pandas DataFrame'e dönüştür
    df = pd.DataFrame(data)
    
    # Excel dosyasına kaydet
    try:
        df.to_excel(output_excel, index=False)
        print(f"Çeviriler {output_excel} dosyasına kaydedildi.")
    except Exception as e:
        print(f"Excel dosyasına kaydedilirken hata oluştu: {str(e)}")

    # Zaman ölçümünü bitir ve süreyi hesapla
    end_time = time.time()
    elapsed_time = end_time - start_time

    # Geçen süreyi saat, dakika, saniye ve milisaniye olarak göster
    hours = int(elapsed_time // 3600)
    minutes = int((elapsed_time % 3600) // 60)
    seconds = int(elapsed_time % 60)
    milliseconds = int((elapsed_time - int(elapsed_time)) * 1000)

    time_str = f"Toplam süre: {seconds} saniye"
    if minutes > 0:
        time_str = f"Toplam süre: {minutes} dakika {time_str}"
    if hours > 0:
        time_str = f"{hours} saat {time_str}"
    time_str += f" ({minutes} dakika, {seconds} saniye, {milliseconds} milisaniye)"

    print(f"Çeviri işlemi tamamlandı. {time_str}")

def open_docx(file_path):
    if os.path.isfile(file_path) and file_path.endswith('.docx'):

        windows = gw.getWindowsWithTitle(os.path.basename(file_path))

        # Eğer dosya açık değilse aç
        if not any(window.title.endswith(os.path.basename(file_path)) for window in windows):
            os.startfile(file_path)
        else:
            # Dosya zaten açıksa ön plana getir
            windows = gw.getWindowsWithTitle(os.path.basename(file_path))
            if windows:
                # Dosya zaten açık, ön plana getir
                time.sleep(2)  # 2 saniye bekle
                window = windows[0]
                window.activate()
    else:
        print("Hata: Geçerli bir .docx dosyası seçilmedi.")

def close_docx(file_path):
    # Açık pencereleri kontrol et
    windows = gw.getWindowsWithTitle(os.path.basename(file_path))

    if windows:
        # İlk bulunan pencereyi kapat
        window = windows[0]
        window.close()
    else:
        pass

# Bu fonksiyon, bulunduğunuz dizindeki .docx dosyalarını listeler
def list_docx_files():
    return [file for file in os.listdir() if file.endswith(".docx")]

# Kullanıcıdan dosya yolu al
def get_file_path():
    docx_files = list_docx_files()
    
    # Dosya seçim soruları
    if docx_files:
        questions = [
            inquirer.List('file_choice',
                          message="Çevrilecek .docx dosyasını seçin veya yolunu girin:",
                          choices=docx_files + ['Dosya yolunu seçin'])
        ]
    else:
        questions = [
            inquirer.List('file_choice',
                          message="Bulunduğunuz dizinde .docx dosyası bulunamadı, lütfen dosya yolunu seçin.",
                          choices=['Dosya yolunu seçin'])
        ]

    # Seçimi al
    answer = inquirer.prompt(questions)['file_choice']
    file_path = ""
    # Eğer kullanıcı 'Dosya yolunu seçin' seçeneğini seçerse, askopenfilename ile dosya yolunu al
    if answer == 'Dosya yolunu seçin':
        while not file_path:
            # Tk arayüzünü gizlemeden önce oluştur ve öne getir
            root = Tk()
            root.withdraw()  # Ana pencereyi gizle
            root.lift()  # Pencereyi ön plana getir
            root.attributes('-topmost', True)  # Pencerenin en üstte görünmesini sağla
            file_path = askopenfilename(filetypes=[("Word files", "*.docx")])
            
            if not file_path:
                print("Hata: Hiçbir dosya seçilmedi, lütfen bir dosya seçin.")

        # Seçilen dosya adını yazdır
        print(f"Seçilen dosya: {file_path}")
        return file_path
    
    # Seçilen dosya adını yazdır
    print(f"Seçilen dosya: {file_path}")

    # Dizinde seçilen dosya adı tam dosya yoluna dönüştürülür
    return os.path.abspath(answer)


def get_output_excel(docx_filename):
    # Varsayılan dosya adı
    default_filename = "translated_sentences.xlsx"
    
    # .docx dosyasının adını kullanarak yeni bir ad türet
    # .docx dosyasının adını kullanarak yeni bir ad türet (os modülü olmadan)
    base_name = docx_filename.split('/')[-1].split('\\')[-1]  # Yolun sonundaki dosya adını al
    base_name = base_name.rsplit('.', 1)[0]
    suggested_filename = base_name + "_translated.xlsx"
    
    # Kullanıcıya seçenekler sun
    questions = [
        inquirer.List(
            'filename_choice',
            message="Oluşturulacak Excel dosyasının adını seçin:",
            choices=[
                suggested_filename,  # Seçilen .docx adının "_translated" eklenmiş hali
                default_filename,    # Varsayılan dosya adı
                'Kendi dosya adını gir'  # Kullanıcıdan özel ad girmesini iste
            ],
            carousel=True  # Menü sonuna gelince başa dönülmesini sağlar
        )
    ]

    # Tekrar konsol odaklanması sorunu olmasın diye `inquirer.prompt()` bir kere çalıştırılır
    filename_choice = inquirer.prompt(questions)['filename_choice']
    
    # Kullanıcı özel bir dosya adı girmek isterse
    if filename_choice == 'Kendi dosya adını gir':
        questions = [
            inquirer.Text('custom_filename',
                          message="Oluşturulacak Excel dosyasının adını girin (uzantı eklemenize gerek yok)")
        ]
        custom_filename = inquirer.prompt(questions)['custom_filename'].strip()
        
        # Kullanıcıdan gelen dosya adına ".xlsx" uzantısı ekle
        if not custom_filename.endswith(".xlsx"):
            custom_filename += ".xlsx"
            
        return custom_filename
    
    # Seçilen dosya adını döndür
    return filename_choice

def ask_chatgpt_usage():
    # Kullanıcıya ChatGPT çevirisi yapıp yapmak istemediğini sor
    questions = [
        inquirer.Confirm('use_chatgpt',
                         message="ChatGPT çevirisi eklemek ister misiniz?",
                         default=False)
    ]
    return inquirer.prompt(questions)['use_chatgpt']

if __name__ == "__main__":
    # 1. Dosya yolunu al
    file_path = get_file_path()

    # Performans Kaybı Yüzünden Rafa Kaldırıldı.
    # Tüm paragrafları al
    #all_sentences = extract_text_from_docx(file_path)
    #print(f"Dosyadaki toplam satır sayısı: {len(all_sentences)}")

    # 2. Excel dosyasının adını al
    output_excel = get_output_excel(file_path)
    print(f"Oluşturulacak Excel dosyasının adı: {output_excel}")

    # 3. ChatGPT kullanımını al
    use_chatgpt = ask_chatgpt_usage()

    api_key = None
    if use_chatgpt:
        print("OpenAI API anahtarına ihtiyacınız var. Anahtarınızı şu adresten alabilirsiniz: https://platform.openai.com/account/api-keys")
        api_key = input("Lütfen OpenAI API anahtarınızı girin: ")

    # Fonksiyonu çalıştır
    print("Yükleniyor...")
    translate_and_save_to_excel(file_path, output_excel, use_chatgpt, api_key)

    excel_formatter.format_excel(output_excel)
