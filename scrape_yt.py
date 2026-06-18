import pandas as pd
from googleapiclient.discovery import build
import re

# ================= KONFIGURASI =================
API_KEY = 'API DISINI'
TOPIK = 'AI di Era Digital' # Sesuaikan dengan topik kamu

# Masukkan hingga 20 URL (atau lebih) di dalam daftar ini
DAFTAR_URL_VIDEO = [
'https://www.youtube.com/live/ENIvRF0Bb2c?si=uYXXAH6NrvTmMLRP'

    # Tambahkan koma di akhir setiap link kecuali yang paling terakhir
]
# ===============================================

def get_video_id(url):
    """Fungsi untuk mengekstrak Video ID dari URL YouTube"""
    match = re.search(r'(?:v=|\/)([0-9A-Za-z_-]{11}).*', url)
    return match.group(1) if match else None

def scrape_multiple_videos(api_key, video_urls, topik):
    youtube = build('youtube', 'v3', developerKey=api_key)
    semua_data_komentar = []
    
    print(f"🚀 Memulai proses scraping untuk {len(video_urls)} video...\n")

    for index, url in enumerate(video_urls, start=1):
        video_id = get_video_id(url)
        if not video_id:
            print(f"[{index}/{len(video_urls)}] ⚠️ URL tidak valid, melewati: {url}")
            continue

        print(f"[{index}/{len(video_urls)}] ⏳ Memproses Video ID: {video_id}...")

        # 1. Mengambil informasi Video (Judul dan Channel)
        try:
            video_request = youtube.videos().list(part="snippet", id=video_id)
            video_response = video_request.execute()
            
            if not video_response['items']:
                print(f"    ❌ Video tidak ditemukan atau disetel privat.")
                continue

            judul_video = video_response['items'][0]['snippet']['title']
            channel = video_response['items'][0]['snippet']['channelTitle']
            print(f"    ✅ Judul: {judul_video[:40]}... | Channel: {channel}")
            
        except Exception as e:
            print(f"    ❌ Gagal mengambil data video: {e}")
            continue

        # 2. Mengambil Seluruh Komentar (Tanpa Batas)
        next_page_token = None
        jumlah_diambil = 0

        while True:
            try:
                request = youtube.commentThreads().list(
                    part="snippet",
                    videoId=video_id,
                    maxResults=100,
                    order="relevance", # Mengambil komentar "Teratas" (Top Comments)
                    pageToken=next_page_token,
                    textFormat="plainText"
                )
                response = request.execute()

                for item in response['items']:
                    thread_id = item['id']
                    comment = item['snippet']['topLevelComment']['snippet']
                    
                    semua_data_komentar.append({
                        "Topik": topik,
                        "Video_ID": video_id,
                        "Judul_Video": judul_video,
                        "Channel": channel,
                        "URL_Video": f"https://www.youtube.com/watch?v={video_id}",
                        "Nama_Komentator": comment.get('authorDisplayName', 'Anonim'),
                        "Teks_Komentar": comment.get('textDisplay', ''),
                        "Jumlah_Like_Komentar": comment.get('likeCount', 0),
                        "Tanggal_Komentar": comment.get('publishedAt', ''),
                        "Tipe_Komentar": "Komentar Utama"
                    })
                    jumlah_diambil += 1

                    # Mengambil SELURUH balasan (replies) dengan pagination
                    total_reply_count = item['snippet']['totalReplyCount']
                    if total_reply_count > 0:
                        try:
                            reply_page_token = None
                            while True:
                                reply_request = youtube.comments().list(
                                    part="snippet",
                                    parentId=thread_id,
                                    maxResults=100,
                                    pageToken=reply_page_token,
                                    textFormat="plainText"
                                )
                                reply_response = reply_request.execute()
                                
                                for reply_item in reply_response.get('items', []):
                                    reply = reply_item['snippet']
                                    semua_data_komentar.append({
                                        "Topik": topik,
                                        "Video_ID": video_id,
                                        "Judul_Video": judul_video,
                                        "Channel": channel,
                                        "URL_Video": f"https://www.youtube.com/watch?v={video_id}",
                                        "Nama_Komentator": reply.get('authorDisplayName', 'Anonim'),
                                        "Teks_Komentar": reply.get('textDisplay', ''),
                                        "Jumlah_Like_Komentar": reply.get('likeCount', 0),
                                        "Tanggal_Komentar": reply.get('publishedAt', ''),
                                        "Tipe_Komentar": "Balasan (Reply)"
                                    })
                                    jumlah_diambil += 1
                                
                                reply_page_token = reply_response.get('nextPageToken')
                                if not reply_page_token:
                                    break # Tidak ada balasan lagi
                                    
                        except Exception as e:
                            print(f"        ⚠️ Gagal mengambil balasan untuk thread {thread_id}: {e}")

                # Cek halaman berikutnya untuk komentar utama
                next_page_token = response.get('nextPageToken')
                if not next_page_token:
                    break # Jika tidak ada komentar utama lagi
                    
            except Exception as e:
                print(f"    ❌ Fitur komentar dinonaktifkan atau terjadi error.")
                break
        
        print(f"    📥 Berhasil mengambil total {jumlah_diambil} komentar beserta balasannya.\n")

    # 3. Export semua data ke dalam satu file Excel
    if semua_data_komentar:
        df = pd.DataFrame(semua_data_komentar)
        
        # Merapikan format zona waktu tanggal agar didukung oleh Excel
        df['Tanggal_Komentar'] = pd.to_datetime(df['Tanggal_Komentar']).dt.tz_localize(None)
        
        nama_file = "Kumpulan_Komentar_YouTube.xlsx"
        
        print(f"💾 Mengekspor total {len(semua_data_komentar)} komentar ke Excel...")
        df.to_excel(nama_file, index=False, engine='openpyxl')
        print(f"🎉 Selesai! Data berhasil disimpan di: {nama_file}")
    else:
        print("⚠️ Tidak ada satupun komentar yang berhasil diambil dari seluruh link.")

if __name__ == "__main__":
    scrape_multiple_videos(API_KEY, DAFTAR_URL_VIDEO, TOPIK)
