import streamlit as st
import google.generativeai as genai

# Gantilah dengan API key Anda
API_KEY = "AIzaSyAyVeEZmXYLEW_Hy5xQ0oSrxi-DrMGdqSY"

# Konfigurasi API key
genai.configure(api_key=API_KEY)

# Pilih model Gemini (misalnya, Gemini 1.5 Pro)
model = genai.GenerativeModel("gemini-1.5-pro")
if 'kond' not in st.session_state:
    st.session_state.kond = False
if 'kond1' not in st.session_state:
    st.session_state.kond1 = {'kondisi1':False,'kondisi2':False,'kondisi3':False}
if 'kumpulan' not in st.session_state:
    st.session_state.kumpulan = []
if 'nilai' not in st.session_state:
    st.session_state.nilai = 0
st.markdown('''
    <style>
           #judul{
             font-family:broadway;
             color:green;
            font-size:30px;
            text-shadow: 2px 2px 2px red
           }
            .pengantar{
                text-align:justify;
                font-family:"comic sans ms";
                font-size:20px;
                border:2px solid black;
                border-radius:10px;
                padding: 5px;
                margin:5px;
                background-color:pink;
                color:black;
            }
            #rps{
                font-family:"bauhaus 93";
                font-size:16px;
                color:white;
                border: 2px solid black;
                background-color:red;
                border-radius:10px;
                width:300px;
                padding:3px;
                text-align:center;
                margin:5px;
                box-shadow: 2px 2px 2px 2px yellow;
                animation-name:jalankan;
                animation-duration:2s;
                animation-iteration-count:infinite;
                animation-direction:alternate;
            }
            @keyframes jalankan{
                0%{background-color:red; color:white;box-shadow: -2px -2px 2px 2px yellow;}
                100%{background-color:blue; color:orange;box-shadow: -5px -5px 5px 5px green;}
            }
    </style>
    ''',unsafe_allow_html=True)
st.markdown('''
            
            <div id="judul">
            Media Belajar Coding Visual Basic Application for Excel
            </div>
            ''',unsafe_allow_html=True)
st.sidebar.title("Masukan identitas Anda")
if "nama" not in st.session_state:
    st.session_state["nama"]=""
if "NIM" not in st.session_state:
    st.session_state["NIM"]=""
st.session_state["nama"]=st.sidebar.text_input("Masukan Nama Anda: ",value=st.session_state["nama"])
st.session_state["NIM"]=st.sidebar.text_input("Masukan Nim Anda: ",value=st.session_state["NIM"])
if st.sidebar.button("Pelajaran pertama"):
    st.session_state.kond = False
    bagian = st.tabs(["Kata Pengantar","Mengenal Module", "Contoh Media", "Penilaian dan Masukan"])
    with bagian[0]:
        st.markdown('''
            <div class="pengantar">Mathematical Entrepreneurship 3 merupakan integrasi kemampuan kewirusaan mahasiswa
            dengan ilmu matematika melalui pengembangan media matematika dengan pembelajaran
            koding salah satu adalah Visual Basic Application for Excel. Dengan menciptakan Media
            Pembelajaran berbasis koding dapat membuka peluang besar dalam bisnis pengembangan
            media alat peraga matematika ke sekolah-sekolah dari SD, SMP, dan SMA. tentunya pengembangan ini
            melibatkan AI generator sebagai petunjuk dasar-dasar pembuatan media berbasis koding</div>
            <div class="pengantar">
                    <ol>Kemampuan yang harus dikuasai diantaranya 
                        <li> penguasaan Aplikasi TIK dalam Pembelajaran Matematika </li>
                        <li> Mathematical Entrepereneurship 1 </li>
                    </ol>
             </div>
        ''', unsafe_allow_html=True)
        st.markdown('''
            <div id="rps"> Silahkan baca RPS di bawah ini </div>
            <iframe src="https://drive.google.com/file/d/1AhRco8iua5UON7bEIaGT_z6gTlOILfGH/preview" 
                        width="800" height="600 title="_konsep">
            </iframe>'
            <div style="font-family:'brush script mt'; font-size:20px; font-weight:bold">Referensi</div>
            ''',unsafe_allow_html=True)
        gambar = st.columns(2,vertical_alignment='center',border=True)
        with gambar[0]:
            st.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRkbXZuI8FeiRZZ4gphV9d1CECRNqnXpgGazw&s")
            st.markdown("<a href='https://shopee.co.id/BUKU-MEDIA-DASAR-MATEMATIKA-MENGGUNAKAN-VISUAL-BASIC-FOR-APPLICATION-MICROSOFT-EXCEL-i.311279133.20714537804'>Buku Media VBA Dasar-dasar</a>",unsafe_allow_html=True)
        with gambar[1]:
            st.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQWoIK5cS452bB-aYYLOha_LnnZLEB9CFbbuQ&s")
            st.markdown("<a href='https://www.tokopedia.com/irfanbooks/buku-media-pembelajaran-matematika-berbasis-ict-dengan-vba-ms-excel'>Aplikasi TIK dalam Pembelajaran Matematika</a>",unsafe_allow_html=True)
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:20px;">Media Pembelajaran Matematika menurut Jean Piaget</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:yellow">Jean Piaget mengembangkan teori perkembangan kognitif yang membagi perkembangan 
                anak ke dalam empat tahap:</div>
            <div class="pengantar">
                    <ol>
                        <li> Sensori Motor (0-2 tahun) 
                        <ul>
                            <li> Anak belajar melalui pengalaman langsung dan indera.</li>
                            <li>Media yang sesuai: Objek fisik seperti balok angka, mainan interaktif, atau alat manipulatif sederhana.</li>
                        </ul>
                        </li>
                        <li>Praoperasional (2-7 tahun)
                        <ul>
                            <li>Anak mulai menggunakan simbol, tetapi masih egosentris dan berpikir konkret.</li>
                            <li>Media yang sesuai: Gambar, cerita visual, kartu angka, permainan peran dengan benda nyata.</li>
                        </ul>
                        </li>
                        <li>Operasi Konkret (7-11 tahun)
                        <ul>
                            <li>Anak mulai berpikir logis tetapi masih membutuhkan benda konkret.</li>
                            <li>Media yang sesuai: Papan permainan matematika, balok pecahan, diagram, eksperimen langsung.</li>
                        </ul>
                        </li>
                        <li>Operasi Formal (11 tahun ke atas)
                        <ul>
                            <li>Anak mampu berpikir abstrak dan logis.</li>
                            <li>Media yang sesuai: Simulasi komputer, soal cerita, diskusi masalah kompleks, dan media digital seperti Pygame untuk memvisualisasikan konsep abstrak.</li>
                        </ul>
                        </li>
                    </ol>
            </div>
            ''',unsafe_allow_html=True)
    with bagian[1]:
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:20px;">Koding yang Perlu Dikuasai dalam Module Sub VBA for Excel</div>
                    ''',unsafe_allow_html=True)
            
        st.markdown('''
            <div class="pengantar" style="background-color:blue;color:yellow">Agar dapat menggunakan VBA untuk Excel secara efektif,
                     ada beberapa aspek penting dalam Module Sub yang perlu 
                    dikuasai. Berikut adalah daftar konsep dan koding utama 
                    yang harus dipahami beserta contohnya:</div>''',unsafe_allow_html=True)
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Mengakses dan Mengubah Data dalam Sel</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Menulis dan membaca data dari sel Excel.<br>
            🔹  Contoh: Mengisi Sel dengan Teks</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub IsiSel()
                    Range("A1").Value = "Belajar VBA"
                End Sub
                ''', language="vb")
        st.write("contoh video pertama")
        with st.expander("Lihat Video Pertama"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740981398/Recording_33_tjwtwh.mp4")
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">
            🔹  Contoh: Membaca Data dari Sel</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub BacaSel()
                    Dim teks As String
                    teks = Range("A1").Value
                    MsgBox "Isi sel A1 adalah: " & teks
                End Sub
                ''', language="vb")
        st.write("contoh video kedua")
        with st.expander("Lihat Video kedua"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740982532/Recording_34_cighxl.mp4")
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Looping (Perulangan)(Subroutine)</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Looping digunakan untuk mengulang proses.<br>
            🔹 Contoh: Mengisi Kolom A dengan Angka 1-10</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub LoopIsiKolom()
                    Dim i As Integer
                    For i = 1 To 10
                        Cells(i, 1).Value = i
                    Next i
                End Sub
                ''', language="vb")
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">
            🔹  Contoh: Looping Menggunakan While</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub LoopWhile()
                    Dim i As Integer
                    i = 1
                    While i <= 5
                        Cells(i, 2).Value = "Baris " & i
                        i = i + 1
                    Wend
                End Sub
                ''', language="vb")
        st.write("Contoh Video Ketiga")
        with st.expander("Lihat Video Ketiga"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740983645/Recording_35_ahpbrp.mp4")
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Kondisi IF dan Select Case</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Digunakan untuk pengambilan keputusan dalam VBA.<br>
            🔹  Contoh: IF ELSE</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub CekNilai()
                    Dim nilai As Integer
                    nilai = Range("B1").Value
                    If nilai >= 60 Then
                        MsgBox "Lulus"
                    Else
                        MsgBox "Tidak Lulus"
                    End If
                End Sub
                ''', language="vb")
        st.write("Contoh Video Keempat")
        with st.expander("Lihat Video keempat"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740991400/Recording_36_ieybn1.mp4")
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">
            🔹  Contoh: Select Case</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub CekGrade()
                    Dim nilai As Integer
                    nilai = Range("B1").Value
    
                    Select Case nilai
                    Case 80 To 100
                        MsgBox "A"
                    Case 70 To 79
                        MsgBox "B"
                    Case 60 To 69
                        MsgBox "C"
                    Case Else
                        MsgBox "D"
                    End Select
                End Sub
                ''', language="vb")
        st.write("Contoh Video Kelima")
        with st.expander("Lihat Video kelima"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740992391/Recording_37_yefvcq.mp4")
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Menggunakan Variabel dan Tipe Data</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Variabel digunakan untuk menyimpan data sementara.<br>
            🔹  Contoh: Deklarasi Variabel</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub VariabelVBA()
                    Dim nama As String
                    Dim usia As Integer
                    Dim tinggi As Double
                    nama = "Budi"
                    usia = 25
                    tinggi = 170.5
                    MsgBox "Nama: " & nama & vbNewLine & "Usia: " & usia & vbNewLine & "Tinggi: " & tinggi
                End Sub
                ''', language="vb")
        st.write("Contoh Video Keenam")
        with st.expander("Lihat Video keenam"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740993059/Recording_38_st1oma.mp4")
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Menggunakan InputBox</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Digunakan untuk meminta input dari pengguna.<br>
            🔹  Contoh: InputBox untuk Mengisi Sel</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub IsiDenganInput()
                    Dim inputData As String
                    inputData = InputBox("Masukkan teks:")
                    Range("A1").Value = inputData
                End Sub
                ''', language="vb")
        st.write("Contoh Video Ketujuh")
        with st.expander("Lihat Video ketujuh"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740993648/Recording_39_t7a37c.mp4")
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Menggunakan Arrays</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Array digunakan untuk menyimpan banyak nilai dalam satu variabel.<br>
            🔹  Contoh: Array Sederhana</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub ArrayVBA()
                    Dim angka(4) As Integer
                    Dim i As Integer
    
                    ' Mengisi array
                    For i = 0 To 4
                        angka(i) = (i + 1) * 10
                    Next i
    
                    ' Menampilkan isi array
                    MsgBox "Nilai pertama: " & angka(0) & vbNewLine & "Nilai terakhir: " & angka(4)
                End Sub
                ''', language="vb")
        st.write("Contoh Video Kedelapan")
        with st.expander("Lihat Video kedelapan"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740994201/Recording_40_spsfsu.mp4")
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Mengontrol Sheet dan Workbook</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Manipulasi lembar kerja (worksheet) dan buku kerja (workbook).<br>
            🔹  Contoh: Menambahkan Sheet Baru</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub TambahSheet()
                    Sheets.Add
                End Sub
                ''', language="vb")
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">
            🔹  Contoh: Mengganti Nama Sheet</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub UbahNamaSheet()
                    ActiveSheet.Name = "Data Baru"
                End Sub
                ''', language="vb")
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">
            🔹  Contoh: Menutup Workbook</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub TutupWorkbook()
                        ThisWorkbook.Close SaveChanges:=False
                End Sub
                ''', language="vb")
        st.write("Contoh Video Kesembilan")
        with st.expander("Lihat Video kesembilan"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740995478/Recording_41_jwdzxh.mp4")
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Event Handling (Menjalankan Kode Saat Event Terjadi)</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Event menangani aksi tertentu dalam Excel.<br>
            🔹  Contoh: Menjalankan Sub Saat Workbook Dibuka</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Private Sub Workbook_Open()
                    MsgBox "Selamat datang di Excel VBA!"
                End Sub
                ''', language="vb")
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Event menangani aksi tertentu dalam Excel.<br>
            🔹  Contoh: Menjalankan Sub Saat Sel Diklik</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Private Sub Worksheet_SelectionChange(ByVal Target As Range)
                    MsgBox "Anda memilih sel " & Target.Address
                End Sub
                ''', language="vb")
        st.write("Contoh Video Kesepuluh")
        with st.expander("Lihat Video kesepuluh"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740996240/Recording_42_zbqpi5.mp4")
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:16px; color:green">Event Handling (Menjalankan Kode Saat Event Terjadi)</div>
                    ''',unsafe_allow_html=True)
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Menggunakan Application.OnTime untuk Timer<br>
            🔹  Contoh: Menjalankan Sub secara Otomatis dalam 5 Detik</div>''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Sub TimerVBA()
                    Application.OnTime Now + TimeValue("00:00:05"), "TampilkanPesan"
                End Sub
                ''', language="vb")
        st.write("Contoh Video Kesebelas")
        with st.expander("Lihat Video kesebelas"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1740997523/Recording_43_b2mmv6.mp4")
        st.markdown('''
            <div class="pengantar" style="background-color:cyan;color:black">Menyimpan dan Membuka Workbook<br>
            🔹  Contoh: Menjalankan Sub secara Otomatis dalam 5 Detik</div>''',unsafe_allow_html=True)
        st.write("Contoh Video Kedua belas (Menyimpan dan Membuka workbook)")
        with st.expander("Lihat Video kedua belas (Menyimpan dan Membuka Workbook)"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1742123425/Recording_50_cz9o3u.mp4")
    with bagian[2]:
        st.markdown('''
                    <div style="font-family:'snap itc'; font-size:20px; color:green">Contoh Operasi Hitung</div>
                    ''',unsafe_allow_html=True)
        with st.expander("Langkah Pertama (Penjumlahan)"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1741000807/Recording_44_fwmeap.mp4")
        with st.expander("Langkah Kedua (Pengurangan)"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1741002705/Recording_45_uamq6j.mp4")
        with st.expander("Langkah Ketiga (Perkalian)"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1741003182/Recording_46_xcglrz.mp4")
        with st.expander("Langkah Keempat (Pembagian)"):
            st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1741003881/Recording_47_jwzf5r.mp4")   
        st.markdown(''' <div style="font-family:'comic sans ms'; font-size:16px">
            Dari Video di atas, kita dapat mengembangkan lebih lagi tentang operasi hitung lainnya
            </div>
            <div style="font-family:'comic sans ms'; font-size:16px">Kita dapat mencari informasinya dari AI Generator contoh ChatGPT</div>
            ''',unsafe_allow_html=True)
        st.image('https://res.cloudinary.com/ikip-siliwangi/image/upload/v1741004945/Bantuan_AI_knpzee.png')
        tulisan_html='''
        <!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload File to Firebase</title>
    <style>
      body {
         font-family: Arial, sans-serif;
         display: flex;
         flex-direction: column;
         align-items: center;
         justify-content: center;
         height: 100vh;
         background-color: #f0f0f0;
         margin: 0;
      }
      h1 {
         color: #333;
      }
      #fileInput {
         margin-bottom: 20px;
      }
      #progressContainer {
         margin-top: 20px;
         width: 100%;
         max-width: 400px;
      }

      progress {
         width: 100%;
      }

</style>
</head>
<body>
    <h1>Upload File VBA for Excel</h1>
    <input type="file" id="fileInput">
    <button id="uploadButton">Upload</button>
    <div id="progressContainer">
        <progress id="uploadProgress" value="0" max="100"></progress>
    </div>
    
    <script type="module">
            // Import modul Firebase yang diperlukan
      import { initializeApp } from "https://www.gstatic.com/firebasejs/11.4.0/firebase-app.js";
      import { getStorage, ref, uploadBytesResumable, getDownloadURL } from "https://www.gstatic.com/firebasejs/11.4.0/firebase-storage.js";

      // Firebase configuration
      const firebaseConfig = {
         apiKey: "AIzaSyCS-04HW1WAL3aLJkA7Zrmz2iedWVeaKKk",
        authDomain: "helpful-rope-333907.firebaseapp.com",
        databaseURL: "https://helpful-rope-333907-default-rtdb.firebaseio.com",
        projectId: "helpful-rope-333907",
        storageBucket: "helpful-rope-333907.appspot.com",
        messagingSenderId: "854982010261",
        appId: "1:854982010261:web:586875587b2a81f679829c",
        measurementId: "G-8XEM0H16H0"
      };

      // Initialize Firebase
      const app = initializeApp(firebaseConfig);
      const storage = getStorage(app);

      // Get elements
      const fileInput = document.getElementById('fileInput');
      const uploadButton = document.getElementById('uploadButton');
      const uploadProgress = document.getElementById('uploadProgress');

      uploadButton.addEventListener('click', () => {
         const file = fileInput.files[0];
         if (file) {
            uploadFile(file);
         } else {
            alert('Please choose a file first.');
         }
      });

      function uploadFile(file) {
         const storageRef = ref(storage, 'uploads/' + file.name);
         const uploadTask = uploadBytesResumable(storageRef, file);

         uploadTask.on('state_changed', 
            (snapshot) => {
                  // Observe state change events such as progress, pause, and resume
                  const progress = (snapshot.bytesTransferred / snapshot.totalBytes) * 100;
                  uploadProgress.value = progress;
                  console.log('Upload is ' + progress + '% done');
            }, 
            (error) => {
                  // Handle unsuccessful uploads
                  console.error('Upload failed:', error);
            }, 
            () => {
                  // Handle successful uploads on complete
                  getDownloadURL(uploadTask.snapshot.ref).then((downloadURL) => {
                     console.log('File available at', downloadURL);
                  });
            }
         );
      }

    </script>
</body>
</html>

        '''
        st.components.v1.html(tulisan_html, height=400)

    with bagian[3]:
        tulisan_html1=f'''
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <style>
                #komentar{{
                    border : 2px solid black;
                    border-radius:10px;
                    padding:5px;
                    background-color:yellow;
                }}
            </style>
        </head>
        <body>
            <div id="komentar"><strong><h3 style="color:red; font-family:'snap itc';font-size:18px">Masukan Komentar</h3></strong><br>
                <label for="nama">Nama:</label>
                <input type="text" id="nama" name="nama" value={st.session_state.nama}>
                <label for="nim">NIM: </label>
                <input type="text" id="nim" name="nim" value={st.session_state.NIM}><br><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:15px">Masukan Pendapat Anda Ketika Mempelajari Media ini</h4></strong><br>
                <textarea id="masukan"></textarea><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:16px">Masukan Kesulitan Anda Ketika Mempelajari Media ini</h4></strong><br>
                <textarea id="masuk1"></textarea><br>
                <button id="kirim">Kirim</button>
            </div>
            <script type="module">
            // Firebase configuration
            const firebaseConfig = {{
                apiKey: "AIzaSyCS-04HW1WAL3aLJkA7Zrmz2iedWVeaKKk",
                authDomain: "helpful-rope-333907.firebaseapp.com",
                databaseURL: "https://helpful-rope-333907-default-rtdb.firebaseio.com",
                projectId: "helpful-rope-333907",
                storageBucket: "helpful-rope-333907.appspot.com",
                messagingSenderId: "854982010261",
                appId: "1:854982010261:web:586875587b2a81f679829c",
                measurementId: "G-8XEM0H16H0"
            }};

            // Initialize Firebase
            import {{ initializeApp }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-app.js";
            import {{ getDatabase, ref, set  }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-database.js";

            const app = initializeApp(firebaseConfig);
            const db = getDatabase(app);
            var nim = document.getElementById("nim")
            var nama = document.getElementById("nama")
            var masukan = document.getElementById("masukan")
            var masuk1 = document.getElementById("masuk1")
            document.getElementById("kirim").addEventListener("click",()=>{{
                set(ref(db, 'masukan1/' + nim.value), {{
                    Nama: nama.value,
                    Nim: nim.value,
                    Masukan:masukan.value,
                    Masukan1:masuk1.value
                }})
                .then(()=>{{
                        alert("Sudah Masuk")}})
                .catch((error)=>{{
                    alert(error)
                }})
            }})
            </script>
        
        </body>
        </html>
        '''
        st.components.v1.html(tulisan_html1,height=600)
        tulisan_html2=f'''
    <!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <style>
        #judul{{
            text-align:center;
            font-family:broadway;
            color:green;
            font-size:30px;
        }}
    </style>
</head>
<body style="background-color:cyan">
    <div id="judul"> ANGKET EVALUASI PEMBELAJARAN VBA FOR EXCEL DALAM PENDIDIKAN MATEMATIKA</div>
    <div>
        <div>
            <label for="nama">Nama</label>
            <input type="text" name="nama" id="nama" value={st.session_state['nama']}>
        </div>
        <div>
            <label for="nim">NIM</label>
            <input type="text" name="nim" id="nim" value={st.session_state['NIM']}>
        </div>
    </div>
    <div><strong>Petunjuk pengisian</strong></div>
    <div>Angket ini bertujuan untuk mengetahui pemahaman dan pengalaman Anda setelah mempelajari VBA for Excel dalam pendidikan matematika.<br>
        Berikan penilaian terhadap setiap pernyataan berikut dengan skala:<br>
        <ul>
            <li>1 = Sangat Kurang</li>
            <li>2 = Kurang</li>
            <li>3 = Cukup</li>
            <li>4 = Baik</li>
            <li>5 = Sangat Baik</li>
        </ul>
    <div>
        <ol type="A">
            <li>Pemahaman Konsep VBA for Excel</li>
            <ol>
                <li>Saya memahami konsep dasar VBA dan fungsinya dalam Excel.</li>
                <input type="radio" name="A1" value="1">
                <label for="A1">1</label>
                <input type="radio" name="A1" value="2">
                <label for="A1">2</label>
                <input type="radio" name="A1" value="3">
                <label for="A1">3</label>
                <input type="radio" name="A1" value="4">
                <label for="A1">4</label>
                <input type="radio" name="A1" value="5">
                <label for="A1">5</label>
                <li>Saya dapat menjelaskan perbedaan antara Macro dan VBA.</li>
                <input type="radio" name="A2" value="1">
                <label for="A2">1</label>
                <input type="radio" name="A2" value="2">
                <label for="A2">2</label>
                <input type="radio" name="A2" value="3">
                <label for="A2">3</label>
                <input type="radio" name="A2" value="4">
                <label for="A2">4</label>
                <input type="radio" name="A2" value="5">
                <label for="A2">5</label>
                <li>Saya dapat membuka Editor VBA dan membuat kode sederhana.</li>
                <input type="radio" name="A3" value="1">
                <label for="A3">1</label>
                <input type="radio" name="A3" value="2">
                <label for="A3">2</label>
                <input type="radio" name="A3" value="3">
                <label for="A3">3</label>
                <input type="radio" name="A3" value="4">
                <label for="A3">4</label>
                <input type="radio" name="A3" value="5">
                <label for="A3">5</label>
                <li>Saya memahami konsep variabel dan tipe data dalam VBA.</li>
                <input type="radio" name="A4" value="1">
                <label for="A4">1</label>
                <input type="radio" name="A4" value="2">
                <label for="A4">2</label>
                <input type="radio" name="A4" value="3">
                <label for="A4">3</label>
                <input type="radio" name="A4" value="4">
                <label for="A4">4</label>
                <input type="radio" name="A4" value="5"">
                <label for="A4">5</label>
                <li>Saya dapat menggunakan struktur kontrol seperti If...Then dan Loop dalam VBA.</li>
                <input type="radio" name="A5" value="1"">
                <label for="A5">1</label>
                <input type="radio" name="A5" value="2">
                <label for="A5">2</label>
                <input type="radio" name="A5" value="3">
                <label for="A5">3</label>
                <input type="radio" name="A5" value="4">
                <label for="A5">4</label>
                <input type="radio" name="A5" value="5">
                <label for="A5">5</label>
            </ol>
            <li>Implementasi VBA dalam Matematika</li>
            <ol>
                <li>Saya dapat membuat program VBA untuk perhitungan matematika sederhana (misalnya, menghitung rata-rata atau mencari nilai maksimum).</li>
                <input type="radio" name="B1" value="1">
                <label for="B1">1</label>
                <input type="radio" name="B1" value="2">
                <label for="B1">2</label>
                <input type="radio" name="B1" value="3">
                <label for="B1">3</label>
                <input type="radio" name="B1" value="4">
                <label for="B1">4</label>
                <input type="radio" name="B1" value="5">
                <label for="B1">5</label>
                <li>Saya dapat membuat fungsi VBA khusus (User Defined Function) untuk keperluan matematika.</li>
                <input type="radio" name="B2" value="1">
                <label for="B2">1</label>
                <input type="radio" name="B2" value="2">
                <label for="B2">2</label>
                <input type="radio" name="B2" value="3">
                <label for="B2">3</label>
                <input type="radio" name="B2" value="4">
                <label for="B2">4</label>
                <input type="radio" name="B2" value="5">
                <label for="B2">5</label>
                <li>Saya dapat menggunakan VBA untuk mengolah dan menganalisis data numerik secara otomatis.</li>
                <input type="radio" name="B3" value="1"">
                <label for="B3">1</label>
                <input type="radio" name="B3" value="2">
                <label for="B3">2</label>
                <input type="radio" name="B3" value="3">
                <label for="B3">3</label>
                <input type="radio" name="B3" value="4">
                <label for="B3">4</label>
                <input type="radio" name="B3" value="5">
                <label for="B3">5</label>
                <li>Saya dapat mengembangkan makro VBA untuk menyajikan data dalam bentuk grafik atau tabel.</li>
                <input type="radio" name="B4" value="1">
                <label for="B4">1</label>
                <input type="radio" name="B4" value="2">
                <label for="B4">2</label>
                <input type="radio" name="B4" value="3">
                <label for="B4">3</label>
                <input type="radio" name="B4" value="4">
                <label for="B4">4</label>
                <input type="radio" name="B4" value="5">
                <label for="B4">5</label>
                <li>Saya memahami cara mengintegrasikan VBA dengan formula Excel untuk menyelesaikan masalah matematika.</li>
                <input type="radio" name="B5" value="1">
                <label for="B5">1</label>
                <input type="radio" name="B5" value="2">
                <label for="B5">2</label>
                <input type="radio" name="B5" value="3">
                <label for="B5">3</label>
                <input type="radio" name="B5" value="4">
                <label for="B5">4</label>
                <input type="radio" name="B5" value="5">
                <label for="B5">5</label>
            </ol>
            <li>Efektivitas Pembelajaran</li>
            <ol>
                <li>Pembelajaran VBA for Excel membantu saya memahami konsep matematika lebih dalam.</li>
                <input type="radio" name="C1" value="1">
                <label for="C1">1</label>
                <input type="radio" name="C1" value="2">
                <label for="C1">2</label>
                <input type="radio" name="C1" value="3">
                <label for="C1">3</label>
                <input type="radio" name="C1" value="4">
                <label for="C1">4</label>
                <input type="radio" name="C1" value="5">
                <label for="C1">5</label>
                <li>Penggunaan VBA membuat saya lebih mudah dalam memproses data matematika dibandingkan dengan perhitungan manual.</li>
                <input type="radio" name="C2" value="1">
                <label for="C2">1</label>
                <input type="radio" name="C2" value="2">
                <label for="C2">2</label>
                <input type="radio" name="C2" value="3">
                <label for="C2">3</label>
                <input type="radio" name="C2" value="4">
                <label for="C2">4</label>
                <input type="radio" name="C2" value="5">
                <label for="C2">5</label>
                <li>Saya merasa lebih percaya diri dalam menggunakan Excel setelah mempelajari VBA.</li>
                <input type="radio" name="C3" value="1">
                <label for="C3">1</label>
                <input type="radio" name="C3" value="2">
                <label for="C3">2</label>
                <input type="radio" name="C3" value="3">
                <label for="C3">3</label>
                <input type="radio" name="C3" value="4">
                <label for="C3">4</label>
                <input type="radio" name="C3" value="5">
                <label for="C3">5</label>
                <li>Materi yang disampaikan dalam modul VBA mudah dipahami dan diterapkan.</li>
                <input type="radio" name="C4" value="1">
                <label for="C4">1</label>
                <input type="radio" name="C4" value="2">
                <label for="C4">2</label>
                <input type="radio" name="C4" value="3">
                <label for="C4">3</label>
                <input type="radio" name="C4" value="4">
                <label for="C4">4</label>
                <input type="radio" name="C4" value="5">
                <label for="C4">5</label>
                <li>Pembelajaran VBA membuat saya lebih tertarik dalam mengembangkan aplikasi matematika berbasis Excel.</li>
                <input type="radio" name="C5" value="1">
                <label for="C5">1</label>
                <input type="radio" name="C5" value="2">
                <label for="C5">2</label>
                <input type="radio" name="C5" value="3">
                <label for="C5">3</label>
                <input type="radio" name="C5" value="4">
                <label for="C5">4</label>
                <input type="radio" name="C5" value="5">
                <label for="C5">5</label>
            </ol>
            <li>Efektivitas Pembelajaran</li>
            <ol>
                <li>Saya dapat memodifikasi kode VBA yang sudah ada sesuai dengan kebutuhan saya.</li>
                <input type="radio" name="D1" value="1">
                <label for="D1">1</label>
                <input type="radio" name="D1" value="2">
                <label for="D1">2</label>
                <input type="radio" name="D1" value="3">
                <label for="D1">3</label>
                <input type="radio" name="C1" value="4">
                <label for="D1">4</label>
                <input type="radio" name="D1" value="5">
                <label for="D1">5</label>
                <li>Saya merasa mampu membuat program VBA sendiri untuk membantu perhitungan dalam penelitian atau tugas matematika.</li>
                <input type="radio" name="D2" value="1">
                <label for="D2">1</label>
                <input type="radio" name="D2" value="2">
                <label for="D2">2</label>
                <input type="radio" name="D2" value="3">
                <label for="D2">3</label>
                <input type="radio" name="D2" value="4">
                <label for="D2">4</label>
                <input type="radio" name="D2" value="5">
                <label for="D2">5</label>
                <li>Saya tertarik untuk mempelajari lebih lanjut VBA dan penggunaannya dalam bidang lain.</li>
                <input type="radio" name="D3" value="1">
                <label for="D3">1</label>
                <input type="radio" name="D3" value="2">
                <label for="D3">2</label>
                <input type="radio" name="D3" value="3">
                <label for="D3">3</label>
                <input type="radio" name="D3" value="4">
                <label for="D3">4</label>
                <input type="radio" name="D3" value="5">
                <label for="D3">5</label>
                <li>Saya dapat menyelesaikan permasalahan yang muncul saat menjalankan kode VBA secara mandiri atau dengan bantuan sumber lain.</li>
                <input type="radio" name="D4" value="1">
                <label for="C4">1</label>
                <input type="radio" name="D4" value="2">
                <label for="D4">2</label>
                <input type="radio" name="D4" value="3">
                <label for="D4">3</label>
                <input type="radio" name="D4" value="4">
                <label for="D4">4</label>
                <input type="radio" name="D4" value="5">
                <label for="D4">5</label>
                <li>Pembelajaran VBA membuat saya lebih tertarik dalam mengembangkan aplikasi matematika berbasis Excel.</li>
                <input type="radio" name="D5" value="1">
                <label for="D5">1</label>
                <input type="radio" name="D5" value="2">
                <label for="D5">2</label>
                <input type="radio" name="D5" value="3">
                <label for="D5">3</label>
                <input type="radio" name="D5" value="4">
                <label for="D5">4</label>
                <input type="radio" name="D5" value="5">
                <label for="D5">5</label>
            </ol>
        </ol>
    </div>
    </div>
    <button id="kirim">Kirimkan</button>
    <script type="module">
        var A1 = document.getElementsByName("A1")
        var A2 = document.getElementsByName("A2")
        var A3 = document.getElementsByName("A3")
        var A4 = document.getElementsByName("A4")
        var A5 = document.getElementsByName("A5")
        var B1 = document.getElementsByName("B1")
        var B2 = document.getElementsByName("B2")
        var B3 = document.getElementsByName("B3")
        var B4 = document.getElementsByName("B4")
        var B5 = document.getElementsByName("B5")
        var C1 = document.getElementsByName("C1")
        var C2 = document.getElementsByName("C2")
        var C3 = document.getElementsByName("C3")
        var C4 = document.getElementsByName("C4")
        var C5 = document.getElementsByName("C5")
        var D1 = document.getElementsByName("D1")
        var D2 = document.getElementsByName("D2")
        var D3 = document.getElementsByName("D3")
        var D4 = document.getElementsByName("D4")
        var D5 = document.getElementsByName("D5")
        var pilihan = [A1,A2,A3,A4,A5,B1,B2,B3,B4,B5,C1,C2,C3,C4,C5,D1,D2,D3,D4,D5]
        var jawaban = []
        const firebaseConfig = {{
                apiKey: "AIzaSyCS-04HW1WAL3aLJkA7Zrmz2iedWVeaKKk",
                authDomain: "helpful-rope-333907.firebaseapp.com",
                databaseURL: "https://helpful-rope-333907-default-rtdb.firebaseio.com",
                projectId: "helpful-rope-333907",
                storageBucket: "helpful-rope-333907.appspot.com",
                messagingSenderId: "854982010261",
                appId: "1:854982010261:web:586875587b2a81f679829c",
                measurementId: "G-8XEM0H16H0"
            }};

            // Initialize Firebase
            import {{ initializeApp }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-app.js";
            import {{ getDatabase, ref, set  }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-database.js";
            const app = initializeApp(firebaseConfig);
            const db = getDatabase(app);
            var nama = document.getElementById("nama")
            var nim = document.getElementById("nim")
        document.getElementById("kirim").addEventListener("click",()=>{{
            jawaban=[]
            for(var j in pilihan){{
                console.log(j)
                for(var i of pilihan[j]){{
                    if(i.checked){{
                        jawaban.push(i.value)
                    }}
                }}
            }}
            console.log(jawaban)
            set(ref(db, 'pilihan/' + nim.value), {{
                    Nama: nama.value,
                    Nim: nim.value,
                    pilihan:JSON.stringify(jawaban)
                }})
                .then(()=>{{
                        alert("Sudah Masuk")}})
                .catch((error)=>{{
                    alert(error)
                }})
        }})
    </script>
</body>
</html>
    '''
        st.components.v1.html(tulisan_html2,height=1600)
if "genai" not in st.session_state:
    st.session_state["genai"] = False
if "isian" not in st.session_state:
    st.session_state["isian"]=""
def nyalakan():
    st.session_state['genai'] = True
def matikan():
    st.session_state['genai'] = False

if st.sidebar.button("Pelajaran Kedua"):
    st.session_state.kond = False
    bagian1 = st.tabs(["Kata Pengantar","Contoh Penggunaan AI", "Media AI", "Penilaian dan Masukan"])
    with bagian1[2]:
        st.session_state['genai'] = False
        st.session_state["isian"]=""
        st.markdown('''
        <div style="font-family:'snap itc';font-size:20px">Contoh Koding hasil dari AI</div>
        ''',unsafe_allow_html=True)
        st.code('''
            Sub SolveEquation()
            ' Deklarasi variabel
            Dim x As Double

            ' Selesaikan persamaan 2x + 1 = 7
            x = (7 - 1) / 2

            ' Tampilkan solusi di MsgBox
            MsgBox "Solusi dari persamaan 2x + 1 = 7 adalah x = " & x

            ' Atau, tampilkan solusi di cell tertentu (misalnya, A1)
            'ThisWorkbook.Sheets("Sheet1").Range("A1").Value = x
            End Sub
        ''',language="vba")
        st.markdown('''
        <div style="font-family:'snap itc';font-size:20px">Hasil Media</div>
        ''',unsafe_allow_html=True)
        st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1741873180/Recording_48_xmn49z.mp4")
        st.markdown('''
        <div style="font-family:'snap itc';font-size:20px">Rancangan Pengembangan</div>
        ''',unsafe_allow_html=True)
        st.image("https://res.cloudinary.com/ikip-siliwangi/image/upload/v1741874890/rancangan1_ddpurw.jpg")
        st.markdown('''
        <div style="font-family:'snap itc';font-size:20px">Koding Hasil Pengembangan</div>
        ''',unsafe_allow_html=True)
        st.code('''
            Sub tampilan()
    'Deklarasi persamaan dengan tipe string
    Dim persamaan As String
    'Deklarasi lembar dengan tipe worksheet
    Dim lembar As Worksheet
    'Aktifkan sheet1
    Set lembar = Worksheets("Sheet1")
    'persamaan akan menyimpan teks dari inputbox sesuai pengguna memasukan
    persamaan = InputBox("Masukan Persamaan")
    'Menampilkan hasil penulisan dari pengguna
    lembar.Shapes("persamaan_linear").TextFrame.Characters.Text = persamaan
End Sub

Sub hasilnya()
    'Deklarasi persamaan dengan tipe string
    Dim persamaan As String
    'Deklarasi untuk variabel
    Dim nilai1, nilai2, nilai3 As Integer
    'Deklarasi lembar dengan tipe worksheet
    Dim lembar As Worksheet
    'Aktifkan sheet1
    Set lembar = Worksheets("Sheet1")
    'Ambil teks dari persamaan_linear ke persamaan
    persamaan = lembar.Shapes("persamaan_linear").TextFrame.Characters.Text
    'pisahkan teks persamaan dengan = menggunakan fungsi split
    pisahkan = Split(persamaan, "=")
    'ambil nilai pada ruas kanan pada urutan kedua atau indeks 1 dari pisahkan ke nilai3
    nilai3 = Trim(pisahkan(1))
    'pada ruas kiri pisahkan urutan pertama atau indeks 0 dari pisahkan dengan simbol + menggunakan fungsi split simpan di pisahkan1
    pisahkan1 = Split(pisahkan(0), "+")
    'ambil nilai pada bagian pisahkan1 pada urutan kedua atau indeks 1 dari pisahkan1 ke nilai2
    nilai2 = Trim(pisahkan1(1))
    'pada variabel x pisahkan urutan pertama atau indeks 0 dari pisahkan1 dengan simbol x menggunakan fungsi split simpan di pisahkan2
    pisahkan2 = Split(pisahkan1(0), "x")
    'ambil nilai pada bagian pisahkan2 pada urutan pertama atau indeks 0 dari pisahkan1 ke nilai1
    nilai1 = Trim(pisahkan2(0))
    'Menampilkan hasil akhir di gamabr teks yang bernama hasil_akhir
    lembar.Shapes("hasil_akhir").TextFrame.Characters.Text = "jika " & persamaan & " maka hasil x = " & (nilai3 - nilai2) / nilai1
End Sub

        ''', language="vba")
        st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1741877752/Recording_49_qfrpfp.mp4")
        st.markdown('''
        <div style="font-family:'snap itc';font-size:20px">Pengembangan</div>
        ''',unsafe_allow_html=True)
        st.write("Buat skenario untuk mengembangkan di atas misalkan kondisi penjumlahan atau pengurangan dan kondisi bilangan positif, nol, dan negatif")
        st.write("Masukan hasilnya jika sudah melakukan pengembangan")
        tulisan_html1='''
        <!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload File to Firebase</title>
    <style>
      body {
         font-family: Arial, sans-serif;
         display: flex;
         flex-direction: column;
         align-items: center;
         justify-content: center;
         height: 100vh;
         background-color: #f0f0f0;
         margin: 0;
      }
      h1 {
         color: #333;
      }
      #fileInput {
         margin-bottom: 20px;
      }
      #progressContainer {
         margin-top: 20px;
         width: 100%;
         max-width: 400px;
      }

      progress {
         width: 100%;
      }

</style>
</head>
<body>
    <h1>Upload File VBA for Excel</h1>
    <input type="file" id="fileInput">
    <button id="uploadButton">Upload</button>
    <div id="progressContainer">
        <progress id="uploadProgress" value="0" max="100"></progress>
    </div>
    
    <script type="module">
            // Import modul Firebase yang diperlukan
      import { initializeApp } from "https://www.gstatic.com/firebasejs/11.4.0/firebase-app.js";
      import { getStorage, ref, uploadBytesResumable, getDownloadURL } from "https://www.gstatic.com/firebasejs/11.4.0/firebase-storage.js";

      // Firebase configuration
      const firebaseConfig = {
         apiKey: "AIzaSyCS-04HW1WAL3aLJkA7Zrmz2iedWVeaKKk",
        authDomain: "helpful-rope-333907.firebaseapp.com",
        databaseURL: "https://helpful-rope-333907-default-rtdb.firebaseio.com",
        projectId: "helpful-rope-333907",
        storageBucket: "helpful-rope-333907.appspot.com",
        messagingSenderId: "854982010261",
        appId: "1:854982010261:web:586875587b2a81f679829c",
        measurementId: "G-8XEM0H16H0"
      };

      // Initialize Firebase
      const app = initializeApp(firebaseConfig);
      const storage = getStorage(app);

      // Get elements
      const fileInput = document.getElementById('fileInput');
      const uploadButton = document.getElementById('uploadButton');
      const uploadProgress = document.getElementById('uploadProgress');

      uploadButton.addEventListener('click', () => {
         const file = fileInput.files[0];
         if (file) {
            uploadFile(file);
         } else {
            alert('Please choose a file first.');
         }
      });

      function uploadFile(file) {
         const storageRef = ref(storage, 'hasilnya/' + file.name);
         const uploadTask = uploadBytesResumable(storageRef, file);

         uploadTask.on('state_changed', 
            (snapshot) => {
                  // Observe state change events such as progress, pause, and resume
                  const progress = (snapshot.bytesTransferred / snapshot.totalBytes) * 100;
                  uploadProgress.value = progress;
                  console.log('Upload is ' + progress + '% done');
            }, 
            (error) => {
                  // Handle unsuccessful uploads
                  console.error('Upload failed:', error);
            }, 
            () => {
                  // Handle successful uploads on complete
                  getDownloadURL(uploadTask.snapshot.ref).then((downloadURL) => {
                     console.log('File available at', downloadURL);
                  });
            }
         );
      }

    </script>
</body>
</html>

        '''
        st.components.v1.html(tulisan_html1, height=400)
    with bagian1[1]:
        st.button("Tampilkan AI",on_click=nyalakan)
    with bagian1[0]:
        st.markdown('''
        <div style="font-family:'snap itc';font-size:20px">Pendahuluan</div>
        ''',unsafe_allow_html=True)
        st.image("https://res.cloudinary.com/ikip-siliwangi/image/upload/v1741915021/rancangan2_ss1lkh.jpg")
        st.write("Tratner, dkk. (2022)")
        with st.container(border=True):
            st.markdown('''
        <div style="font-family:'snap itc';font-size:16px">penjelasan</div>
        ''',unsafe_allow_html=True)
            tulisan = '''
        <!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Penjelasan Diagram Media Pembelajaran AI</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            margin: 20px;
            padding: 20px;
        }
        h2 {
            color: #333;
        }
        .section {
            margin-bottom: 20px;
        }
        ul {
            margin-left: 20px;
        }
    </style>
</head>
<body style="background-color:orange">
    <h3>Penjelasan Diagram Media Pembelajaran AI</h3>

    <div class="section">
        <h4>1. Memahami Pengalaman Media</h4>
        <p>Dimulai dengan <strong>analisis penyelesaian soal menggunakan AI</strong>, yang membantu dalam memahami bagaimana AI dapat menyelesaikan soal matematika.</p>
        <ul>
            <li>AI mengumpulkan data dan menyesuaikan dengan kemampuan berpikir pengguna.</li>
            <li>Membantu dalam menghasilkan media pembelajaran berbasis AI.</li>
            <li>Contoh: Hasil coding VBA untuk Excel menurut AI.</li>
        </ul>
    </div>

    <div class="section">
        <h4>2. Pemodelan Pengguna, Personalisasi, dan Keterlibatan</h4>
        <p>Berdasarkan pemahaman terhadap pengalaman media, sistem mulai <strong>memodelkan pengguna</strong>, termasuk aspek personalisasi dan keterlibatan dalam pembelajaran.</p>
        <p>Hasil pemodelan ini dievaluasi sebelum melanjutkan ke tahap produksi konten.</p>
    </div>

    <div class="section">
        <h4>3. Analisis dan Produksi Konten Media</h4>
        <p>Dari hasil pemodelan pengguna, sistem mulai <strong>menganalisis dan memproduksi konten media</strong>.</p>
        <ul>
            <li>Menyesuaikan skenario media pembelajaran.</li>
            <li>Menambahkan coding VBA untuk menyesuaikan rancangan skenario.</li>
            <li>Mengembangkan konten berdasarkan hasil evaluasi pemodelan pengguna.</li>
        </ul>
    </div>

    <div class="section">
        <h4>4. Interaksi dan Aksesibilitas Konten Media</h4>
        <p>Setelah konten diproduksi, langkah selanjutnya adalah <strong>implementasi dan evaluasi oleh pengguna</strong>.</p>
        <p>Reaksi pengguna menjadi faktor penting dalam menyesuaikan aksesibilitas dan efektivitas media pembelajaran.</p>
    </div>

    <div class="section">
        <h4>5. Teknologi Bahasa Alami</h4>
        <p>Hasil akhir yang diharapkan dari seluruh proses ini adalah model pembelajaran berbasis <strong>teknologi bahasa alami</strong>.</p>
        <ul>
            <li>Dengan NLP (Natural Language Processing), media pembelajaran dapat berinteraksi dengan pengguna secara lebih cerdas.</li>
            <li>Menyesuaikan bahasa dengan pemahaman siswa.</li>
        </ul>
    </div>

    <h4>Kesimpulan</h4>
    <p>Diagram ini menunjukkan pendekatan berbasis AI dalam pengembangan media pembelajaran interaktif menggunakan <strong>VBA for Excel</strong>. Setiap langkah membantu membangun sistem pembelajaran yang lebih efektif dan adaptif.</p>
</body>
</html>
        '''
            st.components.v1.html(tulisan,height=1300)
        st.markdown('''
        <div style="font-family:'snap itc';font-size:16px">Referensi</div>
        ''',unsafe_allow_html=True)
        st.write("Trattner, C., Jannach, D., Motta, E., Costera Meijer, I., Diakopoulos, N., Elahi, M., ... & Moe, H. (2022). Responsible media technology and AI: challenges and research directions. AI and Ethics, 2(4), 585-594.")
    with bagian1[3]:
        st.markdown('''
        <div style="font-family:'snap itc';font-size:16px">Pendapat dan Masukan</div>
        ''',unsafe_allow_html=True)
        tulisan1=f'''
            <!DOCTYPE html>
        <html lang="en">
        <head>
            <style>
                #komentar{{
                    border : 2px solid black;
                    border-radius:10px;
                    padding:5px;
                    background-color:yellow;
                }}
            </style>
        </head>
        <body>
            <div id="komentar"><strong><h3 style="color:red; font-family:'snap itc';font-size:18px">Masukan Komentar</h3></strong><br>
                <label for="nama">Nama:</label>
                <input type="text" id="nama" name="nama" value={st.session_state.nama}>
                <label for="nim">NIM: </label>
                <input type="text" id="nim" name="nim" value={st.session_state.NIM}><br><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:15px">Masukan Pendapat Anda Ketika Mempelajari Media ini</h4></strong><br>
                <textarea id="masukan"></textarea><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:16px">Masukan Kesulitan Anda Ketika Mempelajari Media ini</h4></strong><br>
                <textarea id="masuk1"></textarea><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:16px">Masukan deskripsikan keinginan Anda untuk membuat media matematika selanjutnya </h4></strong><br>
                <textarea id="masuk2"></textarea><br>
                <button id="kirim">Kirim</button>
            </div>
            <script type="module">
            // Firebase configuration
            const firebaseConfig = {{
                apiKey: "AIzaSyCS-04HW1WAL3aLJkA7Zrmz2iedWVeaKKk",
                authDomain: "helpful-rope-333907.firebaseapp.com",
                databaseURL: "https://helpful-rope-333907-default-rtdb.firebaseio.com",
                projectId: "helpful-rope-333907",
                storageBucket: "helpful-rope-333907.appspot.com",
                messagingSenderId: "854982010261",
                appId: "1:854982010261:web:586875587b2a81f679829c",
                measurementId: "G-8XEM0H16H0"
            }};

            // Initialize Firebase
            import {{ initializeApp }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-app.js";
            import {{ getDatabase, ref, set  }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-database.js";

            const app = initializeApp(firebaseConfig);
            const db = getDatabase(app);
            var nim = document.getElementById("nim")
            var nama = document.getElementById("nama")
            var masukan = document.getElementById("masukan")
            var masuk1 = document.getElementById("masuk1")
            var masuk2 = document.getElementById("masuk2")
            document.getElementById("kirim").addEventListener("click",()=>{{
                set(ref(db, 'masukan2/' + nim.value), {{
                    Nama: nama.value,
                    Nim: nim.value,
                    Masukan:masukan.value,
                    Masukan1:masuk1.value,
                    Masukan2:masuk2.value
                }})
                .then(()=>{{
                        alert("Sudah Masuk")}})
                .catch((error)=>{{
                    alert(error)
                }})
            }})
            </script>
        
        </body>
        </html>
        '''
        st.components.v1.html(tulisan1,height=800)
        tulisan2 = f'''
    <!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <style>
        #judul{{
            text-align:center;
            font-family:broadway;
            color:green;
            font-size:30px;
        }}
    </style>
</head>
<body>
    <div id="judul"> ANGKET EVALUASI PEMBELAJARAN VBA FOR EXCEL DALAM PENDIDIKAN MATEMATIKA</div>
    <div>
        <div>
            <label for="nama">Nama</label>
            <input type="text" name="nama" id="nama" value={st.session_state['nama']}>
        </div>
        <div>
            <label for="nim">NIM</label>
            <input type="text" name="nim" id="nim" value={st.session_state['NIM']}>
        </div>
    </div>
    <div><strong>Petunjuk pengisian</strong></div>
    <div>Angket ini bertujuan untuk mengetahui pemahaman dan pengalaman Anda setelah mempelajari VBA for Excel dalam pendidikan matematika.<br>
        Berikan penilaian terhadap setiap pernyataan berikut dengan skala:<br>
        <ul>
            <li>1 = Sangat Kurang</li>
            <li>2 = Kurang</li>
            <li>3 = Cukup</li>
            <li>4 = Baik</li>
            <li>5 = Sangat Baik</li>
        </ul>
    <div>
        <ol type="A">
            <li>Pemahaman Dasar VBA for Excel</li>
            <ol>
                <li>Saya memahami konsep dasar VBA for Excel.</li>
                <input type="radio" name="A1" value="1">
                <label for="A1">1</label>
                <input type="radio" name="A1" value="2">
                <label for="A1">2</label>
                <input type="radio" name="A1" value="3">
                <label for="A1">3</label>
                <input type="radio" name="A1" value="4">
                <label for="A1">4</label>
                <input type="radio" name="A1" value="5">
                <label for="A1">5</label>
                <li>Saya dapat menulis dan menjalankan kode VBA sederhana.</li>
                <input type="radio" name="A2" value="1">
                <label for="A2">1</label>
                <input type="radio" name="A2" value="2">
                <label for="A2">2</label>
                <input type="radio" name="A2" value="3">
                <label for="A2">3</label>
                <input type="radio" name="A2" value="4">
                <label for="A2">4</label>
                <input type="radio" name="A2" value="5">
                <label for="A2">5</label>
                <li>Saya dapat membuat massagebox dan inputbox dalam VBA untuk media interaktif.</li>
                <input type="radio" name="A3" value="1">
                <label for="A3">1</label>
                <input type="radio" name="A3" value="2">
                <label for="A3">2</label>
                <input type="radio" name="A3" value="3">
                <label for="A3">3</label>
                <input type="radio" name="A3" value="4">
                <label for="A3">4</label>
                <input type="radio" name="A3" value="5">
                <label for="A3">5</label>
                <li>Saya dapat menghubungkan VBA dengan lembar kerja Excel secara efektif.</li>
                <input type="radio" name="A4" value="1">
                <label for="A4">1</label>
                <input type="radio" name="A4" value="2">
                <label for="A4">2</label>
                <input type="radio" name="A4" value="3">
                <label for="A4">3</label>
                <input type="radio" name="A4" value="4">
                <label for="A4">4</label>
                <input type="radio" name="A4" value="5"">
                <label for="A4">5</label>
            </ol>
            <li>Pembuatan Media Matematika dengan VBA</li>
            <ol>
                <li>Saya dapat membuat media pembelajaran berbasis VBA untuk konsep matematika tertentu.</li>
                <input type="radio" name="B1" value="1">
                <label for="B1">1</label>
                <input type="radio" name="B1" value="2">
                <label for="B1">2</label>
                <input type="radio" name="B1" value="3">
                <label for="B1">3</label>
                <input type="radio" name="B1" value="4">
                <label for="B1">4</label>
                <input type="radio" name="B1" value="5">
                <label for="B1">5</label>
                <li>Saya mampu membuat tombol dan shape interaktif dan grafik dinamis menggunakan VBA.</li>
                <input type="radio" name="B2" value="1">
                <label for="B2">1</label>
                <input type="radio" name="B2" value="2">
                <label for="B2">2</label>
                <input type="radio" name="B2" value="3">
                <label for="B2">3</label>
                <input type="radio" name="B2" value="4">
                <label for="B2">4</label>
                <input type="radio" name="B2" value="5">
                <label for="B2">5</label>
                <li>Saya dapat mengotomatisasi perhitungan matematika menggunakan VBA.</li>
                <input type="radio" name="B3" value="1"">
                <label for="B3">1</label>
                <input type="radio" name="B3" value="2">
                <label for="B3">2</label>
                <input type="radio" name="B3" value="3">
                <label for="B3">3</label>
                <input type="radio" name="B3" value="4">
                <label for="B3">4</label>
                <input type="radio" name="B3" value="5">
                <label for="B3">5</label>
                <li>Saya mampu membuat simulasi atau animasi matematika dengan VBA.</li>
                <input type="radio" name="B4" value="1">
                <label for="B4">1</label>
                <input type="radio" name="B4" value="2">
                <label for="B4">2</label>
                <input type="radio" name="B4" value="3">
                <label for="B4">3</label>
                <input type="radio" name="B4" value="4">
                <label for="B4">4</label>
                <input type="radio" name="B4" value="5">
                <label for="B4">5</label>
            </ol>
            <li>Integrasi VBA dengan Generative AI</li>
            <ol>
                <li>Saya memahami konsep dasar Generative AI.</li>
                <input type="radio" name="C1" value="1">
                <label for="C1">1</label>
                <input type="radio" name="C1" value="2">
                <label for="C1">2</label>
                <input type="radio" name="C1" value="3">
                <label for="C1">3</label>
                <input type="radio" name="C1" value="4">
                <label for="C1">4</label>
                <input type="radio" name="C1" value="5">
                <label for="C1">5</label>
                <li>Saya mengetahui cara mengembangkan Media VBA hasil dari AI seperti OpenAI.</li>
                <input type="radio" name="C2" value="1">
                <label for="C2">1</label>
                <input type="radio" name="C2" value="2">
                <label for="C2">2</label>
                <input type="radio" name="C2" value="3">
                <label for="C2">3</label>
                <input type="radio" name="C2" value="4">
                <label for="C2">4</label>
                <input type="radio" name="C2" value="5">
                <label for="C2">5</label>
                <li>Saya mampu mengembangkan media matematika yang menggunakan AI dari soal otomatis.</li>
                <input type="radio" name="C3" value="1">
                <label for="C3">1</label>
                <input type="radio" name="C3" value="2">
                <label for="C3">2</label>
                <input type="radio" name="C3" value="3">
                <label for="C3">3</label>
                <input type="radio" name="C3" value="4">
                <label for="C3">4</label>
                <input type="radio" name="C3" value="5">
                <label for="C3">5</label>
                <li>Saya dapat memanfaatkan Generative AI untuk memberikan solusi atau penjelasan otomatis dalam media pembelajaran.</li>
                <input type="radio" name="C4" value="1">
                <label for="C4">1</label>
                <input type="radio" name="C4" value="2">
                <label for="C4">2</label>
                <input type="radio" name="C4" value="3">
                <label for="C4">3</label>
                <input type="radio" name="C4" value="4">
                <label for="C4">4</label>
                <input type="radio" name="C4" value="5">
                <label for="C4">5</label>
                <li>Saya dapat mengintegrasikan AI dengan VBA untuk meningkatkan interaktivitas media pembelajaran.</li>
                <input type="radio" name="C5" value="1">
                <label for="C5">1</label>
                <input type="radio" name="C5" value="2">
                <label for="C5">2</label>
                <input type="radio" name="C5" value="3">
                <label for="C5">3</label>
                <input type="radio" name="C5" value="4">
                <label for="C5">4</label>
                <input type="radio" name="C5" value="5">
                <label for="C5">5</label>
            </ol>
            <li>Implementasi dan Evaluasi Media</li>
            <ol>
                <li>Saya dapat menguji keefektifan media pembelajaran yang dibuat dengan VBA dan AI.</li>
                <input type="radio" name="D1" value="1">
                <label for="D1">1</label>
                <input type="radio" name="D1" value="2">
                <label for="D1">2</label>
                <input type="radio" name="D1" value="3">
                <label for="D1">3</label>
                <input type="radio" name="C1" value="4">
                <label for="D1">4</label>
                <input type="radio" name="D1" value="5">
                <label for="D1">5</label>
                <li>Saya mampu memperbaiki dan mengembangkan media berdasarkan evaluasi pengguna.</li>
                <input type="radio" name="D2" value="1">
                <label for="D2">1</label>
                <input type="radio" name="D2" value="2">
                <label for="D2">2</label>
                <input type="radio" name="D2" value="3">
                <label for="D2">3</label>
                <input type="radio" name="D2" value="4">
                <label for="D2">4</label>
                <input type="radio" name="D2" value="5">
                <label for="D2">5</label>
                <li>Saya dapat mengajarkan atau mendemonstrasikan media yang saya buat kepada orang lain.</li>
                <input type="radio" name="D3" value="1">
                <label for="D3">1</label>
                <input type="radio" name="D3" value="2">
                <label for="D3">2</label>
                <input type="radio" name="D3" value="3">
                <label for="D3">3</label>
                <input type="radio" name="D3" value="4">
                <label for="D3">4</label>
                <input type="radio" name="D3" value="5">
                <label for="D3">5</label>
                <li>Saya percaya bahwa penggunaan VBA dan Generative AI dapat meningkatkan efektivitas pembelajaran matematika.</li>
                <input type="radio" name="D4" value="1">
                <label for="C4">1</label>
                <input type="radio" name="D4" value="2">
                <label for="D4">2</label>
                <input type="radio" name="D4" value="3">
                <label for="D4">3</label>
                <input type="radio" name="D4" value="4">
                <label for="D4">4</label>
                <input type="radio" name="D4" value="5">
                <label for="D4">5</label>
            </ol>
        </ol>
    </div>
    </div>
    <button id="kirim">Kirimkan</button>
    <script type="module">
        var A1 = document.getElementsByName("A1")
        var A2 = document.getElementsByName("A2")
        var A3 = document.getElementsByName("A3")
        var A4 = document.getElementsByName("A4")
        var A5 = document.getElementsByName("A5")
        var B1 = document.getElementsByName("B1")
        var B2 = document.getElementsByName("B2")
        var B3 = document.getElementsByName("B3")
        var B4 = document.getElementsByName("B4")
        var B5 = document.getElementsByName("B5")
        var C1 = document.getElementsByName("C1")
        var C2 = document.getElementsByName("C2")
        var C3 = document.getElementsByName("C3")
        var C4 = document.getElementsByName("C4")
        var C5 = document.getElementsByName("C5")
        var D1 = document.getElementsByName("D1")
        var D2 = document.getElementsByName("D2")
        var D3 = document.getElementsByName("D3")
        var D4 = document.getElementsByName("D4")
        var D5 = document.getElementsByName("D5")
        var pilihan = [A1,A2,A3,A4,A5,B1,B2,B3,B4,B5,C1,C2,C3,C4,C5,D1,D2,D3,D4,D5]
        var jawaban = []
        const firebaseConfig = {{
                apiKey: "AIzaSyCS-04HW1WAL3aLJkA7Zrmz2iedWVeaKKk",
                authDomain: "helpful-rope-333907.firebaseapp.com",
                databaseURL: "https://helpful-rope-333907-default-rtdb.firebaseio.com",
                projectId: "helpful-rope-333907",
                storageBucket: "helpful-rope-333907.appspot.com",
                messagingSenderId: "854982010261",
                appId: "1:854982010261:web:586875587b2a81f679829c",
                measurementId: "G-8XEM0H16H0"
            }};

            // Initialize Firebase
            import {{ initializeApp }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-app.js";
            import {{ getDatabase, ref, set  }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-database.js";
            const app = initializeApp(firebaseConfig);
            const db = getDatabase(app);
            var nama = document.getElementById("nama")
            var nim = document.getElementById("nim")
        document.getElementById("kirim").addEventListener("click",()=>{{
            jawaban=[]
            for(var j in pilihan){{
                console.log(j)
                for(var i of pilihan[j]){{
                    if(i.checked){{
                        jawaban.push(i.value)
                    }}
                }}
            }}
            console.log(jawaban)
            set(ref(db, 'pilihan1/' + nim.value), {{
                    Nama: nama.value,
                    Nim: nim.value,
                    pilihan:JSON.stringify(jawaban)
                }})
                .then(()=>{{
                        alert("Sudah Masuk")}})
                .catch((error)=>{{
                    alert(error)
                }})
        }})
    </script>
</body>
</html>
    '''
        st.components.v1.html(tulisan2,height=1400)


if st.sidebar.button("Pelajaran Ketiga"):
    st.session_state.kond = True

if st.session_state.kond:
    daftar = st.tabs(["Pendahuluan","contoh Media","Menggunakan Generative AI", "Penilaian dan Masukan"])
    with daftar[0]:
        st.write("Pendahuluan")
        st.image("https://res.cloudinary.com/ikip-siliwangi/image/upload/v1742192974/diagram1_oo4s86.png")
        tulisan3='''
            <!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Penjelasan Diagram</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; }
        h2 { color: #2c3e50; }
        .section { margin-bottom: 20px; }
        ul { margin-left: 20px; }
    </style>
</head>
<body>
    <h2>Penjelasan Diagram</h2>
    
    <div class="section">
        <h3>1. Input Bahan Pembelajaran</h3>
        <p>Proses dimulai dengan bahan pembelajaran berupa:</p>
        <ul>
            <li><strong>Soal</strong></li>
            <li><strong>Materi</strong></li>
        </ul>
        <p>Bahan ini digunakan untuk mengidentifikasi <strong>Permasalahan dalam Materi Matematika</strong>.</p>
    </div>
    
    <div class="section">
        <h3>2. Proses Pemanfaatan Generative AI</h3>
        <p>Permasalahan dalam materi matematika diproses dengan:</p>
        <ul>
            <li><strong>Generative AI</strong></li>
            <li><strong>Pemanfaatan API</strong></li>
            <li><strong>Streamlit dan Python</strong></li>
        </ul>
        <p>Hasilnya berupa <strong>program dan produk</strong> yang akan digunakan dalam media pembelajaran.</p>
    </div>
    
    <div class="section">
        <h3>3. Hasil Generative AI → Media Pembelajaran Pertama</h3>
        <p>Program dan produk yang dihasilkan digunakan untuk membangun <strong>Media Pembelajaran Pertama</strong>.</p>
        <p>Dalam tahap ini, digunakan:</p>
        <ul>
            <li><strong>Bahasa Program VBA for Excel</strong></li>
            <li><strong>Bentuk Gambar (Shape)</strong></li>
        </ul>
    </div>
    
    <div class="section">
        <h3>4. Evaluasi dan Pengembangan Media Selanjutnya</h3>
        <p>Media Pembelajaran Pertama dievaluasi, kemudian dikembangkan menjadi:</p>
        <ul>
            <li><strong>Media Pembelajaran Kedua</strong></li>
        </ul>
        <p>Analisis dilakukan terhadap berbagai aspek termasuk bentuk gambar.</p>
    </div>
    
    <div class="section">
        <h3>5. Penilaian dan Pengembangan Lebih Lanjut</h3>
        <p>Setelah evaluasi, dilakukan:</p>
        <ul>
            <li><strong>Penilaian dan Masukan</strong></li>
            <li>Pengembangan menjadi <strong>Media Pembelajaran Ketiga</strong></li>
        </ul>
        <p>Proses ini memastikan media pembelajaran menjadi lebih efektif.</p>
    </div>
    
    <h3>Kesimpulan</h3>
    <p>Diagram ini menggambarkan proses iteratif pengembangan media pembelajaran berbasis teknologi, yang mencakup:</p>
    <ul>
        <li>Identifikasi masalah dalam materi matematika</li>
        <li>Penggunaan AI dan pemrograman untuk menghasilkan solusi</li>
        <li>Evaluasi dan pengembangan media pembelajaran secara berkelanjutan</li>
    </ul>
    <p>Tujuan akhirnya adalah menciptakan <strong>media pembelajaran yang lebih efektif</strong> melalui siklus analisis, penilaian, dan perbaikan terus-menerus.</p>
</body>
</html>

    '''
        st.components.v1.html(tulisan3,height=1000)
    with daftar[1]:
        st.header("contoh media")
        st.subheader("Media pembelajaran yang dipilih")
        with st.container(border=True):
            st.write("Buatkan media pembelajaran matematika menggunakan shape di VBA for Excel")
        st.markdown('''
        <div style="font-family:'snap itc';font-size:16px">Hasil Generative AI</div>
        ''',unsafe_allow_html=True)
        st.code('''
            Sub PeluangKoin()
' Deklarasi variabel
  Dim sisiKoin As Integer, hasilLemparan As Integer
  Dim peluangKepala As Double, peluangEkor As Double
  Dim i As Integer, jumlahLemparan As Integer

  ' Input jumlah lemparan
  jumlahLemparan = InputBox("Masukkan jumlah lemparan koin:", "Simulasi Lemparan Koin", 100)
  If jumlahLemparan < 1 Then Exit Sub ' Keluar jika input tidak valid


  ' Buat Shapes untuk representasi koin
  With ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 50, 50, 50)
    .Name = "Koin"
    .Fill.ForeColor.RGB = RGB(255, 255, 0) ' Kuning
    .Line.Visible = msoFalse
  End With

  ' Buat Shapes untuk representasi hasil
  With ActiveSheet.Shapes.AddShape(msoShapeRectangle, 200, 50, 100, 50)
    .Name = "Hasil"
    .Fill.ForeColor.RGB = RGB(255, 255, 255) ' Putih
    .TextFrame.Characters.Text = ""
  End With

  ' Simulasi lemparan koin
  For i = 1 To jumlahLemparan

    ' Generate angka acak 0 atau 1 (0 = Kepala, 1 = Ekor)
    sisiKoin = Int(Rnd * 2)

    ' Update tampilan koin (animasi sederhana)
    If sisiKoin = 0 Then
      ActiveSheet.Shapes("Koin").Fill.ForeColor.RGB = RGB(255, 255, 0) ' Kuning (Kepala)
    Else
      ActiveSheet.Shapes("Koin").Fill.ForeColor.RGB = RGB(192, 192, 192) ' Abu-abu (Ekor)
    End If

    DoEvents ' Penting untuk update tampilan

    ' Hitung hasil
    If sisiKoin = 0 Then
      hasilLemparan = hasilLemparan + 1
    End If


    ' Hitung dan tampilkan peluang
    peluangKepala = hasilLemparan / i
    peluangEkor = (i - hasilLemparan) / i

    ActiveSheet.Shapes("Hasil").TextFrame.Characters.Text = "Kepala: " & Format(peluangKepala, "0.0%") & vbCrLf & "Ekor: " & Format(peluangEkor, "0.0%")

    ' Jeda sebentar untuk melihat animasi (opsional)
    Application.Wait Now + TimeValue("0:00:01") ' Jeda 1 detik

  Next i

  MsgBox "Simulasi selesai!", vbInformation

End Sub
        ''',language="vb")
        st.markdown('''
        <div style="font-family:'snap itc';font-size:16px">Penjelasan Analisis Koding</div>
        ''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                Dim sisiKoin As Integer, hasilLemparan As Integer
                Dim peluangKepala As Double, peluangEkor As Double
                Dim i As Integer, jumlahLemparan As Integer
            ''',language="vb")
        st.markdown('''
            <div style="font-family:'comic sans ms' font-size:16px">
                    <ol>
                        <li>Deklarasi sisiKoin, hasilLemparan sebagai Integer</li>
                        <li>peluangKepala, peluangEkor sebagai Double</li>
                        <li>i dan jumlahLemparan sebagai Integer</li>
                    </ol>
                    <div>Integer adalah bilangan bulat positif sedangkan Double bilangan positif desimal</div>
            </div>
        ''', unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                jumlahLemparan = InputBox("Masukkan jumlah lemparan koin:", "Simulasi Lemparan Koin", 100)
                If jumlahLemparan < 1 Then Exit Sub ' Keluar jika input tidak valid
            ''',language="vb")
        st.markdown('''
            <div style="font-family:'comic sans ms' font-size:16px">
                    Ketika program dijalankan Kotak masukan akan ditampilkan dan berisi 100 dan nilai tersebut
                    tersimpan di jumlahLemparan. Jika jumlahLemparan kurang dari 1 maka keluar dari fungsi sub 
                    peluangKoin atau program berhenti bekerja.
            </div>
        ''', unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                ' Buat Shapes untuk representasi koin
                With ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 50, 50, 50)
                    .Name = "Koin"
                    .Fill.ForeColor.RGB = RGB(255, 255, 0) ' Kuning
                    .Line.Visible = msoFalse
                End With
            ''',language="vb")
        st.markdown('''
            <div style="font-family:'comic sans ms' font-size:16px">
                    Buat gambar oval (lingkaran) yang posisinya x=100px dan y=50px,
                    tinggi gambar 50px dan tinggi gambar 50px. Gambar tersebut
                    diberi nama Koin. Warna backgroundnya kuning. Dan sisi gambar
                    dihilangkan.
            </div>
        ''', unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                ' Buat Shapes untuk representasi hasil
                With ActiveSheet.Shapes.AddShape(msoShapeRectangle, 200, 50, 100, 50)
                    .Name = "Hasil"
                    .Fill.ForeColor.RGB = RGB(255, 255, 255) ' Putih
                    .TextFrame.Characters.Text = ""
                End With
            ''',language="vb")
        st.markdown('''
            <div style="font-family:'comic sans ms' font-size:16px">
                    Buat gambar persegi panjang yang posisinya x=200px dan y=50px,
                    tinggi gambar 100px dan tinggi gambar 50px. Gambar tersebut
                    diberi nama Hasil. Warna backgroundnya potih. Tulisan dikosongkan.
            </div>
        ''', unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                ' Simulasi lemparan koin
  For i = 1 To jumlahLemparan

    ' Generate angka acak 0 atau 1 (0 = Kepala, 1 = Ekor)
    sisiKoin = Int(Rnd * 2)

    ' Update tampilan koin (animasi sederhana)
    If sisiKoin = 0 Then
      ActiveSheet.Shapes("Koin").Fill.ForeColor.RGB = RGB(255, 255, 0) ' Kuning (Kepala)
    Else
      ActiveSheet.Shapes("Koin").Fill.ForeColor.RGB = RGB(192, 192, 192) ' Abu-abu (Ekor)
    End If

    DoEvents ' Penting untuk update tampilan

    ' Hitung hasil
    If sisiKoin = 0 Then
      hasilLemparan = hasilLemparan + 1
    End If


    ' Hitung dan tampilkan peluang
    peluangKepala = hasilLemparan / i
    peluangEkor = (i - hasilLemparan) / i

    ActiveSheet.Shapes("Hasil").TextFrame.Characters.Text = "Kepala: " & Format(peluangKepala, "0.0%") & vbCrLf & "Ekor: " & Format(peluangEkor, "0.0%")

    ' Jeda sebentar untuk melihat animasi (opsional)
    Application.Wait Now + TimeValue("0:00:01") ' Jeda 1 detik

  Next i
            ''',language="vb")
        st.markdown('''
            <div style="font-family:'comic sans ms' font-size:16px; text-align:justify">
                    Kemudian program akan berjalan secara berulang 1 sampai 100 kali
                    (sesuai dengan jumlahLemparan). SisiKoin akan bernilai 0 sampai 1 secara acak.
                    Jika sisiKoin sama dengan 0 maka warna background menjadi kuning.
                    Jika sisiKoin sama dengan 1 maka warna background menjadi Abu-abu.
                    Jika sisiKoin sama dengan 0 maka hasilLemparan akan bertambah 1.
                    PeluangKepala akan terhitung sama dengan hasilLemparan dibagi i.
                    PeluangEkor akan terhitung sama dengan (i-hasilLemparan) dibagi i.
                    Program pengulangan akan berjalan setiap 1 detik akan terupdate.
            </div>
        ''', unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
                MsgBox "Simulasi selesai!", vbInformation
            ''',language="vb")
        st.markdown('''
            <div style="font-family:'comic sans ms' font-size:16px; text-align:justify">
                Kotak keterangan akan muncul dan menjelaskan bahwa permainan sudah selesai.
            </div>
        ''', unsafe_allow_html=True)
        st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1742216535/Recording_52_kxcqi5.mp4")
        st.markdown('''
        <div style="font-family:'snap itc';font-size:16px; color:blue">Evaluasi Media Kedua</div>
        ''',unsafe_allow_html=True)
        st.markdown('''
            <div style="font-family:'comic sans ms' font-size:16px; text-align:justify">
                <ol>
                    <li>Pada pertama dijalankan gambar lingkaran berjalan dengan baik
                        Namun menjalankan program kembali gambar lingkaran tidak berjalan
                        ini dikarenakan lingkaran yang bernama koin terjadi duplikat, 
                        disebabkan fungsi AddShape (membuat gambar secara otomatis)</li>
                    <li> Demikian juga pada persegi panjang yang bernama Hasil 
                        Pada pertama dijalankan gambar persegi panjang berjalan dengan baik
                        Namun menjalankan program kembali gambar persegi panjang tidak berjalan
                        ini dikarenakan persegi panjang yang bernama Hasil terjadi duplikat
                        disebabkan fungsi AddShape (membuat gambar secara otomatis)</li>
                    <li>Tulisan di dalam persegi panjang tidak ditampilkan karena warna huruf
                        dan warna background warnanya sama yaitu putih.</li>
                </ol>
            </div>
        ''', unsafe_allow_html=True)
        st.markdown('''
        <div style="font-family:'snap itc';font-size:16px; color:blue">Perbaikan Media Kedua</div>
        ''',unsafe_allow_html=True)
        st.markdown('''
            <div style="font-family:'comic sans ms' font-size:16px; text-align:justify">
                <ol>
                    <li>Untuk mengantisipasi dari poin 1 dan 2 bagian Evaluasi Media Kedua
                        maka diperlukan fungsi tambahan untuk menghapus duplikat (contoh: 
                        Activesheet.Shapes("Koin").Delete), Fungsi tersebut harus di atas
                        fungsi AddShape. Namun perlu ditambahkan <i>on Error esume Next</i>
                        tempaykan bagian paling atas tapi di bawah sub peluangKoin</li>
                    <li>Tulisan di dalam persegi panjang tidak ditampilkan karena warna huruf
                        ditambahkan perintah fungsi contoh: ActiveSheet.Shapes("Hasil").textFrame.
                        characters.Font.Color=vbBlack.</li>
                </ol>
            </div>
        ''', unsafe_allow_html=True)
        st.markdown('''
        <div style="font-family:'snap itc';font-size:16px; color:blue">Hasil Media Kedua</div>
        ''',unsafe_allow_html=True)
        with st.container(border=True):
            st.code('''
            Sub PeluangKoin()

On Error Resume Next 'Tambahan fungsi untuk mengabaikan terjadi error

' Deklarasi variabel
  Dim sisiKoin As Integer, hasilLemparan As Integer
  Dim peluangKepala As Double, peluangEkor As Double
  Dim i As Integer, jumlahLemparan As Integer

  ' Input jumlah lemparan
  jumlahLemparan = InputBox("Masukkan jumlah lemparan koin:", "Simulasi Lemparan Koin", 100)
  If jumlahLemparan < 1 Then Exit Sub ' Keluar jika input tidak valid

ActiveSheet.Shapes("Koin").Delete 'Menghapus gambar yang bernama Koin
ActiveSheet.Shapes("Hasil").Delete 'Menghapus gambar yang bernama Hasil

  ' Buat Shapes untuk representasi koin
  With ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 50, 50, 50)
    .Name = "Koin"
    .Fill.ForeColor.RGB = RGB(255, 255, 0) ' Kuning
    .Line.Visible = msoFalse
  End With

  ' Buat Shapes untuk representasi hasil
  With ActiveSheet.Shapes.AddShape(msoShapeRectangle, 200, 50, 100, 50)
    .Name = "Hasil"
    .Fill.ForeColor.RGB = RGB(255, 255, 255) ' Putih
    .TextFrame.Characters.Text = ""
    .TextFrame.Characters.Font.Color = vbBlack 'Memberikan warna huruf hitam
  End With

  ' Simulasi lemparan koin
  For i = 1 To jumlahLemparan

    ' Generate angka acak 0 atau 1 (0 = Kepala, 1 = Ekor)
    sisiKoin = Int(Rnd * 2)

    ' Update tampilan koin (animasi sederhana)
    If sisiKoin = 0 Then
      ActiveSheet.Shapes("Koin").Fill.ForeColor.RGB = RGB(255, 255, 0) ' Kuning (Kepala)
    Else
      ActiveSheet.Shapes("Koin").Fill.ForeColor.RGB = RGB(192, 192, 192) ' Abu-abu (Ekor)
    End If

    DoEvents ' Penting untuk update tampilan

    ' Hitung hasil
    If sisiKoin = 0 Then
      hasilLemparan = hasilLemparan + 1
    End If


    ' Hitung dan tampilkan peluang
    peluangKepala = hasilLemparan / i
    peluangEkor = (i - hasilLemparan) / i

    ActiveSheet.Shapes("Hasil").TextFrame.Characters.Text = "Kepala: " & Format(peluangKepala, "0.0%") & vbCrLf & "Ekor: " & Format(peluangEkor, "0.0%")

    ' Jeda sebentar untuk melihat animasi (opsional)
    Application.Wait Now + TimeValue("0:00:01") ' Jeda 1 detik

  Next i

  MsgBox "Simulasi selesai!", vbInformation
End Sub
        ''',language="vb")
        st.write("Video Evaluasi Media kedua")
        st.video("https://res.cloudinary.com/ikip-siliwangi/video/upload/v1742220594/Recording_53_ksaak0.mp4")
    with daftar[2]:
        st.write("Hasil Generative AI dan pengembangannya ")
    with daftar[3]:
        tulisan4 = f'''
            <!DOCTYPE html>
        <html lang="en">
        <head>
            <style>
                #komentar{{
                    border : 2px solid black;
                    border-radius:10px;
                    padding:5px;
                    background-color:yellow;
                }}
            </style>
        </head>
        <body>
            <div id="komentar"><strong><h3 style="color:red; font-family:'snap itc';font-size:18px">Masukan Komentar</h3></strong><br>
                <label for="nama">Nama:</label>
                <input type="text" id="nama" name="nama" value={st.session_state.nama}>
                <label for="nim">NIM: </label>
                <input type="text" id="nim" name="nim" value={st.session_state.NIM}><br><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:15px">Masukan tentang Pemahaman dan Penguasaan Teknis</h4></strong>
                <ol>
                    <li> Sejauh mana Anda merasa memahami dasar-dasar pemrograman VBA for Excel?</li>
                    <li> Apa kesulitan utama yang Anda hadapi saat mengembangkan media matematika dengan VBA?</li>
                    <li> Seberapa nyaman Anda dalam mengintegrasikan Generative AI dengan VBA?</li>
                    <li>Menurut Anda, fitur apa yang perlu ditambahkan untuk meningkatkan efektivitas media pembelajaran berbasis VBA dan AI?</li>
                </ol>
                <textarea id="masukan"></textarea><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:16px">Masukan tentang Efektivitas Media</h4></strong>
                <ol>
                    <li> Bagaimana pendapat Anda tentang manfaat penggunaan Generative AI dalam pembuatan media matematika?</li>
                    <li> Apakah media yang dibuat dapat membantu pemahaman konsep matematika dengan lebih baik? Jika ya, bagaimana? Jika tidak, mengapa?</li>
                    <li> Apakah ada aspek tertentu dalam media yang perlu ditingkatkan agar lebih interaktif dan mudah digunakan?</li>
                </ol>
                <textarea id="masuk1"></textarea><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:16px">Masukan tentang Implementasi dan Penggunaan</h4></strong>
                <ol>
                    <li> Bagaimana tingkat kesulitan dalam menerapkan media pembelajaran ini dalam proses belajar mengajar?</li>
                    <li> Menurut Anda, apakah media ini dapat digunakan secara luas oleh pengajar dan siswa? Apa yang perlu ditingkatkan agar lebih mudah diadopsi?</li>
                    <li> Apa saran Anda untuk pengembangan lebih lanjut agar media ini lebih efektif dan bermanfaat?</li>
                </ol>
                <textarea id="masuk2"></textarea><br>
                <strong><h4 style="color:black; font-family:'snap itc';font-size:16px">Koding apa yang Anda mau pelajari lebih dalam</h4></strong>
                <textarea id="masuk3"></textarea><br>
                <button id="kirim">Kirim</button>
            </div>
            <script type="module">
            // Firebase configuration
            const firebaseConfig = {{
                apiKey: "AIzaSyCS-04HW1WAL3aLJkA7Zrmz2iedWVeaKKk",
                authDomain: "helpful-rope-333907.firebaseapp.com",
                databaseURL: "https://helpful-rope-333907-default-rtdb.firebaseio.com",
                projectId: "helpful-rope-333907",
                storageBucket: "helpful-rope-333907.appspot.com",
                messagingSenderId: "854982010261",
                appId: "1:854982010261:web:586875587b2a81f679829c",
                measurementId: "G-8XEM0H16H0"
            }};

            // Initialize Firebase
            import {{ initializeApp }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-app.js";
            import {{ getDatabase, ref, set  }} from "https://www.gstatic.com/firebasejs/11.4.0/firebase-database.js";

            const app = initializeApp(firebaseConfig);
            const db = getDatabase(app);
            var nim = document.getElementById("nim")
            var nama = document.getElementById("nama")
            var masukan = document.getElementById("masukan")
            var masuk1 = document.getElementById("masuk1")
            var masuk2 = document.getElementById("masuk2")
            var masuk3 = document.getElementById("masuk3")
            document.getElementById("kirim").addEventListener("click",()=>{{
                set(ref(db, 'masukan3/' + nim.value), {{
                    Nama: nama.value,
                    Nim: nim.value,
                    Masukan:masukan.value,
                    Masukan1:masuk1.value,
                    Masukan2:masuk2.value,
                    Masukan3:masuk3.value
                }})
                .then(()=>{{
                        alert("Sudah Masuk")}})
                .catch((error)=>{{
                    alert(error)
                }})
            }})
            </script>
        
        </body>
        </html>

        '''
        st.components.v1.html(tulisan4,height=1600)
if st.session_state.kond:
    st.markdown('''
        <div style="font-family:'snap itc';font-size:16px;color:red">Menggunakan AI Generative</div>
        ''',unsafe_allow_html=True)
    for i in st.session_state.kumpulan:
        with st.chat_message(i["role"]):
            st.markdown(i["content"])
        st.write("Masukan perintah")
        
if st.session_state.kond:
    masukan = st.chat_input("Masukan perintah")
    if masukan:
        try:
            st.session_state.kumpulan.append({"role":"user","content":masukan})
            with st.chat_message("user"):
                st.markdown(masukan)
                response = model.generate_content(masukan)
            st.session_state.kumpulan.append({"role":"assistant","content":response.text})
            with st.chat_message("assistant"):
                st.markdown(response.text)
        except TypeError:
            st.write("Masukan Perintah")
if st.session_state['genai']:
    st.session_state.isian = st.text_input("Masukan permintaan", value = st.session_state.isian)
    if st.session_state.isian:
        # Buat prompt dan kirim ke model
        response = model.generate_content(st.session_state.isian)
        st.write(response.text)



