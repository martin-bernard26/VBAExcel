import streamlit as st
import google.generativeai as genai
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
            Media Visual Basic Application for Excel
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
if st.sidebar.button("Pelajaran Kedua"):
    genai.configure(api_key="AIzaSyDdh2-QQc15-gLC4T3WJz4eatMc21bv0iA")
    model = genai.GenerativeModel("gemini-1.5-pro")

    st.title("🤖 Chatbot Gemini AI")

    if "chat" not in st.session_state:
        st.session_state.chat = model.start_chat()

    user_input = st.text_input("Ketik pesan...")

    if st.button("Kirim") and user_input:
        response = st.session_state.chat.send_message(user_input)
        st.write("🧠 **Gemini AI:**", response.text)

