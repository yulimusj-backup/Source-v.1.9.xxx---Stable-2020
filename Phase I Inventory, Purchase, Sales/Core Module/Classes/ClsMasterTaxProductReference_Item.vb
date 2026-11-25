Option Explicit On
Option Strict Off

Public Class ClsMasterTaxProductReference_Item
    Private Function GetUrutan2() As String
        Dim retval As String = ""
        Try
            retval = "Insert Into `tblmaster_taxproduct_reference` (`type`,`Code`,`Description`) Values" & _
"('0','391500','Sis reja dan skrap, dari plastik.')," & _
"('0','391600','Monofilamen yang ukuran setiap penampang silangnya melebihi 1 mm, berbentuk batang kecil, stik dan profil, yang permukaannya dikerjakan maupun tidak tetapi tidak dikerjakan secara lain, dari plastik.')," & _
"('0','391700','Pembuluh, pipa dan selang, serta alat kelengkapan dari padanya (misalny penyambung pip elbow, flensa), dari plastik.')," & _
"('0','391800','Penutup lantai dari plastik, dengan perekat maupun tidak, dalam gulungan atau dalam bentuk ubin penutup dinding atau langit-langit dari plastik, sebagaimana dinyatakan dalam Catatan 9 pada Bab ini.')," & _
"('0','391900','Pelat, lembaran, film, foil, pit strip dan bentuk pipih lainnya berperekat, dari plastik, dalam gulungan maupun tidak.')," & _
"('0','392000','Pelat, lembaran, film, foil dan strip lainny dari plastik, non seluler dan tidak diperkuat, tidak dilaminasi, tidak didukung atau tidak dikombinasi dengan cara semacam itu dengan bahan lain.')," & _
"('0','392100','Pelat, lembaran, film, foil dan strip lainnya dari plastik.')," & _
"('0','392200','Bak mandi, bak mandi dengan pancuran, bak cuci, wastafel, bidet, mangkuk, tempat duduk dan penutup kloset, tangki air penyiram dan perlengkapan saniter semacam itu, dari plastik.')," & _
"('0','392300','Wadah untuk mengangkut atau mengemas barang, dari plastik sumbat, tutup, tutup botol dan penutup lainny dari plastik.')," & _
"('0','392400','Perangkat makan, perangkat dapur, peralatan rumah tangga lainnya dan peralatan toilet atau higienis, dari plastik.')," & _
"('0','392500','Perangkat bangunan dari plastik, tidak dirinci atau termasuk dalam pos lainnya.')," & _
"('0','392600','Barang lain dari plastik dan barang dari bahan lain yang dimaksud dalam pos 39.01 sampai dengan 39.14.')," & _
"('0','400000','Karet dan barang daripadanya')," & _
"('0','400100','Karet alam, balat getah perc guayule, chicle dan getah alam semacam itu, dalam bentuk asal atau pelat, lembaran atau strip.')," & _
"('0','400200','Karet sintetik dan factice diperoleh dari minyak, dalam bentuk asal atau pelat, lembaran atau strip campuran produk dari pos 40.01 dengan produk apapun dari pos ini, dalam bentuk asal atau pelat, lembaran atau strip.')," & _
"('0','400300','Reclaimed rubber dalam bentuk asal atau pelat, lembaran atau strip.')," & _
"('0','400400','Sis reja dan skrap karet (selain karet keras) dan bubuk serta butir yang diperoleh daripadanya.')," & _
"('0','400500','Karet campuran, tidak divulkanisasi, dalam bentuk asal atau pelat, lembaran atau strip.')," & _
"('0','400600','Bentuk lainnya (misalny bentuk batang kecil, pembuluh dan profil) dan barang (misalny cakram dan cincin), dari karet tidak divulkanisasi.')," & _
"('0','400700','Benang dan tali karet divulkanisasi.')," & _
"('0','400800','Bentuk pelat, lembaran, strip, batang kecil dan profil, dari karet divulkanisasi selain karet keras.')," & _
"('0','400900','Pembuluh, pipa dan selang dari karet divulkanisasi selain karet keras, dengan atau tanpa alat kelengkapannya (misalny penghubung, siku, flensa).')," & _
"('0','401000','Ban atau belting pengangkut atau penggerak dari karet divulkanisasi.')," & _
"('0','401100','Ban bertekanan, baru, dari karet.')," & _
"('0','401200','Ban bertekanan, bekas atau ditelapaki ulang, dari karet ban padat atau bantalan, telapak ban dan penutup ban dari karet.')," & _
"('0','401300','Ban dalam, dari karet.')," & _
"('0','401400','Barang higienis atau farmasi (termasuk dot), dari karet divulkanisasi selain karet keras, dengan atau tanpa alat kelengkapan dari karet keras.')," & _
"('0','401500','Pakaian dan aksesori pakaian (termasuk sarung tangan, mitten dan mitt), untuk segala keperluan, dari karet divulkanisasi selain karet keras.')," & _
"('0','401600','Barang lain dari karet divulkanisasi selain karet keras.')," & _
"('0','401700','Karet keras (misalny ebonit) dalam segala bentuk, termasuk sisa dan skrap barang dari karet keras.')," & _
"('0','410000','Jangat dan kulit mentah (selain kulit berbulu dan kulit samak')," & _
"('0','410100','Hide dan skin mentah dari hewan jenis lembu (termasuk kerbau) atau hewan jenis kuda (segar, atau diasinkan, dikeringkan, dikapur, diasamkan atau diawetkan secara lain, tetapi tidak disamak, tidak diolah secara perkamen atau tidak diolah lebih lanjut), dihilangkan bulunya atau split maupun tidak.')," & _
"('0','410200','Skin mentah dari biri - biri atau biri - biri muda (segar atau diasinkan, dikeringkan, dikapur, diasamkan atau diawetkan secara lain tetapi tidak disamak, tidak diolah secara perkamen atau tidak diolah lebih lanjut), dengan wol atau split maupun tidak, selain yang dikecualikan oleh Catatan 1 (c) pada Bab ini.')," & _
"('0','410300','Hide dan skin mentah lainnya (segar atau diasinkan, dikeringkan, dikapur, diasamkan atau diawetkan secara lain, tetapi tidak disamak, tidak diolah secara perkamen atau tidak diolah lebih lanjut), dihilangkan bulunya atau split maupun tidak, selain yang dikecualikan oleh Catatan 1 (b) atau Catatan 1 (c) dalam Bab ini.')," & _
"('0','410400','Hide dan skin dari hewan jenis lembu (termasuk kerbau) atau hewan jenis kuda yang disamak atau crust, tanpa bulu, split maupun tidak, tetapi tidak diolah lebih lanjut.')," & _
"('0','410500','Skin dari biri-biri atau biri-biri muda disamak atau crust, tanpa wol, split maupun tidak, tetapi tidak diolah lebih lanjut.')," & _
"('0','410600','Hide dan skin dari hewan lainnya disamak atau crust, tanpa wol atau bulu, split maupun tidak, tetapi tidak diolah lebih lanjut.')," & _
"('0','410700','Kulit samak yang diolah lebih lanjut setelah penyamakan atau crusting, termasuk kulit samak yang diolah secara perkamen, dari hewan jenis lembu (termasuk kerbau) atau hewan jenis kud tanpa bulu, split maupun tidak, selain kulit samak dari pos 41.14.')," & _
"('0','411200','Kulit samak yang diolah lebih lanjut setelah penyamakan atau crusting, termasuk kulit samak yang diolah secara perkamen dari biri-biri atau biri-biri mud tanpa wol, split maupun tidak, selain kulit samak dari pos 41.14.')," & _
"('0','411300','Kulit samak yang diolah lebih lanjut setelah penyamakan atau crusting, termasuk kulit samak yang diolah secara perkamen dari hewan lainny tanpa wol atau bulu, split maupun tidak, selain kulit samak dari pos 41.14.')," & _
"('0','411400','Kulit chamois (termasuk kulit chamois kombinasi) kulit paten dan kulit paten dilaminasi, kulit samak dilapisi logam.')," & _
"('0','411500','Kulit komposisi dengan bahan dasar kulit samak atau serat kulit samak, dalam lembaran tebal, lembaran atau strip, dalam gulungan maupun tidak reja dan sisa lainnya dari kulit samak atau dari kulit komposisi, tidak cocok untuk pembuatan barang dari kulit samak debu, bubuk dan tepung kulit samak.')," & _
"('0','420000','Barang dari kulit samak saddlery dan harness barang untuk bepergian, tas tangan dan kemasan semacam itu barang dari usus binatang (selain benang ulat sutra)')," & _
"('0','420100','Saddlery dan harness untuk semua macam binatang (termasuk tali kekang, kekang, penutup lutut, penutup mulut, tutup sadel, tas sadel, jaket anjing dan sejenisnya), dari berbagai bahan.')," & _
"('0','420200','Peti, koper, vanity-case, tas eksekutif, tas kantor, tas sekolah, dompet kacamat tas teropong, tas kamer tas peralatan musik, koper senjat sarung pistol dan kemasan semacam itu tas untuk bepergian, tas makanan dan minuman bersekat, tas rias, ransel, tas tangan, tas belanj dompet, pundi, tempat pet tempat rokok, kantong tembakau, tas perkakas, tas olahrag tempat botol, kotak perhiasan, kotak bedak, tempat pisau dan kemasan semacam itu dari kulit samak atau dari kulit komposisi, dengan lembaran dari plastik, atau dari bahan tekstil, atau dari serat vulkanisasi atau dari kertas karton seluruhnya atau sebagian besar dibungkus bahan tersebut atau dengan kertas.')," & _
"('0','420300','Pakaian dan aksesori pakaian, dari kulit samak atau dari kulit komposisi.')," & _
"('0','420500','Barang lainnya dari kulit samak atau dari kulit komposisi.')," & _
"('0','420600','Barang dari usus (selain benang ulat sutra), dari kulit goldbeater, dari kandung kemih atau dari urat.')," & _
"('0','430000','Kulit berbulu dan bulu artifisial barang daripadanya')," & _
"('0','430100','Kulit berbulu mentah (termasuk kepal ekor, cakar dan bagian atau potongan lainny cocok digunakan untuk pakaian kulit berbulu), selain hides dan skins mentah dari pos 41.01, 41.02 atau 41.03.')," & _
"('0','430200','Kulit berbulu disamak atau dikerjakan (termasuk kepal ekor, cakar dan bagian atau potongan lainnya), dipisah, atau disatukan (tanpa penambahan bahan lainnya) selain yang dimaksud dalam pos 43.03.')," & _
"('0','430300','Pakaian, aksesori pakaian dan barang lainnya dari kulit berbulu.')," & _
"('0','430400','Bulu artifisial dan barang terbuat dari padanya.')," & _
"('0','440000','Kayu dan barang dari kayu arang kayu')," & _
"('0','440100','Kayu bakar, berbentuk log, billet, ranting, ikatan cabang atau bentuk semacam itu kayu dalam bentuk keping atau pecahan kayu serbuk gergaji dan sisa serta skrap kayu, diaglomerasi berbentuk log, briket, pelet atau bentuk semacam itu maupun tidak diaglomerasi')," & _
"('0','440200','Arang kayu (termasuk arang dari tempurung atau arang dari buah bertempurung), diaglomerasi maupun tidak.')," & _
"('0','440300','Kayu kasar, dihilangkan kulit atau kayu gubalnya maupun tidak, atau dibentuk bujur sangkar secara kasar.')," & _
"('0','440400','Kayu simpai galah belahan piles, tiang pancang dan tonggak dari kayu, runcing tetapi tidak digergaji memanjang tongkat kayu, dipotong secara kasar tetapi tidak dibubut, dibengkokkan atau dikerjakan secara lain, cocok untuk pembuatan tongkat jalan, payung, gagang perkakas atau sejenisnya kepingan kayu dan sejenisnya.')," & _
"('0','440500','Wol kayu tepung kayu.')," & _
"('0','440600','Bantalan (cross-tie) rel kereta api atau trem dari kayu.')," & _
"('0','440700','Kayu digergaji atau dibelah memanjang, diiris atau dikuliti, diketam, diampelas atau end-jointed maupun tidak, dengan ketebalan melebihi 6 mm.')," & _
"('0','440800','Lembaran untuk veneering (termasuk yang diperoleh dengan cara mengiris kayu yang dilaminasi), untuk kayu lapis atau kayu yang dilaminasi semacam itu dan kayu lainny digergaji memanjang, diiris atau dikuliti, diketam, diampelas, disambung atau end-jointed maupun tidak, dengan ketebalan tidak melebihi 6 mm.')," & _
"('0','440900','Kayu (termasuk strip dan frieze untuk lantai papan, tidak dipasang) dibentuk tidak terputus (diberi lidah, diberi alur, tepinya dikorok, diberi lereng, V-jointed, beaded, diberi pola bentukan, dibundarkan atau sejenis itu), sepanjang tepi, ujung atau permukaanny diketam, diampelas atau end-jointed maupun tidak.')," & _
"('0','441000','Papan partikel, papan oriented strand (OSB) dan papan semacam itu (misalny papan wafer) dari kayu atau bahan mengandung lignin lainny diaglomerasi dengan resin atau dengan zat pengikat organik lainnya maupun tidak.')," & _
"('0','441100','Papan fiber dari kayu atau bahan mengandung lignin lainny direkatkan dengan resin atau zat organik lainnya maupun tidak.')," & _
"('0','441200','Kayu lapis, panel veneer dan kayu dilaminasi semacam itu.')," & _
"('0','441300','Kayu dipadatkan, berbentuk blok, pelat, strip atau profil.')," & _
"('0','441400','Bingkai kayu untuk lukisan, foto, cermin atau benda semacam itu.')," & _
"('0','441500','Peti, kotak, krat, drum dan pengemas semacam itu, dari kayu gelendong kabel dari kayu palet, palet kotak dan papan untuk muatan lainny dari kayu kerah palet dari kayu.')," & _
"('0','441600','Tahang, tong, bejan pasu dan produk lainnya dari pembuat tong/pasu dan bagianny dari kayu, termasuk stave.')," & _
"('0','441700','Perkakas, badan perkakas, gagang perkakas, badan dan gagang sapu atau sikat dan gagangnya dari kayu kelebut bot atau sepatu dan boot tree atau shoe tree, dari kayu.')," & _
"('0','441800','Produk pertukangan dan bahan bangunan rumah dari kayu, termasuk panel kayu seluler, rakitan panel penutup lantai, atap sirap dan shake.')," & _
"('0','441900','Perangkat makan dan perangkat dapur, dari kayu.')," & _
"('0','442000','Tatakan kayu dan kayu bertatah kotak dan peti untuk perhiasan atau peralatan makan dan barang semacam itu, dari kayu patung kecil dan ornamen lainny dari kayu perabotan dari kayu yang tidak termasuk Bab 94.')," & _
"('0','442100','Barang lainnya dari kayu.')," & _
"('0','450000','Gabus dan barang dari gabus')," & _
"('0','450100','Gabus alam, mentah atau diolah sederhana gabus sisa gabus pecahan, butiran atau gilingan.')," & _
"('0','450200','Gabus alam, dikupas atau dibentuk bujur sangkar secara kasar, atau blok empat persegi panjang (termasuk bujur sangkar), pelat, lembaran atau strip, termasuk sharp-edged blanks untuk sumbat atau penutup).')," & _
"('0','450300','Barang dari gabus alam.')," & _
"('0','450400','Gabus diaglomerasi (dengan atau tanpa zat pengikat) dan barang dari gabus diaglomerasi.')," & _
"('0','460000','Barang dari jerami, dari rumput esparto dan dari bahan anyaman lainnya keranjang dan barang anyaman')," & _
"('0','460100','Anyaman dan produk semacam itu dari bahan anyaman, strip maupun tidak bahan anyaman, anyaman dan produk semacam itu dari bahan anyaman, diikat bersama menjadi untaian atau tenunan sejajar, dalam bentuk lembaran, merupakan barang jadi maupun tidak (misalny mat, matting, tirai).')," & _
"('0','460200','Keranjang, barang anyaman dan barang lainny dibuat secara langsung menjadi berbentuk dari bahan anyaman atau dibuat dari barang dalam pos 46.01 barang dari loofah.')," & _
"('0','470000','Pulp dari kayu atau dari bahan selulosa berserat lainnya Kertas atau kertas karton yang dipulihkan (sisa dan skrap)')," & _
"('0','470100','Pulp kayu mekanik.')," & _
"('0','470200','Pulp kayu kimi dissolving grade.')," & _
"('0','470300','Pulp kayu kimi soda atau sulfat, selain dissolving grade.')," & _
"('0','470400','Pulp kayu kimi sulfit, selain dissolving grade.')," & _
"('0','470500','Pulp kayu yang diperoleh melalui kombinasi proses pembuatan pulp secara mekanik dan kimia.')," & _
"('0','470600','Pulp dari serat yang berasal dari kertas atau kertas karton yang dipulihkan (sisa dan skrap) atau dari bahan selulosa berserat lainnya.')," & _
"('0','470700','Kertas atau kertas karton yang dipulihkan (sisa dan skrap).')," & _
"('0','480000','Kertas dan kertas karton barang dari pulp kertas, dari kertas atau dari kertas karton')," & _
"('0','480100','Kertas koran, dalam gulungan atau lembaran.')," & _
"('0','480200','Kertas dan kertas karton tidak dilapisi, dari jenis yang digunakan untuk penulisan, pencetakan atau keperluan grafik lainny serta kertas untuk punch card serta punch tape tidak dilubangi, dalam gulungan atau lembaran empat persegi panjang (termasuk bujur sangkar) dari berbagai ukuran, selain dari pos 48.01 atau 48.03 kertas dan kertas karton buatan tangan.')," & _
"('0','480300','Kertas toilet atau kertas tisu untuk kulit muk kertas handuk atau kertas serbet dan kertas semacam itu dari jenis yang digunakan untuk keperluan rumah tangga atau saniter, gumpalan selulosa dan web dari serat selulos dikisutkan, dikerutkan, diembos, dilubangi, diwarnai permukaanny dihias atau dicetak permukaannya maupun tidak, dalam gulungan atau lembaran.')," & _
"('0','480400','Kertas kraft dan kertas karton tidak dilapisi, dalam gulungan atau lembaran, selain yang dimaksud dalam pos 48.02 atau 48.03.')," & _
"('0','480500','Kertas dan kertas karton tidak dilapisi lainny dalam gulungan atau lembaran, tidak dikerjakan atau diproses lebih lanjut selain yang dirinci dalam Catatan 3 pada Bab ini.')," & _
"('0','480600','Perkamen nabati, kertas tahan lemak, kertas kalkir dan kertas glasin serta kertas transparan dikilapkan atau kertas bening lainny dalam gulungan atau lembaran.')," & _
"('0','480700','Kertas komposit dan kertas karton komposit (dibuat dengan merekatkan beberapa lapisan datar kertas atau kertas karton dengan perekat), permukaannya tidak dilapisi atau diresapi , bagian dalamnya diperkuat maupun tidak, dalam gulungan atau lembaran.')," & _
"('0','480800','Kertas dan kertas karton, bergelombang (dengan atau tanpa dilekati lembaran yang datar permukaannya), dikisutkan, dikerutkan, diembos atau dilubangi, dalam gulungan atau lembaran, selain kertas dari jenis yang diuraikan dalam pos 48.03.')," & _
"('0','480900','Kertas karbon, kertas self-copy dan kertas kopi atau kertas transfer lainnya (termasuk kertas dilapisi atau kertas diresapi untuk stensil duplikator atau pelat offset), dicetak maupun tidak, dalam gulungan atau lembaran.')," & _
"('0','481000','Kertas dan kertas karton, dilapisi satu atau kedua sisinya dengan kaolin (tanah liat Cina) atau zat anorganik lainny dengan atau tanpa bahan pengikat, dan tanpa pelapis lainny diwarnai, dihias atau dicetak permukaannya maupun tidak, dalam gulungan atau lembaran empat persegi panjang (termasuk bujur sangkar) dari berbagai ukuran.')," & _
"('0','481100','Kertas, kertas karton, gumpalan selulosa dan web dari serat selulos dilapisi, diresapi, ditutupi, diwarnai permukaanny dihias atau dicetak permukaanny dalam gulungan atau lembaran empat persegi panjang (termasuk bujur sangkar), dari berbagai ukuran, selain barang dari jenis yang diuraikan dalam pos 48.03, 48.09 atau 48.10.')," & _
"('0','481200','Blok, lempengan dan pelat saring, dari pulp kertas.')," & _
"('0','481300','Kertas sigaret, dipotong menurut ukuran maupun tidak, baik dalam bentuk buklet atau tabung.')," & _
"('0','481400','Wallpaper dan penutup dinding semacam itu kertas transparansi untuk jendela.')," & _
"('0','481600','Kertas karbon, kertas self-copy dan kertas kopi atau kertas transfer lainnya (selain yang dimaksud dalam pos 48.09), stensil duplikator dan pelat offset, dari kertas, disiapkan dalam kotak maupun tidak.')," & _
"('0','481700','Amplop, kartu surat, kartu pos polos dan kartu korespondensi, dari kertas atau kertas karton kotak, kantong, dompet dan kompendium tulisan, dari kertas atau kertas karton, berisi bermacam-macam kertas surat.')," & _
"('0','481800','Kertas toilet dan kertas semacam itu, gumpalan selulosa atau web dari serat selulos dari jenis yang digunakan untuk keperluan rumah tangga atau saniter, dalam gulungan dengan lebar tidak melebihi 36 cm, atau dipotong menurut ukuran atau bentuk tertentu saputangan, tisu pembersih, handuk, taplak mej serbet, seprei dan keperluan rumah tangga semacam itu, barang keperluan saniter atau rumah sakit, pakaian dan aksesori pakaian, dari pulp kertas, kertas, gumpalan selulosa atau web serat selulosa.')," & _
"('0','481900','Kardus, kotak, peti, tas dan kemasan lainnya dari kertas, kertas karton, gumpalan selulosa atau web dari serat selulosa kotak arsip, nampan surat dan barang semacam itu, dari kertas atau kertas karton, dari jenis yang digunakan di kantor, di toko atau sejenisnya.')," & _
"('0','482000','Buku daftar, buku kas, buku catatan, buku pesanan, buku tanda terim blok kertas surat, blok kertas memo, buku harian dan barang semacam itu, buku tulis, blok kertas isap, penjilid (lembaran lepas atau lain-lain), map, penyimpan berkas, formulir bermacam-macam bisnis, karbon set bersela lembaran kosong dan barang tulis-menulis lainny dari kertas atau kertas karton album untuk contoh atau untuk koleksi dan sampul buku, dari kertas atau kertas karton.')," & _
"('0','482100','Label kertas atau kertas karton dari segala jenis, dicetak maupun tidak.')," & _
"('0','482200','Bobbin, kelos, cop dan alat penunjang semacam itu dari pulp kertas, kertas atau kertas karton (dilubangi atau dikeraskan maupun tidak).')," & _
"('0','482300','Kertas, kertas karton, gumpalan selulosa dan web serat selulosa lainny dipotong menurut ukuran atau bentuk barang lainnya dari pulp kertas, kertas, kertas karton, gumpalan selulosa atau web serat selulosa.')," & _
"('0','490000','Buku, koran, gambar cetakan dan produk lainnya dari industri percetakan naskah tulisan tangan, naskah ketikan dan rencana')," & _
"('0','490100','Buku cetakan, brosur, selebaran dan barang cetakan semacam itu, dalam lembaran tunggal maupun tidak.')," & _
"('0','490200','Koran, jurnal dan majalah berkal bergambar atau berisi iklan maupun tidak.')," & _
"('0','490300','Buku bergambar, buku untuk menggambar atau mewarnai untuk anak-anak.')," & _
"('0','490400','Buku musik, dicetak atau dalam bentuk manuskrip, dijilid atau bergambar maupun tidak.')," & _
"('0','490500','Peta dan peta hidrografi atau peta semacam itu dari segala jenis, termasuk atlas, peta dinding, peta topografi serta bola duni dicetak.')," & _
"('0','490600','Rancangan dan gambar untuk arsitektur, permesinan, industri, perdagangan, topografi atau keperluan semacam itu, asli buatan tangan teks tulisan tangan reproduksi fotografi di atas kertas peka cahaya dan carbon copy dari barang-barang tersebut.')," & _
"('0','490700','Perangko, meterai atau perangko semacam itu belum digunakan, yang berlaku atau baru diterbitkan oleh negar yang mempunyai atau akan mempunyai harga yang nilainya diakui kertas meterai uang kertas formulir cek sertifikat saham, andil atau obligasi dan dokumen kepemilikan semacam itu.')," & _
"('0','490800','Gambar tempel (decalcomania).')," & _
"('0','490900','Kartu pos dicetak atau diberi gambar kartu ucapan selamat pribadi, pesan atau pengumuman, diberi gambar maupun tidak, dengan atau tanpa amplop atau hiasan.')," & _
"('0','491000','Berbagai jenis kalender, dicetak, termasuk blok kalender.')," & _
"('0','491100','Barang cetakan lainny termasuk gambar dan foto cetakan.')," & _
"('0','500000','Sutra')," & _
"('0','500100','Kepompong ulat sutra cocok untuk digulung.')," & _
"('0','500200','Sutra mentah (tidak dipintal).')," & _
"('0','500300','Sisa sutra (termasuk kepompong tidak cocok untuk digulung, sisa benang dan garnetted stock).')," & _
"('0','500400','Benang sutra (selain benang pintal dari sisa sutra) tidak disiapkan untuk penjualan eceran.')," & _
"('0','500500','Benang pintal dari sisa sutr tidak disiapkan untuk penjualan eceran.')," & _
"('0','500600','Benang sutra dan benang pintal dari sisa sutr disiapkan untuk penjualan eceran benang ulat sutra.')," & _
"('0','500700','Kain tenunan dari sutra atau sisa sutra.')," & _
"('0','510000','Wol, bulu hewan halus atau kasar benang bulu kuda dan kain tenunan')," & _
"('0','510100','Wol, tidak digaruk atau disisir.')," & _
"('0','510200','Bulu hewan halus atau kasar, tidak digaruk atau disisir.')," & _
"('0','510300','Sisa dari wol atau dari bulu hewan halus atau kasar, termasuk sisa benang tetapi tidak termasuk garnetted stock.')," & _
"('0','510400','Garnetted stock dari wol atau dari bulu hewan halus atau kasar.')," & _
"('0','510500','Wol dan bulu hewan halus atau kasar, digaruk atau disisir (termasuk wol disisir dalam bentuk fragmen).')," & _
"('0','510600','Benang dari wol digaruk, tidak disiapkan untuk penjualan eceran.')," & _
"('0','510700','Benang dari wol disisir, tidak disiapkan untuk penjualan eceran.')," & _
"('0','510800','Benang dari bulu hewan halus (digaruk atau disisir), tidak disiapkan untuk penjualan eceran.')," & _
"('0','510900','Benang dari wol atau dari bulu hewan halus, disiapkan untuk penjualan eceran.')," & _
"('0','511000','Benang dari bulu hewan kasar atau dari bulu kuda (termasuk benang lilit dari bulu kuda), disiapkan untuk penjualan eceran maupun tidak.')," & _
"('0','511100','Kain tenunan dari wol atau bulu hewan halus digaruk.')," & _
"('0','511200','Kain tenunan dari wol atau bulu hewan halus disisir.')," & _
"('0','511300','Kain tenunan dari bulu hewan kasar atau bulu kuda.')," & _
"('0','520000','Kapas')," & _
"('0','520100','Kapas, tidak digaruk atau disisir.')," & _
"('0','520200','Sisa kapas (termasuk sisa benang dan garnetted stock).')," & _
"('0','520300','Kapas, digaruk atau disisir.')," & _
"('0','520400','Benang jahit dari kapas, disiapkan untuk penjualan eceran maupun tidak.')," & _
"('0','520500','Benang kapas (selain benang jahit), mengandung kapas 85 % atau lebih menurut beratny tidak disiapkan untuk penjualan eceran.')," & _
"('0','520600','Benang kapas (selain benang jahit), mengandung kapas kurang dari 85 % menurut beratny tidak disiapkan untuk penjualan eceran.')," & _
"('0','520700','Benang kapas (selain benang jahit) disiapkan untuk penjualan eceran.')," & _
"('0','520800','Kain tenunan dari kapas, mengandung kapas 85 % atau lebih menurut beratny beratnya tidak lebih dari 200 g/m2.')," & _
"('0','520900','Kain tenunan dari kapas, mengandung kapas 85 % atau lebih menurut beratny beratnya lebih dari 200 g/m2.')," & _
"('0','521000','Kain tenunan dari kapas, mengandung kapas kurang dari 85 % menurut beratny dicampur terutama atau semata-mata dengan serat buatan, beratnya tidak lebih dari 200 g/m2.')," & _
"('0','521100','Kain tenunan dari kapas, mengandung kapas kurang dari 85 % menurut beratny dicampur terutama atau semata-mata dengan serat buatan, beratnya lebih dari 200 g/m2.')," & _
"('0','521200','Kain tenunan lainnya dari kapas.')," & _
"('0','530000','Serat tekstil nabati lainnya benang kertas dan kain tenunan Clari benang kertas')," & _
"('0','530100','Len mentah atau sudah dikerjakan tetapi tidak dipintal tow lena dan sisa lena (termasuk sisa benang dan garnetted stock).')," & _
"('0','530200','True hemp (Cannabis sativa L.), mentah atau diolah tetapi tidak dipintal tow dan sisa dari true hemp (termasuk sisa benang dan garnetted stock).')," & _
"('0','530300','Serat jute dan serat tekstil kulit pohon lainnya (tidak termasuk len true hemp dan rami), mentah atau diolah tetapi tidak dipintal tow dan sisa dari serat tersebut (termasuk sisa benang dan garnetted stock).')," & _
"('0','530500','Serat kelap abaca (serat pisang manila atau Musa textilis Nee), rami dan serat tekstil nabati lainny tidak dirinci atau termasuk dalam pos manapun, mentah atau dikerjakan tetapi tidak dipintal tow, noil dan sisa dari serat itu (termasuk sisa benang dan garnetted stock).')," & _
"('0','530600','Benang lena.')," & _
"('0','530700','Benang dari serat jute atau dari serat tekstil kulit pohon lainnya dari pos 53.03.')," & _
"('0','530800','Benang dari serat tekstil nabati lainnya benang kertas.')," & _
"('0','530900','Kain tenunan dari lena.')," & _
"('0','531000','Kain tenunan dari serat jute atau dari serat tekstil kulit pohon lainnya dari pos 53.03.')," & _
"('0','531100','Kain tenunan dari serat tekstil nabati lainnya kain tenunan dari benang kertas.')," & _
"('0','540000','Filamen buatan strip dan sejenisnya dari bahan tekstil buatan')," & _
"('0','540100','Benang jahit dari filamen buatan, disiapkan untuk penjualan eceran maupun tidak.')," & _
"('0','540200','Benang filamen sintetik (selain benang jahit), tidak disiapkan untuk penjualan eceran, termasuk monofilamen sintetik yang kurang dari 67 desiteks.')," & _
"('0','540300','Benang filamen artifisial (selain benang jahit), tidak disiapkan untuk penjualan eceran, termasuk monofilamen artifisial yang kurang dari 67 desiteks.')," & _
"('0','540400','Monofilamen sintetik dengan ukuran 67 desiteks atau lebih dan yang ukuran penampang silangnya tidak ada yang lebih dari 1 mm strip dan sejenisnya (misalny jerami artifisial) dari bahan tekstil sintetik yang mempunyai lebar tidak melebihi 5 mm.')," & _
"('0','540500','Monofilamen artifisial dengan ukuran 67 desiteks atau lebih dan yang ukuran penampang silangnya tidak ada yang lebih dari 1 mm strip dan sejenisnya (misalny jerami artifisial) dari bahan tekstil artifisial yang mempunyai lebar tidak melebihi 5 mm.')," & _
"('0','540600','Benang filamen buatan (selain benang jahit), disiapkan untuk penjualan eceran.')," & _
"('0','540700','Kain tenunan dari benang filamen sintetik, termasuk kain tenunan yang diperoleh dari bahan dari pos 54.04.')," & _
"('0','540800','Kain tenunan dari benang filamen artifisial, termasuk kain tenunan yang diperoleh dari bahan dari pos 54.05.')," & _
"('0','550000','Serat Stapel Buatan')," & _
"('0','550100','Tow filamen sintetik.')," & _
"('0','550200','Tow filamen artifisial.')," & _
"('0','550300','Serat stapel sintetik, tidak digaruk, disisir atau diproses secara lain untuk dipintal.')," & _
"('0','550400','Serat stapel artifisial, tidak digaruk, disisir atau diproses secara lain untuk dipintal.')," & _
"('0','550500','Sisa (termasuk noil, sisa benang dan garnetted stock) dari serat buatan.')," & _
"('0','550600','Serat stapel sintetik, digaruk, disisir atau diproses secara lain untuk dipintal.')," & _
"('0','550700','Serat stapel artifisial, digaruk, disisir atau diproses secara lain untuk dipintal.')," & _
"('0','550800','Benang jahit dari serat stapel buatan, disiapkan untuk penjualan eceran maupun tidak.')," & _
"('0','550900','Benang (selain benang jahit) dari serat stapel sintetik, tidak disiapkan untuk penjualan eceran.')," & _
"('0','551000','Benang (selain benang jahit) dari serat stapel artifisial, tidak disiapkan untuk penjualan eceran.')," & _
"('0','551100','Benang (selain benang jahit) dari serat stapel buatan, disiapkan untuk penjualan eceran.')," & _
"('0','551200','Kain tenunan dari serat stapel sintetik, mengandung serat stapel sintetik 85 % atau lebih menurut beratnya.')," & _
"('0','551300','Kain tenunan dari serat stapel sintetik, mengandung serat tersebut kurang dari 85 % menurut beratny dicampur terutama atau semata-mata dengan kapas, dengan berat tidak melebihi 170 g/m2.')," & _
"('0','551400','Kain tenunan dari serat stapel sintetik, mengandung serat tersebut kurang dari 85 % menurut beratny dicampur terutama atau semata-mata dengan kapas, dengan berat melebihi 170 g/m2.')," & _
"('0','551500','Kain tenunan lainnya dari serat stapel sintetik.')," & _
"('0','551600','Kain tenunan dari serat stapel artifisial.')," & _
"('0','560000','Gumpalan, kain kempa dan bukan tenunan benang khusus benang pintal, tali, tambang dan kabel serta barang daripadanya')," & _
"('0','560100','Gumpalan bahan tekstil dan barang daripadanya serat tekstil, panjangnya tidak melebihi 5 mm (flock), debu tekstil dan mill nep.')," & _
"('0','560200','Kain kemp diresapi, dilapisi, ditutupi atau dilaminasi maupun tidak.')," & _
"('0','560300','Bukan tenunan, diresapi, dilapisi, ditutupi atau dilaminasi maupun tidak.')," & _
"('0','560400','Benang dan tali karet, ditutupi tekstil benang tekstil, dan strip serta sejenisnya dari pos 54.04 atau 54.05, diresapi, dilapisi, ditutupi atau diselubungi dengan karet atau plastik.')," & _
"('0','560500','Benang berlogam, berpalut atau tidak, menjadi benang tekstil, atau strip serta sejenisnya dari pos 54.04 atau 54.05, dikombinasikan dengan logam dalam bentuk benang, strip atau bubuk atau ditutupi dengan logam.')," & _
"('0','560600','Benang berpalut, dan strip serta sejenisnya dari pos 54.04 atau 54.05, berpalut (selain yang disebutkan dalam pos 56.05 dan benang bulu kuda berpalut) benang chenille (termasuk benang flock chenille) benang loop-wale.')," & _
"('0','560700','Benang pintal, tali, tambang dan kabel, dianyam atau dijalin maupun tidak dan diresapi, dilapisi, ditutupi atau diselubungi dengan karet atau plastik maupun tidak.')," & _
"('0','560800','Jaring disimpul dari benang pintal, tali atau tambang jaring ikan jadi dan jaring jadi lainny dari bahan tekstil.')," & _
"('0','560900','Barang dari benang, strip atau sejenisnya dari pos 54.04 atau 54.05, benang pintal, tali, tambang atau kabel, tidak dirinci atau termasuk dalam pos lainnya.')," & _
"('0','570000','Karpet dan tekstil penutup lantai lainnya')," & _
"('0','570100','Karpet dan penutup lantai tekstil lainny disimpul, sudah jadi maupun belum.')," & _
"('0','570200','Karpet dan penutup lantai tekstil lainny tenunan, tidak berumbai atau tidak dibentuk flock, sudah jadi maupun belum, termasuk Kelem')," & _
"('0','570300','Karpet dan penutup lantai tekstil lainnya (termasuk turf), berumbai, sudah jadi maupun belum.')," & _
"('0','570400','Karpet dan penutup lantai tekstil lainny dari kain kemp tidak berumbai atau tidak dibentuk flock, sudah jadi maupun belum.')," & _
"('0','570500','Karpet dan penutup lantai tekstil lainny sudah jadi maupun belum.')," & _
"('0','580000','Kain tenunan khusus kain tekstil berjumbai renda permadani dinding hiasan sulaman')," & _
"('0','580100','Kain tenunan berbulu dan kain chenille, selain kain dari pos 58.02 atau 58.06.')," & _
"('0','580200','Terry towelling dan kain tenunan terry semacam itu, selain kain pita dari pos 58.06 kain tekstil berumbai, selain produk dari pos 57.03.')," & _
"('0','580300','Kas selain kain pita dari pos 58.06.')," & _
"('0','580400','Kain tule dan kain jaring lainny tidak termasuk kain tenunan, rajutan atau kaitan renda dalam lembaran, strip atau motif, selain kain dari pos 60.02 sampai dengan 60.06.')," & _
"('0','580500','Permadani dinding tenunan tangan dari tipe Gobelin, Flander, Aubusson, Beauvais dan sejenisny dan permadani dinding dikerjakan dengan jarum (misalny bintik kecil, jeratan silang), sudah jadi maupun belum.')," & _
"('0','580600','Kain pita tenunan, selain barang dari pos 58.07 kain pita terdiri dari benang lusi tanpa benang pakan yang disatukan dengan perekat (bolduc).')," & _
"('0','580700','Label, lencana dan barang semacam itu dari bahan tekstil, dalam lembaran, strip atau dipotong menjadi berbentuk atau berukuran, tidak disulam.')," & _
"('0','580800','Kain jalinan dalam lembaran kain perapih hiasan dalam lembaran, tanpa sulaman, selain rajutan atau kaitan jumbai, pompon dan barang semacam itu.')," & _
"('0','580900','Kain tenunan dari benang logam dan kain tenunan dari benang berlogam dari pos 56.05, dari jenis yang digunakan dalam pakaian, sebagai kain perabotan atau untuk keperluan semacam itu, tidak dirinci atau termasuk dalam pos lainnya.')," & _
"('0','581000','Kain sulaman dalam lembaran, strip atau motif.')," & _
"('0','581100','Quilted textile products dalam lembaran, disusun dari satu atau lebih lapisan bahan tekstil disatukan dengan padding dengan cara dijahit atau secara lain, selain kain sulaman dari pos 58.10.')," & _
"('0','590000','Kain tekstil diresapi, dilapisi, ditutupi atau dilaminasi barang tekstil dari jenis yang cocok untuk keperluan industri')," & _
"('0','590100','Kain tekstil dilapisi dengan perekat atau zat mengandung pati, dari jenis yang digunakan untuk kulit luar buku atau sejenisnya kain kalkir kanvas lukis siap pakai buckram dan kain tekstil kaku semacam itu dari jenis yang digunakan untuk dasar topi.')," & _
"('0','590200','Kain untuk ban dari benang dengan kekuatan tinggi dari nilon atau poliamida lainny poliester atau rayon viskose.')," & _
"('0','590300','Kain tekstil diresapi, dilapisi, ditutupi atau dilaminasi dengan plastik, selain yang dimaksud dalam pos 59.02.')," & _
"('0','590400','Linoleum, dipotong menjadi bentuk maupun tidak penutup lantai yang terdiri dari lapisan atau penutup di atas dasar tekstil, dipotong menjadi bentuk maupun tidak.')," & _
"('0','590500','Penutup dinding tekstil.')," & _
"('0','590600','Kain tekstil berkaret, selain yang dimaksud dalam pos 59.02.')," & _
"('0','590700','Kain tekstil diresapi, dilapisi atau ditutupi secara lain kanvas dilukis menjadi layar pentas, kain latar belakang studio atau sejenisnya.')," & _
"('0','590800','Sumbu, tenunan, anyaman atau rajutan dari tekstil untuk lampu, kompor, korek api, lilin atau sejenisnya kaos lampu gas pijar dan kain kaos lampu rajutan berbentuk tabung untuk kaos lampu gas, diresapi maupun tidak.')," & _
"('0','590900','Selang tekstil dan tabung tekstil semacam itu, dengan atau tanpa lapisan, pelindung atau aksesori dari bahan lain.')," & _
"('0','591000','Ban atau belting penggerak atau pengangkut, dari bahan tekstil, diresapi, dilapisi, ditutupi atau dilaminasi dengan plastik, atau diperkuat dengan logam atau bahan lain maupun tidak.')," & _
"('0','591100','Produk dan barang tekstil, untuk penggunaan teknis, dirinci dalam Catatan 8 pada Bab ini.')," & _
"('0','600000','Kain rajutan atau kain kaitan')," & _
"('0','600100','Kain berbulu, termasuk kain berbulu panjang dan kain terry')," & _
"('0','600200','Kain rajutan atau kaitan dengan lebar tidak melebihi 30 cm, mengandung benang elastomer atau benang karet 5 % atau lebih menurut beratny selain yang dimaksud dalam pos 60.01.')," & _
"('0','600300','Kain rajutan atau kaitan dengan lebar tidak melebihi 30 cm, selain yang dimaksud dalam pos 60.01 atau 60.02.')," & _
"('0','600400','Kain rajutan atau kaitan dengan lebar melebihi 30 cm, mengandung benang elastomer atau benang karet 5 % atau lebih menurut beratny selain yang dimaksud dalam pos 60.01.')," & _
"('0','600500','Kain rajut lusi (termasuk kain yang dibuat dengan mesin rajut galon), selain yang dimaksud dalam pos 60.01 sampai dengan 60.04.')," & _
"('0','600600','Kain rajutan atau kaitan lainnya.')," & _
"('0','610000','Pakaian dan aksesori pakaian, barang rajutan atau kaitan')," & _
"('0','610100','Mantel panjang, car-coat, jubah bertopi, jubah, anorak (termasuk jaket-ski), wind-cheater, wind-jacket dan barang semacam itu, untuk pria atau anak laki-laki, rajutan atau kaitan, selain yang dimaksud dalam pos 61.03.')," & _
"('0','610200','Mantel panjang, car-coat, jubah bertopi, jubah, anorak (termasuk jaket-ski), wind-cheater, wind-jacket dan barang semacam itu, untuk wanita atau anak perempuan, rajutan atau kaitan, selain yang dimaksud dalam pos 61.04.')," & _
"('0','610300','Setelan, ensemble, jas, blazer, celana panjang, pakaian terusan berpenutup di depan dan bertali, celana panjang sampai lutut dan celana pendek (selain pakaian renang), untuk pria atau anak laki-laki, rajutan atau kaitan.')," & _
"('0','610400','Setelan, ensemble, jas, blazer, gaun, rok, rok terpisah, celana panjang, pakaian terusan berpenutup di depan dan bertali, celana panjang sampai lutut dan celana pendek (selain pakaian renang), untuk wanita atau anak perempuan, rajutan atau kaitan.')," & _
"('0','610500','Kemeja pria atau anak laki-laki, rajutan atau kaitan.')," & _
"('0','610600','Blus, kemeja dan kemeja blus, untuk wanita atau anak perempuan, rajutan atau kaitan.')," & _
"('0','610700','Celana kolor, celana dalam, kemeja tidur, piyam pakaian mandi, dressing gown dan barang semacam itu, untuk pria atau anak laki-laki, rajutan atau kaitan.')," & _
"('0','610800','Rok dalam, petticoat, celana dalam, panty, gaun malam, piyam negligee, pakaian mandi, dressing gown dan barang semacam itu, untuk wanita atau anak perempuan, rajutan atau kaitan.')," & _
"('0','610900','T-shirt, singlet dan kaus kutang lainny rajutan atau kaitan.')," & _
"('0','611000','Jersey, pullover, cardigan, rompi dan barang semacam itu, rajutan atau kaitan.')," & _
"('0','611100','Garmen dan aksesori pakaian untuk bayi, rajutan atau kaitan.')," & _
"('0','611200','Track suit, ski suit dan pakaian renang, rajutan atau kaitan.')," & _
"('0','611300','Garmen, dibuat dari kain rajutan atau kaitan dari pos 59.03, 59.06 atau 59.07.')," & _
"('0','611400','Garmen lainny rajutan atau kaitan.')," & _
"('0','611500','Panty hose, tight, stocking, kaus kaki dan hosiery lainny termasuk hosiery dikempa bergradasi (misalny stocking untuk penderita varises) dan alas kaki tanpa sol, rajutan atau kaitan.')," & _
"('0','611600','Sarung tangan, mitten dan mitt, rajutan atau kaitan.')," & _
"('0','611700','Aksesori pakaian jadi lainny rajutan atau kaitan bagian dari garmen atau dari aksesori pakaian, rajutan atau kaitan.')," & _
"('0','620000','Pakaian dan aksesori pakaian, bukan rajutan atau kaitan')," & _
"('0','620100','Mantel panjang, car-coat, jubah bertopi, cloak, anorak (termasuk jaket-ski), wind-cheater, wind-jacket dan barang semacam itu, untuk pria atau anak laki-laki, selain yang dimaksud dalam pos 62.03.')," & _
"('0','620200','Mantel panjang, car-coat, jubah bertopi, cloak, anorak (termasuk jaket-ski), wind-cheater, wind-jacket dan barang semacam itu, untuk wanita atau anak perempuan, selain yang dimaksud dalam pos 62.04.')," & _
"('0','620300','Setelan, ensemble, jas, blazer, celana panjang, pakaian terusan berpenutup di depan dan bertali, celana panjang sampai lutut dan celana pendek (selain pakaian renang), untuk pria atau anak laki-laki.')," & _
"('0','620400','Setelan, ensemble, jas, blazer, gaun, rok, rok terpisah, celana panjang, pakaian terusan berpenutup di depan dan bertali, celana panjang sampai lutut dan celana pendek (selain pakaian renang), untuk wanita atau anak perempuan.')," & _
"('0','620500','Kemeja pria atau anak laki-laki.')," & _
"('0','620600','Blus, kemeja dan blus kemej untuk wanita atau anak perempuan.')," & _
"('0','620700','Singlet dan kaus kutang lainny celana kolor, celana dalam, pakaian tidur, piyam pakaian mandi, dressing gown dan barang semacam itu, untuk pria atau anak laki-laki.')," & _
"('0','620800','Singlet dan kaus kutang lainny pakaian dalam kombinasi, petticoat, celana dalam, panty, gaun malam, piyam negligee, pakaian mandi, dressing gown dan barang semacam itu, untuk wanita atau anak perempuan.')," & _
"('0','620900','Garmen dan aksesori pakaian bayi.')," & _
"('0','621000','Garmen, dibuat dari kain dari pos 56.02, 56.03, 59.03, 59.06, atau 59.07.')," & _
"('0','621100','Track suit, ski suit dan pakaian renang garmen lainnya.')," & _
"('0','621200','Brassiere, girdle, korset, brace, suspender, garter dan barang semacam itu serta bagianny rajutan atau kaitan maupun tidak.')," & _
"('0','621300','Saputangan.')," & _
"('0','621400','Syal, scarf, muffler, mantill veil dan sejenisnya.')," & _
"('0','621500','Dasi, dasi kupu- kupu dan cravat.')," & _
"('0','621600','Sarung tangan, mitten dan mitt.')," & _
"('0','621700','Aksesori pakaian jadi lainnya bagian dari garmen atau dari aksesori pakaian, selain yang dimaksud dalam pos 62.12.')," & _
"('0','630000','Barang tekstil sudah jadi lainnya set pakaian bekas dan barang tekstil bekas gombal')," & _
"('0','630100','Selimut dan selimut kecil untuk perjalanan.')," & _
"('0','630200','Linen untuk tempat tidur, mej toilet dan dapur.')," & _
"('0','630300','Tirai (termasuk gorden) dan kerai dalam tirai atau bed valances.')," & _
"('0','630400','Barang perabotan lainny tidak termasuk yang dimaksud dalam pos 94.04.')," & _
"('0','630500','Kantong dan karung, dari jenis yang digunakan untuk membungkus barang.')," & _
"('0','630600','Terpal, awning dan kerai matahari tenda (termasuk temporary canopy dan barang semacam itu) layar untuk perahu, papan selancar atau landcraft barang keperluan berkemah.')," & _
"('0','630700','Barang jadi lainny termasuk pola pakaian.')," & _
"('0','630800','Set terdiri dari kain tenunan dan benang, dengan aksesori maupun tidak, untuk dibuat menjadi babut, permadani dinding, kain meja sulaman atau serbet, atau barang tekstil semacam itu, disiapkan dalam kemasan untuk penjualan eceran.')," & _
"('0','630900','Pakaian bekas dan barang bekas lainnya.')," & _
"('0','631000','Rag bekas atau baru, skrap dari benang pintal, tali , tali tambang dan kabel serta barang usang dari benang pintal, tali tambang atau kabel, dari bahan tekstil.')," & _
"('0','640000','Alas kaki, pelindung kaki dan sejenisnya bagian dari barang tersebut')," & _
"('0','640100','Alas kaki tahan air dengan sol luar dan bagian atas dari karet atau dari plastik, bagian atasnya tidak dipasang pada sol dan tidak dirakit dengan cara dijahit, dikeling, dipaku, disekrup, ditusuk atau proses semacam itu.')," & _
"('0','640200','Alas kaki lainnya dengan sol luar dan bagian atas dari karet atau plastik.')," & _
"('0','640300','Alas kaki dengan sol luar dari karet, plastik, kulit samak atau kulit komposisi dan bagian atas sepatu dari kulit samak.')," & _
"('0','640400','Alas kaki dengan sol luar dari karet, plastik, kulit samak atau kulit komposisi dan bagian atasnya dari bahan tekstil.')," & _
"('0','640500','Alas kaki lainnya.')," & _
"('0','640600','Bagian dari alas kaki (termasuk bagian atas dipasang sol maupun tidak selain sol luar) sol dalam yang dapat dilepas, bantalan tumit dan barang semacam itu pelindung kaki, pembalut kaki dan barang semacam itu, serta bagiannya.')," & _
"('0','650000','Tutup kepala dan bagiannya')," & _
"('0','650100','Bentuk topi, badan topi dan tudung topi dari kain kemp tidak diberi bentuk atau tidak dibuat pinggirannya plateux dan manchon (termasuk manchon belah), dari kain kempa.')," & _
"('0','650200','Model topi, dianyam atau dibuat dengan merakit strip dari berbagai bahan, tidak diberi bentuk, tidak dibuat pinggiran, tidak diberi garis dan tidak dirapikan pinggirannya.')," & _
"('0','650400','Topi dan tutup kepala lainny dianyam atau dibuat dengan merakit strip dari berbagai bahan, diberi garis atau dirapikan pinggirannya maupun tidak.')," & _
"('0','650500','Topi dan tutup kepala lainny rajutan atau kaitan, atau dibuat dari rend kain kempa atau kain tekstil lainny dalam lembaran (tetapi tidak dalam bentuk strip), diberi garis, dirapikan pinggirannya maupun tidak jaring rambut dari berbagai bahan, diberi garis dirapikan pinggirannya maupun tidak.')," & _
"('0','650600','Tutup kepala lainny diberi garis atau dirapikan pinggirannya maupun tidak.')," & _
"('0','650700','Ikat kepal pelapis topi, tutup topi, dasar topi, rangka topi, klep dan tali dagu, untuk tutup kepala.')," & _
"('0','660000','Payung, payung panas, tongkat jalan, tongkat duduk, cambuk, pecut dan bagiannya')," & _
"('0','660100','Payung dan payung panas (termasuk payung berbentuk tongkat jalan, payung taman dan payung semacam itu).')," & _
"('0','660200','Tongkat jalan, tongkat duduk, cambuk, pecut dan sejenisnya.')," & _
"('0','660300','Bagian, hiasan dan aksesori barang dari pos 66.01 atau 66.02.')," & _
"('0','670000','Bulu dan bulu halus unggas olahan serta barang dibuat dari bulu atau bulu halus unggas bunga artifisial barang dari rambut manusia')," & _
"('0','670100','Kulit dan bagian lainnya dari unggas dengan bulu atau bulu halusny bulu, bagian dari bulu, bulu halus dan barang terbuat dari padanya (selain barang dari pos 05.05 dan pena bulu angsa serta bulu lepas olahan).')," & _
"('0','670200','Bung daun dan buah artifisial serta bagian daripadanya barang dibuat dari bung daun atau buah artifisial.')," & _
"('0','670300','Rambut manusi ditat ditipiskan, dikelantang atau dikerjakan secara lain wol atau bulu hewan lainnya atau bahan tekstil lainny disiapkan untuk digunakan dalam pembuatan wig atau sejenisnya.')," & _
"('0','670400','Wig, jenggot, alis dan bulu mata palsu, cemara dan sejenisny dari rambut manusia atau bulu hewan atau dari bahan tekstil barang dari rambut manusia tidak dirinci atau termasuk dalam pos lain.')," & _
"('0','680000','Barang dari batu, plester, semen, asbes, mika atau bahan semacam itu')," & _
"('0','680100','Batu jalan, batu tepi jalan dan batu ubin, dari batu alami (kecuali batu sabak).')," & _
"('0','680200','Batu monumen dan batu bangunan dikerjakan (kecuali batu sabak) dan barang terbuat dari padany selain barang dari pos 68.01 kubus mosaik dan sejenisny dari batu alam (termasuk batu sabak), dengan alas maupun tidak butiran, kepingan dan bubuk dengan warna artifisial dari batu alam (termasuk batu sabak).')," & _
"('0','680300','Batu sabak dikerjakan dan barang dari batu sabak atau dari batu sabak diaglomerasi.')," & _
"('0','680400','Batu giling, batu asah, roda gerinda dan sejenisny tanpa rangk untuk pengasah, penajam, pemoles, pemotong, batu penajam atau pemoles yang digerakkan dengan tangan, dan bagian daripadany dari batu alam, dari batu gosok alam atau artifisial diaglomerasi, atau dari keramik, dengan atau tanpa bagian dari bahan lain.')," & _
"('0','680500','Bubuk atau butir gosok alam atau artifisial, dengan dasar dari bahan tekstil, kertas, kertas karton atau bahan lainny dipotong menjadi bentuk atau dijahit atau dibuat secara lain maupun tidak.')," & _
"('0','680600','Wol terak, wol batuan dan wol mineral semacam itu vermikulit melar, tanah liat muai, terak buih dan bahan mineral muai semacam itu campuran dan barang dari bahan mineral isolasi tahan panas, isolasi suara atau serap suar selain yang dimaksud dalam pos 68.11 atau 68.12 atau dari Bab 69.')," & _
"('0','680700','Barang dari aspal atau dari bahan semacam itu (misalny bitumen petroleum atau pek dari ter batu bara).')," & _
"('0','680800','Panel, papan, ubin, blok dan barang semacam itu terbuat dari serat nabati, jerami atau dari serutan, kepingan, partikel, serbuk gergaji atau sisa lainny dari kayu, diaglomerasi dengan semen, plester atau pengikat mineral lainnya.')," & _
"('0','680900','Barang dari plester atau dari komposisinya dengan dasar plester.')," & _
"('0','681000','Barang dari semen, beton atau batu artifisial, diperkuat maupun tidak.')," & _
"('0','681100','Barang dari semen asbes, dari semen serat selulosa atau sejenisnya.')," & _
"('0','681200','Serat asbes pabrikasi campuran dengan dasar dari asbes atau dengan dasar dari asbes dan magnesium karbonat barang dari campuran semacam itu atau dari asbes (misalny benang jahit, kain tenunan, pakaian, tutup kepal alas kaki, gasket), diperkuat maupun tidak, selain barang dari pos 68.11 atau 68.13.')," & _
"('0','681300','Bahan friksi dan barang yang terbuat dari padanya (misalny lembaran, gulungan, strip, segmen, cakram, cincin pipih, bantalan), tidak terpasang, untuk rem, kopling atau sejenisny dengan dasar asbes, bahan mineral lain atau selulos dikombinasi dengan tekstil atau bahan lain maupun tidak.')," & _
"('0','681400','Mika dikerjakan dan barang dari mik termasuk mika diaglomerasi atau dibentuk kembali, di atas dasar kertas, kertas karton atau bahan lainnya maupun tidak.')," & _
"('0','681500','Barang dari batu atau dari bahan mineral lainnya (termasuk serat karbon, barang dari serat karbon dan barang dari gambut), tidak dirinci atau termasuk dalam pos lainnya.')," & _
"('0','690000','Produk keramik')," & _
"('0','690100','Batu bat blok, ubin dan barang keramik lainnya dari tanah diatomea (misalny kieselguhr, tripolite atau diatomite) atau dari tanah mengandung silika semacam itu.')," & _
"('0','690200','Batu bat blok, ubin dan keramik tahan panas untuk barang keperluan konstruksi, selain dari tanah diatomea atau dari tanah mengandung silika semacam itu.')," & _
"('0','690300','Barang keramik tahan panas lainnya (misalny retor, cawan lebur, mofel, mulut pip sumbat, penopang, cangkir lebur, pembuluh, pip sarung, batang kecil dan gerbang geser), selain barang dari tanah diatomea atau tanah mengandung silika semacam itu.')," & _
"('0','690400','Batu bata bangunan, blok lantai, ubin penopang atau pengisi dan sejenisnya dari keramik.')," & _
"('0','690500','Ubin atap, cerobong berbentuk kap, tutup cerobong, lapisan cerobong, ornamen arsitektur dan barang keramik lainnya untuk konstruksi.')," & _
"('0','690600','Pip saluran, talang dan alat kelengkapan pipa dari keramik.')," & _
"('0','690700','Ubin dan paving, ubin perapian dan ubin dinding dari keramik kubus mozaik dari keramik dan sejenisny dengan alas maupun tidak keramik untuk finishing.')," & _
"('0','690900','Barang keramik untuk laboratorium, kimia atau penggunaan teknik lainnya palung keramik, pasu dan wadah semacam itu dari jenis yang digunakan dalam pertanian pot keramik, tempayan dan barang semacam itu dari jenis digunakan untuk mengangkut atau mengepak barang.')," & _
"('0','691000','Bak cuci, wastafel, alas baskom cuci, bak mandi, bidet, bejana kloset, tangki air pembilasan, tempat kencing, dan perlengkapan saniter semacam itu dari keramik.')," & _
"('0','691100','Perangkat makan, perangkat dapur, peralatan rumah tangga lainnya dan peralatan toilet, dari porselin atau keramik cina.')," & _
"('0','691200','Perangkat makan, perangkat dapur, peralatan rumah tangga lainnya dan peralatan toilet dari keramik, selain dari porselin atau keramik cina.')," & _
"('0','691300','Patung dan barang keramik ornamental lainnya.')," & _
"('0','691400','Barang keramik lainnya.')," & _
"('0','700000','Kaca dan barang dari kaca')," & _
"('0','700100','Pecahan dan sisa serta skrap lainnya dari kac kecuali kaca dari tabung sinar katoda atau kaca aktif lainnya dari pos 85.49 kaca dalam bentuk butiran.')," & _
"('0','700200','Kaca dalam bentuk bola (selain microspheres dari pos 70.18), batang kecil atau tabung, tidak dikerjakan.')," & _
"('0','700300','Kaca tuang dan rolled glass, dalam lembaran atau profil, mempunyai lapisan penyerap, pemantul atau bukan pemantul maupun tidak, tetapi tidak dikerjakan secara lain.')," & _
"('0','700400','Kaca tarik dan kaca tiup, dalam lembaran, mempunyai lapisan penyerap, pemantul atau bukan pemantul maupun tidak, tetapi tidak dikerjakan secara lain.')," & _
"('0','700500','Kaca apung dan kaca yang permukaannya digosok atau dipoles, dalam lembaran, mempunyai lapisan penyerap, pemantul atau bukan pemantul maupun tidak, tetapi tidak dikerjakan secara lain.')," & _
"('0','700600','Kaca dari pos 70.03, 70.04 atau 70.05, dibengkokkan, tepinya dikerjakan, diukir, dibor, dilapisi atau dikerjakan secara lain, tetapi tidak dibingkai atau dipasang dengan barang lain.')," & _
"('0','700700','Kaca pengaman, terdiri dari kaca dikeraskan (tempered) atau dilaminasi.')," & _
"('0','700800','Kaca dinding isolasi berlapis dari kaca.')," & _
"('0','700900','Cermin kac dibingkai maupun tidak, termasuk kaca spion.')," & _
"('0','701000','Carboy, botol, termos, kendi, teko, vial, ampul dan kemasan lainnya dari kac dari jenis yang digunakan untuk pengangkutan atau pengemasan barang kendi kaca pengawet sumbat, tutup dan penutup lainny dari kaca.')," & _
"('0','701100','Sungkup kaca (termasuk bulb dan tabung), terbuk dan bagiannya dari kac tanpa alat kelengkapan, untuk lampu listrik dan sumber cahay tabung sinar katoda atau sejenisnya.')," & _
"('0','701300','Barang kaca dari jenis yang digunakan untuk mej dapur, toilet, kantor, dekorasi dalam ruangan atau keperluan semacam itu (selain yang disebut dalam pos 70.10 atau 70.18).')," & _
"('0','701400','Barang kaca pemberi sinyal dan elemen optik dari kaca (selain yang dimaksud dalam pos 70.15), tidak dikerjakan secara optik.')," & _
"('0','701500','Kaca jam atau kaca arloji dan kaca semacam itu, kaca untuk kacamata korektif atau non-korektif, bergelombang, bengkok, berongga atau sejenisny tidak dikerjakan secara optik bola kaca berongga dan bagianny untuk pembuatan barang kaca tersebut di atas.')," & _
"('0','701600','Paving blok, lembaran tebal, batu bat bujur sangkar, ubin dan barang lainnya dari kaca padat atau kaca tuang, dipasang kawat maupun tidak, dari jenis yang digunakan untuk keperluan bangunan atau konstruksi kaca kubus dan barang kaca kecil lainny dengan alas maupun tidak, untuk mosaik atau keperluan dekorasi semacam itu kaca lapis timbal dan sejenisnya kaca multi seluler atau kaca busa dalam bentuk blok, panel, pelat, selongsong atau bentuk semacam itu.')," & _
"('0','701700','Barang kaca untuk keperluan laboratorium, hegienis atau farmasi, dengan tanda skala atau takaran maupun tidak.')," & _
"('0','701800','Manik-manik, mutiara imitasi, batu mulia imitasi atau batu semi mulia imitasi dan barang kecil semacam itu dari kac dan barang daripadanya selain perhiasan imitasi mata kaca selain mata pengganti untuk manusia patung dan ornamen lainnya dari kaca tiup, selain perhiasan imitasi kaca microsphere dengan diameter tidak melebihi 1 mm.')," & _
"('0','701900','Manik-manik, mutiara imitasi, batu mulia imitasi atau batu semi mulia imitasi dan barang kecil semacam itu dari kac dan barang daripadanya selain perhiasan imitasi mata kaca selain mata pengganti untuk manusia patung dan ornamen lainnya dari kaca tiup, selain perhiasan imitasi kaca microsphere dengan diameter tidak melebihi 1 mm.')," & _
"('0','702000','Barang lainnya dari kaca.')," & _
"('0','710000','Mutiara alam atau mutiara budiday batu mulia atau batu semi muli logam muli logam yang dipalut dengan logam muli dan barang daripadanya perhiasan imitasi koin')," & _
"('0','710100','Mutiar alam atau budiday dikerjakan atau ditingkatkan mutunya maupun tidak, tetapi tidak diuntai, tidak dipasang atau tidak disusun mutiar alam atau budiday diuntai sementara untuk memudahkan pengangkutan.')," & _
"('0','710200','Intan, dikerjakan maupun tidak, tetapi tidak dipasang atau tidak disusun.')," & _
"('0','710300','Batu mulia (selain intan) dan batu semi muli dikerjakan atau ditingkatkan mutunya maupun tidak tetapi tidak diuntai, tidak dipasang atau tidak disusun batu mulia yang tidak ditingkatkan mutunya (selain intan) dan batu semi muli diuntai sementara untuk memudahkan pengangkutan.')," & _
"('0','710400','Batu mulia atau semi mulia sintetik atau direkonstruksi, dikerjakan atau ditingkatkan mutunya maupun tidak tetapi tidak diuntai, tidak dipasang atau tidak disusun batu mulia atau semi mulia sintetik atau direkonstruksi, tidak ditingkatkan mutuny diuntai sementara untuk memudahkan pengangkutan.')," & _
"('0','710500','Debu dan bubuk dari batu mulia atau semi muli alam atau sintetik.')," & _
"('0','710600','Perak (termasuk perak disepuh emas atau platina), tidak ditempa atau dalam bentuk setengah jadi, atau dalam bentuk bubuk.')," & _
"('0','710700','Logam tidak mulia dipalut dengan perak, tidak dikerjakan lebih lanjut selain menjadi barang setengah jadi.')," & _
"('0','710800','Emas (termasuk emas disepuh dengan platina) tidak ditempa atau dalam bentuk setengah jadi, atau dalam bentuk bubuk.')," & _
"('0','710900','Logam tidak mulia atau perak, dipalut dengan emas, tidak dikerjakan lebih lanjut selain setengah jadi.')," & _
"('0','711000','Platin tidak ditempa atau dalam bentuk setengah jadi, atau dalam bentuk bubuk.')," & _
"('0','711100','Logam tidak muli perak atau emas, dipalut dengan platin tidak dikerjakan lebih lanjut selain setengah jadi.')," & _
"('0','711200','Sisa dan skrap dari logam mulia atau dari logam yang dipalut dengan logam mulia sisa dan skrap lainnya mengandung logam mulia atau senyawa logam muli dari jenis yang digunakan terutama untuk pemulihan logam mulia selain barang yang dimaksud dalam pos 85.49.')," & _
"('0','711300','Barang perhiasan dan bagianny dari logam mulia atau dari logam yang dipalut dengan logam mulia.')," & _
"('0','711400','Barang hasil tempaan pandai emas atau pandai perak dan bagianny dari logam mulia atau dari logam yang dipalut dengan logam mulia.')," & _
"('0','711500','Barang lainnya dari logam mulia atau dari logam yang dipalut dengan logam mulia.')," & _
"('0','711600','Barang dari mutiara alam atau mutiara budiday batu mulia atau batu semi mulia (alam, sintetik atau direkonstruksi).')," & _
"('0','711700','Perhiasan imitasi.')," & _
"('0','711800','Koin.')," & _
"('0','720000','Besi dan baja')," & _
"('0','720100','Besi pig dan besi cermin (spiegeleisen) dalam bentuk pig, blok atau bentuk asal lainnya.')," & _
"('0','720200','Paduan fero.')," & _
"('0','720300','Produk fero diperoleh dengan reduksi langsung dari bijih besi dan produk besi spon lainny dalam bentuk bongkah, pelet atau semacamnya besi yang mempunyai kemurnian sekurang-kurangnya 99,94 % menurut beratny dalam bentuk bongkah, pelet atau semacamnya.')," & _
"('0','720400','Sisa dan skrap fero ingot hasil peleburan kembali skrap besi atau baja.')," & _
"('0','720500','Butir dan bubuk, dari besi pig, besi cermin (spiegeleisen), besi atau baja.')," & _
"('0','720600','Besi dan baja bukan paduan dalam bentuk ingot atau bentuk asal lainnya (tidak termasuk besi dari pos 72.03).')," & _
"('0','720700','Produk setengah jadi dari besi atau baja bukan paduan.')," & _
"('0','720800','Produk canai lantaian dari besi atau baja bukan paduan, dengan lebar 600 mm atau lebih, dicanai panas, tidak dipalut, tidak disepuh atau tidak dilapisi.')," & _
"('0','720900','Produk canai lantaian dari besi atau baja bukan paduan, dengan lebar 600 mm atau lebih, dicanai dingin (cold-reduced), tidak dipalut, tidak disepuh atau tidak dilapisi.')," & _
"('0','721000','Produk canai lantaian dari besi atau baja bukan paduan, dengan lebar 600 mm atau lebih, dipalut, disepuh atau dilapisi.')," & _
"('0','721100','Produk canai lantaian dari besi atau baja bukan paduan, dengan lebar kurang dari 600 mm, tidak dipalut, tidak disepuh atau tidak dilapisi.')," & _
"('0','721200','Produk canai lantaian dari besi atau baja bukan paduan, dengan lebar kurang dari 600 mm, dipalut, disepuh atau dilapisi.')," & _
"('0','721300','Batang dan batang kecil, dicanai panas, dalam gulungan yang putarannya tidak beraturan, dari besi atau baja bukan paduan.')," & _
"('0','721400','Batang dan batang kecil lainnya dari besi atau baja bukan paduan, tidak dikerjakan lebih lanjut selain ditemp dicanai panas, ditarik panas atau diekstrusi panas, termasuk yang dipuntir setelah dicanai.')," & _
"('0','721500','Batang dan batang kecil lainnya dari besi atau baja bukan paduan.')," & _
"('0','721600','Angle, shape dan section dari besi atau baja bukan paduan.')," & _
"('0','721700','Kawat besi atau baja bukan paduan.')," & _
"('0','721800','Baja stainless berbentuk ingot atau bentuk asal lainnya produk setengah jadi dari baja stainless.')," & _
"('0','721900','Produk canai lantaian dari baja stainless, dengan lebar 600 mm atau lebih.')," & _
"('0','722000','Produk canai lantaian dari baja stainless, dengan lebar kurang dari 600 mm.')," & _
"('0','722100','Batang dan batang kecil, dicanai panas, dalam gulungan yang putarannya tidak beraturan, dari baja stainless.')," & _
"('0','722200','Batang dan batang kecil lainnya dari baja stainless angle, shape dan section dari baja stainless.')," & _
"('0','722300','Kawat dari baja stainless.')," & _
"('0','722400','Baja paduan lainnya dalam bentuk ingot atau bentuk asal lainnya produk setengah jadi dari baja paduan lainnya.')," & _
"('0','722500','Produk canai lantaian dari baja paduan lainny dengan lebar 600 mm atau lebih.')," & _
"('0','722600','Produk canai lantaian dari baja paduan lainny dengan lebar kurang dari 600 mm.')," & _
"('0','722700','Batang dan batang kecil, dicanai panas, dalam gulungan yang putarannya tidak beraturan, dari baja paduan lainnya.')," & _
"('0','722800','Batang dan batang kecil lainnya dari baja paduan lainnya angle, shape, dan section dari baja paduan lainnya batang dan batang kecil bor berongg dari baja paduan atau baja bukan paduan.')," & _
"('0','722900','Kawat dari baja paduan lainnya.')," & _
"('0','730000','Barang dari besi atau baja')," & _
"('0','730100','Sheet piling dari besi atau baj dibor, dilobangi atau dibuat dari rakitan elemen maupun tidak angle, shape dan section dilas, dari besi atau baja.')," & _
"('0','730200','Bahan konstruksi rel kereta api atau trem dari besi atau baj sebagai berikut : rel, rel simpang dan rel gigi, bilah wesel, sambungan persimpangan, batang kecil wesel dan bagian persimpangan lainny bantalan rel (cross-ties), pelat apit, tumpuan, pasak tumpuan, pelat tapak (pelat dasar), jepit rel, pelat alas, ganjal rel dan barang lainnya khusus untuk menyambung atau memasang rel.')," & _
"('0','730300','Pembuluh, pipa dan profil berongg dari besi tuang.')," & _
"('0','730400','Pembuluh, pipa dan profil berongg tanpa kampuh, dari besi (selain besi tuang) atau baja.')," & _
"('0','730500','Pembuluh dan pipa lainnya (misalny dilas, dikeling, atau disambung semacam itu), mempunyai penampang silang lingkaran, diameter luarnya melebihi 406,4 mm, dari besi atau baja.')," & _
"('0','730600','Pembuluh, pipa dan profil berongga lainnya (misalny kampuh atau dilas terbuk dikeling atau sambungan semacam itu), dari besi atau baja.')," & _
"('0','730700','Alat kelengkapan pembuluh atau pipa (misalny penyambung, siku-siku, selongsong) dari besi atau baja.')," & _
"('0','730800','Struktur (tidak termasuk bangunan prapabrikasi dari pos 94.06) dan bagian dari struktur (misalny jembatan dan bagian jembatan, daun pintu air, menar tiang kisi-kisi, atap, rangka atap, pintu dan jendela serta rangkanya dan ambang untuk pintu, penutup jendel langkan pilar dan tiang), dari besi atau baja pelat, batang kecil, angle, shape, section, pembuluh dan sejenisny disiapkan untuk keperluan struktur, dari besi atau baja.')," & _
"('0','730900','Reservoir, tangki, tong dan tempat penyimpanan semacamnya untuk segala jenis bahan (selain gas dibawah tekanan atau gas dicairkan), dari besi atau baj dengan kapasitas melebihi 300 liter, dilapis atau diisolasi panas maupun tidak, tetapi tidak dilengkapi dengan perlengkapan mekanik atau termal.')," & _
"('0','731000','Tangki, tahang, drum, kaleng, kotak dan tempat penyimpanan semacam itu, untuk berbagai bahan (selain dari gas dibawah tekanan atau gas dicairkan), dari besi atau baj dengan kapasitas tidak melebihi 300 liter, dilapisi atau diisolasi panas maupun tidak, tetapi tidak dilengkapi dengan perlengkapan mekanik atau termal.')," & _
"('0','731100','Kemasan untuk gas dibawah tekanan atau gas dicairkan, dari besi atau baja.')," & _
"('0','731200','Kawat dipilin, tali, kabel, ban anyaman, sling dan sejenisny dari besi atau baj tidak diisolasi elektrik.')," & _
"('0','731300','Kawat berduri dari besi atau baja simpai dipuntir atau kawat pipih tunggal, berduri atau tidak, dan kawat rangkap dipilin secara longgar, dari jenis yang digunakan untuk pagar, dari besi atau baja.')," & _
"('0','731400','Kain (termasuk ban tanpa ujung), anyaman kisi, jaring dan pagar, dari kawat besi atau baja expanded metal dari besi atau baja.')," & _
"('0','731500','Rantai dan bagianny dari besi atau baja.')," & _
"('0','731600','Jangkar, jangkar kecil dan bagianny dari besi atau baja.')," & _
"('0','731700','Paku, pasak, paku payung, corrugated nail, paku kokot (selain yang dimaksud pos 83.05) dan barang semacam itu, dari besi atau baj dengan kepala dari bahan lain maupun tidak, tetapi tidak termasuk barang seperti itu dengan kepala dari tembaga.')," & _
"('0','731800','Sekrup, baut, mur, sekrup rel, kait sekrup, paku keling, pasak, pasak kunci, cincin pipih (termasuk cincin pipih pegas) dan barang semacam itu, dari besi atau baja.')," & _
"('0','731900','Jarum jahit, jarum rajut, jarum jelujur, jarum kait, jarum bordir dan barang semacam itu, digunakan dengan tangan, dari besi atau baja pin pengaman dan pin lainnya dari besi atau baj tidak dirinci atau termasuk dalam pos lainnya.')," & _
"('0','732000','Pegas dan daun untuk pegas, dari besi atau baja.')," & _
"('0','732100','Tungku, kompor, tungku terbuk alat masak (termasuk tungku dengan ketel tambahan untuk pemanasan sentral), panggangan besar, anglo, gelang gas, piring pemanas, dan peralatan rumah tangga tanpa listrik semacam itu, dan bagianny dari besi atau baja.')," & _
"('0','732200','Radiator untuk pemanasan sentral, tidak dipanaskan dengan listrik, dan bagianny dari besi atau baja pemanas udara dan distributor udara panas (termasuk distributor yang dapat juga mendistribusikan udara segar atau didinginkan), tidak dipanaskan dengan listrik, didalamnya dipasang kipas atau alat tiup yang dijalankan dengan motor, dan bagianny dari besi atau baja.')," & _
"('0','732300','Barang untuk penggunaan di atas mej di dapur atau peralatan rumah tangga lainnya dan bagianny dari besi atau baja wol besi atau wol baja penggosok pot dan bantalan gosok atau pemoles, sarung tangan dan sejenisny dari besi atau baja.')," & _
"('0','732400','Perangkat saniter dan bagianny dari besi atau baja.')," & _
"('0','732500','Barang tuangan lainnya dari besi atau baja.')," & _
"('0','732600','Barang lainnya dari besi atau baja.')," & _
"('0','740000','Tembaga dan barang daripadanya')," & _
"('0','740100','Mate tembaga tembaga semen (tembaga endapan).')," & _
"('0','740200','Tembaga tidak dimurnikan anoda tembaga untuk pemurnian secara elektrolisa.')," & _
"('0','740300','Tembaga dimurnikan dan paduan tembag tidak ditempa.')," & _
"('0','740400','Sisa dan skrap tembaga.')," & _
"('0','740500','Paduan utama dari tembaga.')," & _
"('0','740600','Bubuk dan serpih tembaga.')," & _
"('0','740700','Batang, batang kecil dan profil tembaga.')," & _
"('0','740800','Kawat tembaga.')," & _
"('0','740900','Pelat, lembaran dan strip tembag dengan ketebalan melebihi 0,15 mm.')," & _
"('0','741000','Foil tembaga (dicetak atau dengan alas kertas, kertas karton, plastik atau alas semacam itu maupun tidak) dengan ketebalan (tidak termasuk alasnya) tidak melebihi 0,15 mm.')," & _
"('0','741100','Pembuluh dan pipa tembaga.')," & _
"('0','741200','Alat kelengkapan pembuluh atau pipa tembaga (misalny penyambung, siku-siku, selongsong).')," & _
"('0','741300','Kawat dipilin, kabel, ban anyaman dan sejenisny dari tembag tidak diisolasi elektrik.')," & _
"('0','741500','Paku, tack, paku payung, kokot (selain barang yang dimaksud dalam pos 83.05) dan barang semacam itu, dari tembaga atau besi atau baja dengan kepala tembaga sekrup, baut, mur, kait sekrup, paku keling, pasak, pasak kunci, cincin pipih (termasuk cincin pipih pegas) dan barang semacam itu, dari tembaga.')," & _
"('0','741800','Barang untuk penggunaan di atas mej di dapur atau peralatan rumah tangga lainnya dan bagianny dari tembaga penggosok pot dan bantalan gosok atau poles, sarung tangan dan sejenisny dari tembaga perangkat saniter dan bagianny dari tembaga.')," & _
"('0','741900','Barang lainnya dari tembaga.')," & _
"('0','750000','Nikel dan barang daripadanya')," & _
"('0','750100','Mate nikel, sinter oksida nikel dan produk-antara lainnya dari metalurgi nikel.')," & _
"('0','750200','Nikel tidak ditempa.')," & _
"('0','750300','Sisa dan skrap nikel.')," & _
"('0','750400','Bubuk dan serpih nikel.')," & _
"('0','750500','Batang, batang kecil, profil dan kawat nikel.')," & _
"('0','750600','Pelat, lembaran, strip dan foil nikel.')," & _
"('0','750700','Pembuluh, pipa dan alat kelengkapan pembuluh atau pipa dari nikel (misalny penyambung, siku-siku, selongsong).')," & _
"('0','750800','Barang lainnya dari nikel.')," & _
"('0','760000','Aluminium dan barang daripadanya')," & _
"('0','760100','Aluminium tidak ditempa.')," & _
"('0','760200','Sisa dan skrap aluminium.')," & _
"('0','760300','Bubuk dan serpih aluminium.')," & _
"('0','760400','Batang, batang kecil dan profil aluminium.')," & _
"('0','760500','Kawat aluminium.')," & _
"('0','760600','Pelat, lembaran dan strip aluminium, dengan ketebalan melebihi 0,2 mm.')," & _
"('0','760700','Aluminium foil (dicetak atau diberi alas kertas, kertas karton, plastik atau alas semacam itu maupun tidak) dengan ketebalan (tidak termasuk alasnya) tidak melebihi 0,2 mm.')," & _
"('0','760800','Pembuluh dan pipa aluminium.')," & _
"('0','760900','Alat kelengkapan pembuluh atau pipa aluminium (misalny penyambung, siku-siku, selongsong).')," & _
"('0','761000','Struktur aluminium (tidak termasuk bangunan prapabrikasi dari pos 94.06) dan bagian dari struktur (misalny jembatan dan bagian jembatan, menar tiang kisi-kisi, atap, rangka atap, pintu dan jendela serta rangkanya dan ambang untuk pintu, langkan, pilar dan tiang) pelat, batang kecil, profil, pembuluh dan sejenisnya dari aluminium, disiapkan untuk keperluan struktur.')," & _
"('0','761100','Reservoir, tangki, tong dan tempat penyimpanan semacam itu, dari aluminium, untuk berbagai jenis bahan (selain gas di bawah tekanan atau gas cair), dengan kapasitas melebihi 300 l, dilapis atau diisolasi panas maupun tidak, tetapi tidak dilengkapi dengan perlengkapan mekanik atau termal.')," & _
"('0','761200','Tahang, drum, kaleng, kotak dan tempat penyimpanan semacam itu dari aluminium (termasuk tempat simpan berbentuk tabung kaku atau dapat dilipat), untuk berbagai jenis bahan (selain gas dibawah tekanan atau gas cair), dengan kapasitas tidak melebihi 300 l, dilapis atau diisolasi panas maupun tidak, tetapi tidak dilengkapi dengan perlengkapan mekanik atau termal.')," & _
"('0','761300','Tempat penyimpanan dari aluminium untuk gas dibawah tekanan atau gas cair.')," & _
"('0','761400','Kawat pilin, kabel, ban anyaman dan sejenisny dari aluminium, tidak diisolasi elektrik.')," & _
"('0','761500','Barang untuk penggunaan di atas mej di dapur atau peralatan rumah tangga lainnya dan bagianny dari aluminium penggosok pot dan bantalan gosok atau pemoles, sarung tangan dan sejenisny dari aluminium perangkat saniter dan bagianny dari aluminium.')," & _
"('0','761600','Barang lainnya dari aluminium.')," & _
"('0','780000','Timbal dan barang daripadanya')," & _
"('0','780100','Timbal tidak ditempa.')," & _
"('0','780200','Sisa dan skrap timbal.')," & _
"('0','780400','Pelat, lembaran, strip dan foil timbal bubuk dan serpih timbal.')," & _
"('0','780600','Barang lainnya dari timbal.')," & _
"('0','790000','Seng dan barang daripadanya')," & _
"('0','790100','Seng tidak ditempa.')," & _
"('0','790200','Sisa dan skrap seng.')," & _
"('0','790300','Debu, bubuk dan serpihan seng.')," & _
"('0','790400','Batang, batang kecil, profil dan kawat seng.')," & _
"('0','790500','Pelat, lembaran, strip dan foil seng.')," & _
"('0','790700','Barang lainnya dari seng.')," & _
"('0','800000','Timah dan barang daripadanya')," & _
"('0','800100','Timah tidak ditempa.')," & _
"('0','800200','Sisa dan skrap timah.');"

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetUrutan1() As String
        Dim retval As String = ""
        Try
            retval = "Insert Into `tblmaster_taxproduct_reference` (`type`,`Code`,`Description`) Values" & _
                                       "('0','000000','Barang')," & _
                                       "('0','010000','Binatang hidup')," & _
                                       "('0','010100','Kud keledai, bagal dan hinnie, hidup')," & _
                                       "('0','010200','Binatang hidup jenis lembu')," & _
                                       "('0','010300','Babi hidup')," & _
                                       "('0','010400','Biri-biri dan kambing, hidup')," & _
                                       "('0','010500','Unggas hidup, yaitu ayam dari spesies Gallus domesticus, bebek, angs kalkun dan ayam guinea')," & _
                                       "('0','010600','Binatang hidup lainnya')," & _
                                       "('0','020000','Daging dan sisa daging yang dapat dimakan')," & _
                                       "('0','020100','Daging binatang jenis lembu, segar atau dingin')," & _
                                       "('0','020200','Daging binatang jenis lembu, beku')," & _
                                       "('0','020300','Daging babi, segar, dingin atau beku')," & _
                                       "('0','020400','Daging bid-bid atau kamblng, segar, dingin atau beku')," & _
                                       "('0','020500','Daging kud keledai, bagal atau hinnie, segar, dingin atau beku')," & _
                                       "('0','020600','Sisa yang dapat dimakan dari binatang jenis lembu, babi, biri-biri, kambing, kud keledai, bagal atau hinnie, segar, dingin atau beku.')," & _
                                       "('0','020700','Daging dan sisanya yang dapat dimakan, dari unggas dari pos 01.05, segar, dingin atau beku')," & _
                                       "('0','020800','Daging dan sisanya yang dapat dimakan dari binatang lainny segar, dingin atau beku')," & _
                                       "('0','020900','Lemak babi tanpa daging dan lemak unggas, tidak dicairkan atau diekstraksi dengan cara lain, segar, dingin, beku, dlasinkan, dalam air garam, dikeringkan atau diasapi')," & _
                                       "('0','021000','Daging dan sisanya yang dapat dimakan, diasinkan, dalam air garam, dikeringkan atau diasapi tepung dan tepung kasar dari daging dan sisanya yang dapat dimakan')," & _
                                       "('0','030000','lkan dan krustase moluska serta invertebrata air lainnya')," & _
                                       "('0','030100','Ikan hidup')," & _
                                       "('0','030200','Ikan, segar atau dingin, tidak termasuk potongan ikan tanpa tulang dan daging ikan lainnya dari pos 03.04')," & _
                                       "('0','030300','Ikan beku, tidak termasuk potongan ikan tanpa tulang dan daging ikan lainnya dari pos 03.04')," & _
                                       "('0','030400','Fillet dan daging ikan lainnya (dicincang maupun tidak), segar, dlngln atau beku')," & _
                                       "('0','030500','lkan, dikeringkan, diasinkan atau dalam air garam lkan diasapi, dimasak maupun tidak sebelum atau selama proses pengasapan')," & _
                                       "('0','030600','Krustase bercangkang maupun tidak, hidup, segar, dingin, beku, dikeringkan, diasinkan atau dalam air garam krustasea diasapi, dimasak maupun tidak sebelum atau selama proses pengasapan krustase bercangkang, dikukus atau direbus, dingin, beku., dilr:eringkan, diasinkan atau dalam air garam maupun tidak')," & _
                                       "('0','030700','Molusk bercangkang maupun tidak, hidup, segar, dingin, beku, dikeringkan, diasinkan atau dalam air garam moluska diasapi, bercangkang maupun tidak, dimasak maupun tidak sebelum atau selama proses pengasapan')," & _
                                       "('0','030800','Invertebrata air selain krustasea dan molusk hidup, segar, dingin, beku, dikeringkan, diasinkan atau dalam air garam Invertebrata air selain krustasea dan moluska diasapi, dimasak maupun tidak sebelum atau selama proses pengasapan')," & _
                                       "('0','030900','Tepung, tepung kasar dan pellet dari ikan atau krustase moluska atau invertebrata air lainny layak untuk dikonsumsi manusia')," & _
                                       "('0','040000','Produk susu telur unggas madu alam produk hewani yang dapat dimakan, tidak dirinci atau termasuk dalam pos lain')," & _
                                       "('0','040100','Susu dan kepala susu, tidak dipekatkan maupun tidak mengandung tambahan gula atau bahan pemanis lainnya')," & _
                                       "('0','040200','Susu dan kepala susu, dipekatkan atau mengandung tambahan gula atau bahan pemanis lainnya')," & _
                                       "('0','040300','Yoghurt susu menteg susu dan kepala susu dikentalkan, kefir dan susu dan krim difermentasi atau diasamkan lainny dipekatkan atau mengandung tambahan gula atau bahan pemanis lainnya atau diberi rasa atau mengandung tambahan buah-buahan, kacang-kacangan atau kakao maupun tidak')," & _
                                       "('0','040400','Whey, dipekatkan atau mengandung tambahan gula atau bahan pemanis lainnya maupun tidak produk terdiri dari susu alam sebagai unsur utam mengandung tambahan gul bahan pemanis lainnya maupun tidak, tidak dirinci atau termasuk dalam pos lainnya')," & _
                                       "('0','040500','Mentega dan lemak serta minyak lainnya yang diperoleh dari susu dairy spreads')," & _
                                       "('0','040600','Keju dan dadih susu')," & _
                                       "('0','040700','Telur unggas bercangkang, segar, diawetkan atau dimasak')," & _
                                       "('0','040800','Telur unggas, tanpa cangkang, dan kuning telur, segar, dikeringkan, dikukus atau direbus, dibentuk, beku atau diawetkan secara lain, mengandung tambahan gula atau bahan pemanis lainnya maupun tidak')," & _
                                       "('0','040900','Madu alam')," & _
                                       "('0','041000','Serangga dan produk lainnya yang dapat dimakan berasal dari hewan, tidak dirinci atau termasuk dalam pos lainnya')," & _
                                       "('0','050000','Produk hewani, tidak dirinci atau termasuk dalam pos lain')," & _
                                       "('0','050100','Rambut manusi tidak dikerjakan, dicuci atau digosok maupun tidak sisa rambut manusia.')," & _
                                       "('0','050200','Bulu dan bulu kasar dari babi, babi ternak atau babi hutan bulu berang-berang dan bulu binatang lainnya yang dapat dibuat sikat sisa dari bulu atau bulu kasar semacam itu.')," & _
                                       "('0','050400','Usus, kandung kemih dan lambung binatang (selain ikan), utuh dan potonganny segar, dingin, beku, diasinkan, dalam air garam, kering atau diasapi.')," & _
                                       "('0','050500','Kulit dan bagian lainnya dari unggas, masih berbulu atau berbulu halus, bulu unggas dan bagiannya (pinggirannya dipangkas maupun tidak) dan bulu halus, tidak dikerjakan lebih lanjut selain dibersihkan, disucihamakan atau dikerjakan untuk pengawetan bubuk dan sisa dari bulu atau bagiannya.')," & _
                                       "('0','050600','Tulang dan teras tanduk, tidak dikerjakan, dihilangkan lemakny dikerjakan secara sederhana (tetapi tidak dipotong menjadi berbentuk), dikerjakan dengan asam atau dihilangkan gelatinnya bubuk dan sisa dari produk tersebut.')," & _
                                       "('0','050700','Gading, tempurung kura-kur whalebone dan whalebone hair, tanduk, ranggah, kuku (binatang sejenis kuda atau sapi), kuku burung, cakar burung dan paruh burung, tidak dikerjakan atau dikerjakan secara sederhana tetapi tidak dipotong menjadi berbentuk bubuk dan sisa dari produk tersebut.')," & _
                                       "('0','050800','Koral dan barang serupa itu, tidak dikerjakan atau dikerjakan secara sederhana tetapi tidak dikerjakan lebih lanjut cangkang molusk krustasea atau echinodermata dan cuttle-bone, tidak dikerjakan atau dikerjakan secara sederhana tetapi tidak dipotong menjadi berbentuk, bubuk dan sisanya.')," & _
                                       "('0','050900','Ambar, kastor, jebat dan kesturi kantaridi empedu, kering maupun tidak kelenjar dan produk binatang lainnya yang digunakan dalam olahan produk farmasi, segar, dingin, beku atau diawetkan sementara secara lain')," & _
                                       "('0','051000','Ambar, kastor, jebat dan kesturi kantaridi empedu, kering maupun tidak kelenjar dan produk binatang lainnya yang digunakan dalam olahan produk farmasi, segar, dingin, beku atau diawetkan sementara secara lain.')," & _
                                       "('0','051100','Produk hewani tidak dirinci atau termasuk dalam pos lain binatang mati dari Bab 1 atau 3, tidak layak untuk dikonsumsi manusia.')," & _
                                       "('0','060000','Pohon hidup dan tanaman lainnya umbi, akar dan sejenisnya bunga potong dan daun ornamen')," & _
                                       "('0','060100','Umbi, bonggol, akar berbonggol, batang dibawah tanah, tajuk dan akar tongkat, dorman, sedang tumbuh atau berbunga tanaman dan akar chicory selain akar dari pos 12.12')," & _
                                       "('0','060200','Tanaman hidup lainnya (termasuk akarnya), potongan dan cangkokan sulur jamur')," & _
                                       "('0','060300','Bunga dan kuncup bunga potong dari jenis yang cocok untuk karangan bunga atau untuk keperluan pajangan, segar, dikeringkan, dicelup, dikelantang, diresapi, atau dikerjakan secara lain')," & _
                                       "('0','060400','Daun, dahan dan bagian lainnya dari tanaman, tanpa bunga atau kuncup bung dan rumput, lumut mosse dan lumut lichen, dari jenis yang cocok untuk karangan bunga atau keperluan pajangan, segar, dikeringkan, dicelup, dikelantang, diresapi atau diolah secara lain')," & _
                                       "('0','070000','Sayuran dan akar serta bonggol tertentu yang dapat dimakan')," & _
                                       "('0','070100','Kentang, segar atau dingin')," & _
                                       "('0','070200','Tomat, segar atau dingin')," & _
                                       "('0','070300','Bawang bombay, bawang merah, bawang putih, bawang bakung/perai dan sayuran sejenis lainny segar atau dingin')," & _
                                       "('0','070400','Kubis, bunga kol, kohlrabi, kale dan brassica sejenis yang dapat dimakan, segar atau dingin')," & _
                                       "('0','070500','Selada {Lactuca sativa) dan chicory (Cichorium spp.), segar atau dingin')," & _
                                       "('0','070600','Wortel, lobak cin akar bit untuk salad, salsify, celeriac, lobak dan akar sejenis yang dapat dimakan, segar atau dingin')," & _
                                       "('0','070700','Ketimun dan ketimun acar, segar atau dingin')," & _
                                       "('0','070800','Sayuran polongan, dikupas atau tidak, segar atau dingin')," & _
                                       "('0','070900','Sayuran lainny segar atau dingin')," & _
                                       "('0','071000','Sayuran (tidak dimasak atau dimasak dengan dikukus atau direbus), beku')," & _
                                       "('0','071100','Sayuran yang diawetkan sementar tetapi tidak cocok untuk konsumsi langsung')," & _
                                       "('0','071200','Sayuran dikeringkan, utuh, potongan, irisan, patahan atau dalam bentuk bubuk, tetapi tidak diolah lebih lanjut')," & _
                                       "('0','071300','Sayuran polongan dikeringkan, dikupas, dikuliti atau dibelah maupun tidak')," & _
                                       "('0','071400','Ubi kayu, arrowroot, salep, Jerusalem artichokes, ubi jalar serta akar-akaran dan bonggol-bonggolan semacam itu yang mengandung banyak pati atau inulin, segar, dingin, beku atau dikeringkan, dalam bentuk irisan maupun tidak atau dalam bentuk pelet empulur sagu')," & _
                                       "('0','080000','Buah dan buah bertempurung yang dapat dimakan kulit dari buah jeruk atau melon')," & _
                                       "('0','080100','Kelap kacang Brasil dan kacang mede, segar atau dikeringkan, dihilangkan tempurungnya atau dikupas maupun tidak')," & _
                                       "('0','080200','Buah bertempurung lainny segar atau dikeringkan, dihilangkan tempurungnya atau dikupas maupun tidak')," & _
                                       "('0','080300','Pisang, termasuk pisang yang tidak cocok dikonsumsi langsung sebagai buah, segar atau dikeringkan')," & _
                                       "('0','080400','Korm buah ar nanas, alpokat, jambu, mangg dan manggis, segar atau dikeringkan')," & _
                                       "('0','080500','Buah jeruk, segar atau dikeringkan')," & _
                                       "('0','080600','Anggur, segar atau dikeringkan')," & _
                                       "('0','080700','Melon (termasuk semangka) dan papaw (pepaya), segar')," & _
                                       "('0','080800','Apel, pir dan quince, segar')," & _
                                       "('0','080900','Aprikot, ceri, persik (termasuk nektarln), plum dan sloe, segar')," & _
                                       "('0','081000','Buah lainny segar')," & _
                                       "('0','081100','Buah dan buah bertempurung, tidak dimasak atau dimasak dengan cara dikukus atau direbus, beku, mengandung tambahan gula atau bahan pemanis lainnya maupun tidak')," & _
                                       "('0','081200','Buah dan buah bertempurung diawetkan sementar tetapi tidak cocok untuk konsumsi langsung')," & _
                                       "('0','081300','Buah, kering, selain yang disebut dalam pos 08.01 sampai dengan 08.06 campuran dari buah bertempurung atau buah kering dari Bab ini')," & _
                                       "('0','081400','Kulit buah jeruk atau melon (termasuk semangka), segar, beku, dikeringkan atau diawetkan sementara dalam air garam, air belerang atau dalam larutan pengawet lainnya')," & _
                                       "('0','090000','Kopi, teh, mate dan rempah-rempah')," & _
                                       "('0','090100','Kopi, digongseng atau dihilangkan kafeinnya maupun tidak sekam dan kulit kopi pengganti kopi mengandung kopi dengan perbandingan berapapun')," & _
                                       "('0','090200','Teh, diberl rasa maupun tidak')," & _
                                       "('0','090300','Mate')," & _
                                       "('0','090400','Lada dari genus Piper buah dari genus Capsicum atau dari genus Pimenta yang dikeringkan atau dihancurkan atau ditumbuk')," & _
                                       "('0','090500','Vanila')," & _
                                       "('0','090600','Kayu manis dan bunga kayu manis')," & _
                                       "('0','090700','Cengkeh (utuh, bunga dan tangkai).')," & _
                                       "('0','090800','Biji pale, bunga pala dan kapulaga')," & _
                                       "('0','090900','Biji adas manis, badian, adas pedas, ketumbar, jintan hitam atau jintan buah juniper')," & _
                                       "('0','091000','Jahe, saffron, turmeric (curcuma), thyme, daun salam, kari dan rempah-rempah lainnya')," & _
                                       "('0','100000','Serealia')," & _
                                       "('0','100100','Gandum dan meslin')," & _
                                       "('0','100200','Gandum hitam')," & _
                                       "('0','100300','Barli')," & _
                                       "('0','100400','Oat')," & _
                                       "('0','100500','Jagung')," & _
                                       "('0','100600','Beras')," & _
                                       "('0','100700','Butiran sorghum')," & _
                                       "('0','100800','Buckwheat, millet dan biji canary serealia lainnya.')," & _
                                       "('0','110000','Produk industri penggilingan malt pati inulin gluten gandum')," & _
                                       "('0','110100','Tepung gandum atau tepung meslin')," & _
                                       "('0','110200','Tepung serelia selain gandum atau meslin')," & _
                                       "('0','110300','Menir, tepung kasar dan pelet serelia')," & _
                                       "('0','110400','Butir serealia dikerjakan secara lain (misalny dikuliti, digiling, dipipihkan, dikikis, diiris atau dipecah), kecuali beras dari pos 10.06 lembaga sereali utuh, digiling, dipipihkan atau ditumbuk')," & _
                                       "('0','110500','Tepung, tepung kasar, bubuk, serpih, butir dan pelet kentang')," & _
                                       "('0','110600','Tepung, tepung kasar dan bubuk dari sayuran polongan dikeringkan dari pos 07.13, dari sagu atau dari akar atau bonggol dari pos 07.14 atau dari produk Bab 8.')," & _
                                       "('0','110700','Malt, digongseng maupun tidak')," & _
                                       "('0','110800','Pati inulin')," & _
                                       "('0','110900','Gluten gandum, dikeringkan maupun tidak')," & _
                                       "('0','120000','Biji dan buah mengandung minyak bermacam macam butir, biji dan buah tanaman industri atau tanaman obat jerami dan makanan ternak')," & _
                                       "('0','120100','Kacang kedelai, pecah maupun tidak')," & _
                                       "('0','120200','Kacang tanah, tidak digongseng atau dimasak secara lain, dihilangkan tempurungnya atau pecah maupun tidak')," & _
                                       "('0','120300','Kopra')," & _
                                       "('0','120400','Biji rami, pecah maupun tidak')," & _
                                       "('0','120500','Biji lobak atau colz pecah maupun tidak')," & _
                                       "('0','120600','Biji bunga matahari, pecah maupun tidak')," & _
                                       "('0','120700','Biji dan buah lainnya yang mengandung minyak, pecah maupun tidak')," & _
                                       "('0','120800','Tepung halus dan tepung kasar dari biji atau buah yang mengandung minyak, selain mustar')," & _
                                       "('0','120900','Biji, buah dan spor dari jenis yang digunakan untuk disemai')," & _
                                       "('0','121000','Kerucut buah hop, segar atau dikeringkan, ditumbuk, dijadikan bubuk atau dalam bentuk pelet maupun tidak lupulin')," & _
                                       "('0','121100','Tanaman dan bagiannya (termasuk biji dan buah), yang terutama dipakai dalam pembuatan wewangian, dalam farmasi atau untuk insektisid fungisida atau untuk tujuan yang semacam itu, segar, didinginkan, beku atau dikeringkan, baik dipotong, dihancurkan atau dijadikan bubuk maupun tidak')," & _
                                       "('0','121200','Kacang karob, rumput laut dan ganggang lainny bit gula dan tebu, segar, dingin, beku atau dikeringkan, ditumbuk maupun tldak kulit keras buah dan kernel serta produk nabati lainnya (termasuk akar chicory yang tidak digongseng dari varietas Cichorium intybus sativum) dari jenis yang terutama digunakan untuk konsumsi manusi tidak dirinci atau termasuk dalam pos lainnya')," & _
                                       "('0','121300','Jerami dan sekam sereali tidak diolah, balk dicacah, ditumbuk, ditekan atau dalam bentuk pelet maupun tidak')," & _
                                       "('0','121400','Swedes (sejenis lobak), mangold (sejenis bit), akar makanan ternak, rumput kering, lucerne (alfalfa), semanggi, sainfoin, Kale makanan ternak, lupine, Vetch dan produk makanan ternak sejenis lainny dalam bentuk pelet maupun tidak')," & _
                                       "('0','130000','Lak getah, damar dan sap serta ekstrak nabati lainnya')," & _
                                       "('0','130100','Lak getah alam, damar, getah-damar dan oleoresins (misalny getah balsem)')," & _
                                       "('0','130200','Sap dan ekstrak nabati zat pektik, pektinat dan pektat Agar-agar dan lendir serta bahan pengental lainny dimodifikasi maupun tidak, berasal dari produk nabati')," & _
                                       "('0','140000','Bahan anyaman nabati produk nabati tidak dirinci atau termasuk dalam pos lain')," & _
                                       "('0','140100','Bahan nabati dari jenis yang terutama dipakai untuk anyam-anyaman (misalny bambu, rotan, buluh, kumbuh, osier, rafi jerami serealia dibersibkan, dikelantang atau dicelup, dan kulit pohon limau)')," & _
                                       "('0','140400','Produk nabati tidak dirinci atau termasuk pos lainnya')," & _
                                       "('0','150000','Lemak dan minyak hewani atau nabati serta produk disosiasinya lemak olahan yang dapat dimakan malam hewani atau malam nabati')," & _
                                       "('0','150100','Lemak babi (termasuk lard) dan lemak unggas, selain dari pos 02.09 atau 15.03')," & _
                                       "('0','150200','Lemak dari binatang jenis lembu, biri-biri atau kambing, selain pos 15.03')," & _
                                       "('0','150300','Lard stearin, minyak lard, oleostearin, minyak oleo dan minyak tallow, tidak diemulsi atau dicampur atau diolah dengan cara lain')," & _
                                       "('0','150400','Lemak dan minyak serta fraksiny dari ikan atau binatang laut menyusui, dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','150500','Wool grease dan zat lemak turunannya (termasuk lanolin).')," & _
                                       "('0','150600','Lemak dan minyak binatang lainnya serta fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','150700','Minyak kacang kedelai dan fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','150800','Minyak kacang tanah dan fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','150900','Minyak zaitun dan fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','151000','Minyak lain dan fraksiny diperoleh semata-mata dari zaitun, dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimi termasuk campuran dari beberapa minyak atau fraksinya dengan minyak atau fraksi dari pos 15.09')," & _
                                       "('0','151100','Minyak kelapa sawit dan fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','151200','Minyak biji bunga matahari, safflower atau biji kapas dan fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','151300','Minyak kelapa (kopra), kernel kelapa sawit atau babassu dan fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikaid secara kimia')," & _
                                       "('0','151400','Minyak lobak, colza atau mustar dan fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','151500','Lemak dan minyak nabati atau mikroba tertentu lainnya (termasuk minyak jojoba) dan fraksiny dimurnikan maupun tidak, tetapi tidak dimodifikasi secara kimia')," & _
                                       "('0','151600','Lemak dan minyak hewani, nabati atau mikroba dan fraksiny sebagian atau seluruhnya dihidrogenasi, diinter-esterifikasi, dire-esterifikasi atau dielaidinisasi, dimurnikan maupun tidak, tetapi tidak diolah lebih lanjut')," & _
                                       "('0','151700','Margarin campuran atau olahan yang dapat dimakan dari lemak atau minyak hewani, nabati atau mikroba atau dari fraksi lemak atau minyak yang berbeda dalam bab ini, selain lemak dan minyak yang dapat dimakan atau fraksinya dari pos 15.16')," & _
                                       "('0','151800','Lemak dan minyak hewani, nabati atau mikroba serta fraksiny dipanaskan, dioksidasi, didehidrasi, disulfurisasi, ditiup, dipolimerisasi dengan panas dalam hampa udara atau dalam gas inert atau dimodifikasi secara kimia lainny tidak termasuk dari pos 15.16 olahan atau campuran yang tidak dapat dimakan dari lemak atau minyak hewani, nabati atau mlkroba atau dari fraksi lemak atau minyak yang berbeda dalam Bab ini, tidak dirinci atau termasuk dalam pos lainnya')," & _
                                       "('0','152000','Gliserol, mentah air gliserol dan larutan alkali gliserol')," & _
                                       "('0','152100','Malam nabati (selain trigliserida), malam lebah, malam serangga lainnya dan spermaceti, dimurnikan atau diwarnai maupun tidak')," & _
                                       "('0','152200','Degra residu yang diperoleh dari pengolahan zat berlemak atau malam hewani atau nabati')," & _
                                       "('0','160000','Olahan dari daging, dari ikan, dari krustase moluska atau invertebrata air lainnya')," & _
                                       "('0','160100','Sosis dan produk semacamny dari daging, sisa daging, darah atau serangga olahan makanan berasal dari produk ini')," & _
                                       "('0','160200','Daging, sisa daging, darah atau serangga lainnya yang diolah atau diawetkan')," & _
                                       "('0','160300','Ekstrak dan jus dari daging, ikan atau krustase moluska atau invertebrata air lainnya')," & _
                                       "('0','160400','Ikan diolah atau diawetkan kaviar dan pengganti kaviar yang diolah dari telur ikan')," & _
                                       "('0','160500','Krustase moluska dan invertebrata air lainny diolah atau diawetkan')," & _
                                       "('0','170000','Gula dan kembang gula')," & _
                                       "('0','170100','Gula tebu atau gula bit dan sukrosa murni Jdmiawi, dalam bentuk padat')," & _
                                       "('0','170200','Gula lainny termasuk laktos maltos glukosa dan fruktosa murni kimiawi, dalam bentuk padat sirop gula tidak mengandung tambahan bahan perasa atau pewarna madu artifisial, dicampur dengan madu alam maupun tidak karamel')," & _
                                       "('0','170300','Tetes hasil dari ekstraksi atau pemurnian gula')," & _
                                       "('0','170400','Kembang gula (termasuk coklat putih), tidak mengandung kakao')," & _
                                       "('0','180000','Kakao dan olahan kakao')," & _
                                       "('0','180100','Biji kakao, utuh atau pecah, mentah atau digongseng')," & _
                                       "('0','180200','Kulit, sekam, selaput dan sisa kakao lainnya')," & _
                                       "('0','180300','Pasta kakao, dihilangkan lemaknya maupun tidak')," & _
                                       "('0','180400','Menteg lemak dan minyak kakao')," & _
                                       "('0','180500','Bubuk kakao, tidak mengandung tambahan gula a.tau bahan pemanis lainnya')," & _
                                       "('0','180600','Coklat dan olahan makanan lainnya mengandung kakao.')," & _
                                       "('0','190000','Olahan dari sereali tepung, pati atau susu produk industri kue')," & _
                                       "('0','190100','Ekstrak malt olahan makanan dari tepung, menir, tepung kasar, pati atau ekstrak malt, tidak mengandung kakao atau mengandung kakao kurang dari 40 % menurut beratnya yang dihitung atas dasar kakao yang dihilangkan seluruh lemakny tidak dirinci atau termasuk pos lainnya olahan makanan dari pos 04.01 sampai dengan 04.04, tidak mengandung kakao atau mengandung kakao kurang dari 5 % menurut beratnya yang dihitung atas dasar kakao yang dihilangkan seluruh lemakny tidak dirinci atau termasuk pos lainnya')," & _
                                       "('0','190200','Past dimasak atau diisi maupun tidak (dengan daging atau bahan lainnya) atau diolah secara lain, seperti spagheti, makaroni, mie, lasagn gnocchi, ravioli, cannelloni couscous, diolah maupun tidak')," & _
                                       "('0','190300','Tapioka dan penggantinya diolah dari pati, dalam bentuk serpih, butir, pearls, hasll ayakan atau bentuk semacam itu')," & _
                                       "('0','190400','Makanan olahan diperoleh dengan cara menggembungkan atau menggongseng serealia atau produk serealia (misalny keripik jagung) serealia (selain jagung), dalam bentuk butir atau serpih atau butir yang dikerjakan secara lain (kecuali tepung, menir dan tepung kasar), belum dimasak atau diolah secara lain, tidak dirinci atau termasuk pos lainnya')," & _
                                       "('0','190500','Roti, kue kering, kue, biskuit dan produk roti lainny mengandung kakao maupun tidak wafer komuni, selongsong kosong dari jenis yang cocok untuk keperluan farmasi, sealing wafers, rice paper dan produk yang semacam itu.')," & _
                                       "('0','200000','Olahan dari sayufan, buah, biji/kacang atau bagian lain dari tanaman')," & _
                                       "('0','200100','Sayuran, buah, kacang dan bagian tanaman lainnya yang dapat dimakan, diolah atau diawetkan dengan cuka atau asam asetat')," & _
                                       "('0','200200','Tomat diolah atau diawetkan selain dengan cuka atau asam asetat')," & _
                                       "('0','200300','Jamur dan cendawan tanah, diolah atau diawetkan selain dengan cuka atau asam asetat')," & _
                                       "('0','200400','Sayuran lainnya yang diolah atau diawetkan selain dengan cuka atau asam asetat, beku, selain produk dari pos 20.06')," & _
                                       "('0','200500','Sayuran lainnya yang diolah atau diawetkan selain dengan cuka atau asam asetat, tidak beku, selain produk dari pos 20.06.')," & _
                                       "('0','200600','Sayuran, buah, kacang, kulit buah dan bagian lain dari tanaman, diawetkan dengan gula (dikeringkan, dilapisi atau dikristalisasi)')," & _
                                       "('0','200700','Selai, jell buah, marmelade, pure dan pasta dari buah atau kacang, diperoleh dari pemasakan, mengandung tambahan gula atau bahan pemanis lainnya maupun tidak')," & _
                                       "('0','200800','Buah, kacang dan bagian tanaman lainnya yang dapat dimakan, diolah atau diawetkan secara lain, mengandung tambahan gula atau bahan pemanis lainnya atau alkohol maupun tidak, tidak dirinci atau termasuk pos lainnya')," & _
                                       "('0','200900','Jus buah atau buah bertempurung (termasuk grape must dan air kelapa) dan jus sayuran, tidak difermentasi dan tidak mengandung tambahan alkohol, mengandung tambahan gula atau pemanis lainnya maupun tidak')," & _
                                       "('0','210000','Bermacam-macam olahan yang dapat dimakan')," & _
                                       "('0','210100','Ekstrak, esens dan konsentrat, dari kopi, teh atau mate dan olahan dengan dasar produk ini atau dengan dasar kopi, teh atau mate chicory digongseng dan pengganti kopi yang digongseng lainny dan ekstrak, esens dan konsentratnya')," & _
                                       "('0','210200','Ragi (aktif atau tidak aktif) mikro-organisme bersel tunggal lainny mati (tetapi tidak termasuk vaksin dari pos 30.02) bubuk pengembang roti')," & _
                                       "('0','210300','Saus dan olahannya campuran bahan penyedap dan campuran bumbu tepung mustar dan tepung kasar mustar serta mustar olahan')," & _
                                       "('0','210400','Sop dan kaldu serta olahannya olahan makanan campuran homogen')," & _
                                       "('0','210500','Es krim dan es lainnya yang dapat dimakan, mengandung kakao maupun tidak')," & _
                                       "('0','210600','Olahan makanan yang tidak dirinci atau termasuk dalam pos lainnya')," & _
                                       "('0','220000','Minuman, alkohol dan cuka')," & _
                                       "('0','220100','Air, termasuk air mineral alam atau artiflsial dan air sod tidak mengandung tambahan gula atau bahan pemanis lainnya maupun pemberi rasa es dan salju')," & _
                                       "('0','220200','Air, termasuk air mineral dan air sod mengandung tambahan gula atau bahan pemanis lainnya atau pemberi ras dan minuman yang tidak mengandung alkohol lainny tidak termasuk jus buah, buah bertempurung atau sayuran dari pos 20.09')," & _
                                       "('0','220300','Bir terbuat dari malt')," & _
                                       "('0','220400','Minuman fermentasi (wine) dari buah anggur segar, termasuk minuman fermentasi (wine) yang diperkuat grape must selain dari pos 20.09.')," & _
                                       "('0','220500','Vermouth dan minuman fermentasi lainnya dari buah anggur segar yang diberi rasa dengan zat nabati atau zat beraroma')," & _
                                       "('0','220600','Minuman fermentasi lainnya (misalny fermentasi buah apel, buah pir, larutan madu dalam air, sake) campuran minuman fermentasi dan campuran minuman fermentasi dengan minuman yang tidak mengandung alkohol, tidak dirinci atau termasuk dalam pos lainnya')," & _
                                       "('0','220700','Etil alkohol yang tidak didenaturasi dengan kadar alkohol 80 % atau lebih menurut volumenya etil alkohol dan alkohol lainny didenaturasi berapapun kadarnya')," & _
                                       "('0','220800','Etil alkohol yang tidak didenaturasi dengan kadar alkohol kurang dari 80 % menurut volumenya alkohol, sopi manis dan minuman beralkohol lainnya')," & _
                                       "('0','220900','Cuka dan pengganti cuka diperoleh dari asam asetat')," & _
                                       "('0','230000','Residu dan sisa dari industri makanan olahan makanan hewan')," & _
                                       "('0','230100','Tepung, tepung kasar dan pelet, dari daging atau sisany dari ikan atau krustase moluska atau invertebrata air lainny tidak layak untuk dikonkumsi manusia greaves')," & _
                                       "('0','230200','Sekam, dedak dan residu lainny berbentuk pelet maupun tidak, berasal dari pengayakan, penggilingan atau pengerjaan lainnya dari serealia atau dari tanaman polongan')," & _
                                       "('0','230300','Residu dari pembuatan pati dan residu semacam itu, pulp bit, ampas tebu dan sisa lainnya dari pembuatan gul endapan dan sisa pembuatan bir atau penyulingan, berbentuk pelet maupun tidak')," & _
                                       "('0','230400','Bungkil dan residu padat lainny ditumbuk maupun tidak atau berbentuk pelet, basil dari ekstraksi minyak kacang kedelai')," & _
                                       "('0','230500','Bungkil dan residu padat lainny ditumbuk maupun tidak atau berbentuk pelet, basil dari ekstraksi minyak kacang tanah')," & _
                                       "('0','230600','Bungkil dan residu padat lainny ditumbuk maupun tidak atau berbentuk pelet, basil dari ekstraksi lemak atau minyak nabati atau mikroba selain dari pos 23.04 atau 23.05')," & _
                                       "('0','230700','Endapan minuman fermentasi kerak minuman fermentasi')," & _
                                       "('0','230800','Bahan nabati dan sisa nabati, residu nabati dan hasil sampinganny dalam bentuk pelet maupun tidak, dari jenis yang digunakan untuk makanan hewan, tidak dirinci atau termasuk dalam pos lainnya')," & _
                                       "('0','230900','Olahan dari jenis yang digunakan untuk makanan hewan')," & _
                                       "('0','240000','Tembakau dan pengganti tembakau dipabrikasi')," & _
                                       "('0','240100','Tembakau belum dipabrikasi sisa tembakau')," & _
                                       "('0','240200','Cerutu, cheroot, cerutu kecil dan sigaret, dari tembakau atau pengganti tembakau')," & _
                                       "('0','240300','Tembakau dipabrlkasi lainnya dan pengganti tembakau dipabrikasi tembakau dihomogenisasi atau dibentuk kembali ekstrak dan esens tembakau')," & _
                                       "('0','240400','Produk yang mengandung tembakau, tembakau dibentuk kembali, nikotin, atau pengganti tembakau atau nikotin, dimaksudkan untuk inhalasi tanpa pembakaran produk mengandung nikotin lainnya yang dimaksudkan untuk memasukkan nikotin ke dalam tubuh manusia')," & _
                                       "('0','250000','Garam belerang tanah dan batu bahan pemlester, kapur dan semen')," & _
                                       "('0','250100','Garam (termasuk garam meja dan garam didenaturasi) dan natrium klorida murni, dalam larutan air atau mengandung tambahan bahan anti-caking atau free-flowing maupuntidak air laut')," & _
                                       "('0','250200','Pirit besi tidak digongseng')," & _
                                       "('0','250300','Belerang dari segala jenis, selain belerang sublimasi, belerang basil endapan dan belerang koloidal.')," & _
                                       "('0','250400','Grafit alam')," & _
                                       "('0','250500','Pasir alam dari segala jenis, berwarna maupun tidak, selain pasir mengandung logam dari Bab 26')," & _
                                       "('0','250600','Kuarsa (selain pasir alam) kuarsit, dikerjakan secara kasar atau semata-mata dipotong maupun tidak, digergaji atau dengan cara lainnya menjadi balok atau lembaran tebal berbentuk empat persegi panjang (termasuk bujur sangkar)')," & _
                                       "('0','250700','Kaolin dan tanah liat kaolin lainny dikalsinasi maupun tidak')," & _
                                       "('0','250800','Tanah liat lainnya (tidak termasuk tanah liat muai dari pos 68.06), andalusite, kyanite dan sillimanite, dikalsinasi maupun tidak mullite tanah chamotte atau tanah dinas')," & _
                                       "('0','250900','Kapur')," & _
                                       "('0','251000','Kalsium fosfat alam, aluminium kalsium fosfat alam dan kapur fosfat')," & _
                                       "('0','251100','Barium sulfat alam (barit) barium karbonat alam (witherite), dikalsinasi maupun tidak, selain barium oksida dari pos 28.16')," & _
                                       "('0','251200','Tanah diatomea (misalny kieselguhr, tripolite dan diatomit) dan tanah semacam itu yang mengandung silik dikalsinasi maupun tidak, dengan berat jenis sebesar 1 atau kurang')," & _
                                       "('0','251300','Batu apung amril korundum alam, garnet alam dan ampelas alam lainny diolah dengan cara dipanaskan maupun tidak')," & _
                                       "('0','251400','Batu sabak, dikerjakan secara kasar atau semata-mata dipotong maupun tidak, digergaji atau dengan cara lain, menjadi balok atau lembaran tebal berbentuk empat persegi panjang (termasuk bujur sangkar)')," & _
                                       "('0','251500','Marmer, travertine, ecaussine dan batu berkapur lainnya untuk monumen atau bangunan dengan berat jenis sebesar 2,5 atau lebih, dan alabaster, dlkerjakan secara kasar atau semata-mata dipotong maupun tidak, digergaji atau dengan cara lain, menjadi balok atau lembaran tebal berbentuk empat persegi panjang (termasuk bujur sangkar)')," & _
                                       "('0','251600','Granit, porphyry, basal, batu pasir dan batu lainnya untuk membuat monumen atau bangunan, dikerjakan secara kasar atau semata-mata dipotong maupun tidak, digergaji atau dengan cara lain, menjadi balok atau lembaran tebal berbentuk empat persegi panjang (termasuk bujur sangkar)')," & _
                                       "('0','251700','Kerikil, gravel, batu pecah atau batu tumbuk, yang lazim digunakan untuk campuran beton, untuk mengeraskan jalan atau untuk rel kereta api atau pemberat lainny shingle dan flint, diolah dengan dipanaskan maupun tidak makadam dari terak, sampah atau sisa buangan industri semacam itu, bercampur dengan bahan tersebut dalam bagian pertama pos ini maupun tidak makadam ter butir, keping dan bubuk, dari batuan dari pos 25.15 atau 25.16, diolah dengan dipanaskan maupun tidak')," & _
                                       "('0','251800','Dolomit, dikalsinasi atau disinter maupun tidak, termasuk dolomit yang dikerjakan secara kasar atau semata-mata dipotong, digergaji atau dengan cara lain, menjadi balok atau lembaran tebal berbentuk empat persegi panjang (termasuk bujur sangkar)')," & _
                                       "('0','251900','Magnesium karbonat alam (magnesit) magnesia leburan magnesia sinter, mengandung sedikit oksida lainnya yang ditambahkan sebelum disinter maupun tidak magnesium oksida lainny murni maupun tidak')," & _
                                       "('0','252000','Gips anhidrit plester (terdiri dari gips dikalsinasi atau kalsium sulfat) diwarnai maupun tidak, tanpa atau dengan sedlldt bahan akselerator atau retarder')," & _
                                       "('0','252100','Fluks batu kapur batu kapur dan batu. mengandung kapur lainny dari jenis yang digunakan untuk pembuatan kapur atau semen')," & _
                                       "('0','252200','Kapur tohor, kapur kembang dan kapur hidrolis, selain kalsium oksida dan hidroksida dari pos 28.25')," & _
                                       "('0','252300','Semen portland, semen alumin semen terak, semen super sulfat dan semen hidrolis semacam itu, berwarna maupun tidak atau dalam bentuk clinker')," & _
                                       "('0','252400','Asbes')," & _
                                       "('0','252500','Mik termasuk belahannya sisa mika')," & _
                                       "('0','252600','Steatit alam, dikerjakan secara kasar atau semata-mata dipotong maupun tidak, digergaji atau dengan cara lain, menjadi balok atau lembaran tebal berbentuk empat persegi panjang (termasuk bujur sangkar) talek')," & _
                                       "('0','252800','Borat slam dan konsentrataya (dikalsinasi maupun tidak), tetapi tidak termasuk borat yang dipisahkan dari air garam alam asam borat alam mergandung H3BO3, tidak lebih dari 85 % dihitung dari berat kering')," & _
                                       "('0','252900','Felspar leucite nepheline dan nepheline syenite fluorspar')," & _
                                       "('0','253000','Bahan mineral yang tidak dirinci atau termasuk dalam pos lainnya')," & _
                                       "('0','260000','Bijih logam, terak dan abu')," & _
                                       "('0','260100','Bijih besi dan konsentratny termasuk pirit besi digongseng')," & _
                                       "('0','260200','Bijih mangan dan konsentratny termasuk bljih mangan mengandung besi dan konsentratnya dengan kandungan mangan 20 % atau lebih, dihitung dari berat kering')," & _
                                       "('0','260300','Bijih tembaga dan konsentratnya')," & _
                                       "('0','260400','Bijih nikel dan konsentratnya.')," & _
                                       "('0','260500','Bijih kobalt dan konsentratnya')," & _
                                       "('0','260600','Bijih aluminium dan konsentratnya')," & _
                                       "('0','260700','Bijih timbal dan konsentratnya')," & _
                                       "('0','260800','Bijih seng dan konsentratnya')," & _
                                       "('0','260900','Bijih timah dan konsentratnya')," & _
                                       "('0','261000','Bijih kromium dan konsentratnya')," & _
                                       "('0','261100','Bijih tungsten dan konsentratnya')," & _
                                       "('0','261200','Bijih uranium atau torium dan konsentratnya')," & _
                                       "('0','261300','Bijih molibdenum dan konsentratnya')," & _
                                       "('0','261400','Bijih titanium dan konsentratnya')," & _
                                       "('0','261500','Bijih niobium, tantalum, vanadium atau zirkonium dan konsentratnya')," & _
                                       "('0','261600','Bijih logam mulia dan konsentratnya')," & _
                                       "('0','261700','Bijih lainnya dan konsentratnya')," & _
                                       "('0','261800','Terak butiran (pasir terak) dari industri pembuatan besi atau baja')," & _
                                       "('0','261900','Terak, sampah (selain terak butiran), kerak logam dan sisa lainnya dari pembuatan besi atau haja')," & _
                                       "('0','262000','Terak, abu dan residu (selain dari pemhuatan besi atau baja) mengandung logam, arsenik atau senyawanya')," & _
                                       "('0','262100','Terak logam dan abu lainny termasuk abu rumput laut (kelp) abu dan residu dari pembakaran limbah rumah tangga')," & _
                                       "('0','270000','Bahan bakar mineral, minyak mineral dan produk sulingannya zat mengandung bitumen malam mineral')," & _
                                       "('0','270100','Batu hara briket, ovoid dan bahan bakar padat semacam itu dibuat dari batu bara')," & _
                                       "('0','270200','Lignit, diaglomerasi maupun tidak, tidak termasuk jet')," & _
                                       "('0','270300','Gambut (termasuk serasah gambut), diaglomerasi maupun tidak.')," & _
                                       "('0','270400','Kokas dan semi-kokas dad batu bara dari lignit atau dari gambut, diaglomerasi maupun tidak retort karbon')," & _
                                       "('0','270500','Gas batu har gas air, gas produser dan gas semacam itu, selain gas minyak bumi dan gas hidrokarbon lainnya')," & _
                                       "('0','270600','Ter sulingan dari batu har dari lignit atau dari gambut, dan ter mineral lainny didehidrasi atau disuling sebagian maupun tidak, termasuk ter yang dibentuk kembali')," & _
                                       "('0','270700','Minyak dan produk lainnya basil penyulingan ter batu bara bersuhu tinggi produk semacam itu yang berat unsur aromatiknya melebihi unsur non aromatik')," & _
                                       "('0','270800','Pek dan pek kokas, diperoleh dari ter batu bara atau dari ter mineral lainnya')," & _
                                       "('0','270900','Minyak petroleum dan minyak yang diperoleh dari mineral mengandung bitumen, mentah')," & _
                                       "('0','271000','Minyak petroleum dan minyak yang diperoleh dari mineral mengandung bitumen, selain mentah preparat tidak dirinci atau termasuk dalam pos lainny mengandung minyak petroleum atau minyak yang diperoleh dari mineral mengandung bitumen 70 % atau lebih menurut beratny minyak ini merupakan unsur dasar dari preparat tersebut minyak sisa')," & _
                                       "('0','271100','Gas petroleum dan gas hidrokarbon lainnya')," & _
                                       "('0','271200','Jell petroleum malam parafin, malam petroleum mikrokristalin, slack wax, ozokerit, malam lignit, malam gambut, malam mineral lainny dan produk semacam itu yang diperoleh melalui sintesa atau proses lainny diwarnai maupun tidak')," & _
                                       "('0','271300','Kokas petroleum, bitumen petroleum dan residu lainnya dari minyak petroleum atau dari minyak yang diperoleh dari mineral mengandung bitumen')," & _
                                       "('0','271400','Bitumen dan aspal, alam shale mengandung bitumen atau minyak dan pasir ter asphaltite dan batu karang mengandung aspal')," & _
                                       "('0','271500','Campuran mengandung bitumen berasal dari aspal alam, dari bitumen alam, dari bitumen petroleum, dari ter mineral atau dari pek ter mineral (misalny mastik mengandung bitumen, cut-back)')," & _
                                       "('0','271600','Batu hara briket, ovoid dan bahan bakar padat semacam itu dibuat dari batu bara')," & _
                                       "('0','280000','Bahan kimia anorganik senyawa organik atau anorganik dari logam muli dari logam tanah jarang, dari unsur radioaktif atau dari isotop')," & _
                                       "('0','280100','Fluorin, klorin, bromin dan iodin')," & _
                                       "('0','280200','Belerang, basil sublimasi atau endapan belerang koloidal')," & _
                                       "('0','280300','Karbon (carbon black dan bentuk lain dari karbon tidak dirinci atau termasuk dalam pos lainnya)')," & _
                                       "('0','280400','Hidrogen, gas mulia dan non logam lainnya')," & _
                                       "('0','280500','Logam alkali atau logam alkali tanah logam tanah jarang, skandium dan itrium, dicampur atau dipadu maupun tidak merkuri')," & _
                                       "('0','280600','Hidrogen klorida (asam klorida) asam klorosulfat')," & _
                                       "('0','280700','Asam sulfat oleum')," & _
                                       "('0','280800','Asam nitrat asam sulfonitrat')," & _
                                       "('0','280900','Difosfor pentaoksida asam fosfat asam polifosfat, mempunyai rumus kimiat tertentu maupun tidak')," & _
                                       "('0','281000','Oksida boron asam borat')," & _
                                       "('0','281100','Asam anorganik: lainnya dan senyawa oksigen anorganik: dari bukan logam lainnya')," & _
                                       "('0','281200','Halida dan halida oksida dari bukan logam')," & _
                                       "('0','281300','Sulfida dari bukan logam fosfor trisulfida komersial')," & _
                                       "('0','281400','Amoni anhidrat atau dalam larutan air')," & _
                                       "('0','281500','Natrium hidroksida (soda api) kalium hidroksida (potas api) peroksida dari natrium atau kalium')," & _
                                       "('0','281600','Hidroksida dan peroksida dari magnesium oksid hidroksida dan peroksid dari strontium atau barium')," & _
                                       "('0','281700','Seng oksida seng peroksida')," & _
                                       "('0','281800','Korundum artifisial, mempunyai rumus kimia tertentu maupun tidak aluminium oksida aluminium hidroksida')," & _
                                       "('0','281900','Kromium oksida dan kromium hidroksida')," & _
                                       "('0','282000','Mangan oksida')," & _
                                       "('0','282100','Besi oksida dan besi hidroksida tanah warna mengandung senyawa besi yang dievaluasi sebagai Fe2O3 sebanyak 70 % atau lebih menurut beratnya')," & _
                                       "('0','282200','Kobalt oksida dan kobalt hidroksida kobalt oksida komersial')," & _
                                       "('0','282300','Titanium oksida')," & _
                                       "('0','282400','Timbal oksida timbal merah dan timbal oranye')," & _
                                       "('0','282500','Hidrazin dan hidroksilamin serta garam anorganiknya basa anorganik lainnya logam oksida lainny logam hidroksida dan logam peroksida lainnya')," & _
                                       "('0','282600','Fluorida fluorosilikat, fluoroaluminat dan garam fluorin kompleks lainnya')," & _
                                       "('0','282700','Klorid klorida oksida dan klorida hidroksida bromida dan bromida oksida iodida dan iodida oksida')," & _
                                       "('0','282800','Hipoklorit kalsium hipoklorit komersial klorit hipobromit')," & _
                                       "('0','282900','Klorat dan perklorat bromat dan perbromat iodat dan periodat')," & _
                                       "('0','283000','Sulfida polisulfid mempunyai rumus kimia tertentu maupun tidak')," & _
                                       "('0','283100','Ditionit dan sulfoksilat')," & _
                                       "('0','283200','Sulfit tiosulfat')," & _
                                       "('0','283300','Sulfat alum peroksosulfat (persulfat)')," & _
                                       "('0','283400','Nitrit nitrat')," & _
                                       "('0','283500','Fosfinat (hipofosfit), fosfonat (fosfit) dan fosfat polifosfat, mempunyai rumus kimia tertentu maupun tidak')," & _
                                       "('0','283600','Karbonat peroksokarbonat (perkarbonat) amonium karbonat komersial mengandung amonium karbamat')," & _
                                       "('0','283700','Sianid sianida oksida dan sianida kompleks')," & _
                                       "('0','283900','Silikat logam alkali silikat komersial.')," & _
                                       "('0','284000','Borat peroksoborat (perborat)')," & _
                                       "('0','284100','Garam dari asam oksometalik atau asam peroksometalik')," & _
                                       "('0','284200','Garam lainnya dari asam anorganik atau asam perokso (termasuk alumino silikat yang mempunyai rumus kimia tertentu maupun tidak), selain azida')," & _
                                       "('0','284300','Logam mulia koloidal senyawa organik atau anorganik dari logam muli mempunyai rumus kimia tertentu maupun tidak amalgam dari logam mulia')," & _
                                       "('0','284400','Unsur kimia radioaktif dan isotop radioaktif (termasuk unsur kimia fisil atau fertil dan isotop fisll atau fertil) dan senyawanya campuran dan residu mengandung produk ini')," & _
                                       "('0','284500','Isotop selain pos 28.44 senyaw anorganik atau organik, dari isotop seperti itu, mempunyai rumus kimia tertentu maupun tidak')," & _
                                       "('0','284600','Senyaw anorganik atau organik, dari logam tanah jarang, dari itrium atau dari skandium atau dari campuran logam ini')," & _
                                       "('0','284700','Hidrogen peroksid dipadatkan dengan urea maupun tidak')," & _
                                       "('0','284900','Karbid mempunyai rumus kimia tertentu maupun tidak')," & _
                                       "('0','285000','Hidrid nitrid azid silisida dan borid mempunyai rumus kimia tertentu maupun tidak selain senyawa yang juga merupakan karbida sebagaimana dimaksud dari pos 28.49')," & _
                                       "('0','285200','Senyawa anorganik atau organik dari merkuri, memiliki rumus kimia tertentu maupun tidak, tidak termasuk amalgam')," & _
                                       "('0','285300','Fosfid memiliki rumus kimia sendiri maupun tidak, tidak termasuk ferofosfor senyawa anorganik lainnya (termasuk air sulingan atau air konduktivitas dan air dengan kemurnian semacam itu) udara cair telah dihilangkan gas mulianya maupun tidak) udara tekan amalgam, selain amalgam dari logam mulia')," & _
                                       "('0','290000','Bahan Kimia Organik')," & _
                                       "('0','290100','Hidrokarbon asiklik')," & _
                                       "('0','290200','Hidrokarbon siklik')," & _
                                       "('0','290300','Turunan halogenasi dari hidrokarbon')," & _
                                       "('0','290400','Turunan sulfonasi, nitrasi atau nitrosasi dari hidrokarbon, dihalogenasi maupun tidak.')," & _
                                       "('0','290500','Alkohol asiklik dan turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','290600','Alkohol siklik dan turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','290700','Fenol fenol-alkohol.')," & _
                                       "('0','290800','Turunan halogenasi, sulfonasi, nitrasi atau nitrosasi dari fenol atau fenol-alkohol.')," & _
                                       "('0','290900','Eter, eter-alkohol, eter-fenol, eter-alkohol-fenol, alkohol peroksid eter peroksid asetal dan hemiasetal peroksid keton peroksida (mempunyai rumus kimia tertentu maupun tidak), dan turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','291000','Epoksid epoksi alkohol, epoksi fenol dan epoksi eter, dengan tiga cincin, dan turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','291100','Asetal dan hemiasetal, dengan fungsi oksigen lainnya maupun tidak, dan turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','291200','Aldehid dengan fungsi oksigen lainnya maupun tidak polimer siklik dari aldehida paraformaldehida.')," & _
                                       "('0','291300','Turunan halogenasi, sulfonasi, nitrasi atau nitrosasi produk dari pos 29.12.')," & _
                                       "('0','291400','Keton dan kinon, dengan fungsi oksigen lainnya maupun tidak, dan turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','291500','Asam monokarboksilat asiklik jenuh dan anhidrid halid peroksida dan asam peroksinya turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','291600','Asam monokarboksilat asiklik tidak jenuh, asam monokarboksilat siklik, anhidrid halid peroksida dan asam peroksinya turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','291700','Asam polikarboksilat, anhidrid halid peroksida dan asam peroksinya turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','291800','Asam karboksilat dengan fungsi oksigen tambahan dan anhidrid halid peroksida dan asam peroksinya turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','291900','Ester fosfat dan garamny termasuk laktofosfat turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','292000','Ester dari asam anorganik bukan logam lainnya (tidak termasuk ester dari hidrogen halida) dan garamnya turunan halogenasi, sulfonasi, nitrasi atau nitrosasinya.')," & _
                                       "('0','292100','Senyawa berfungsi amina.')," & _
                                       "('0','292200','Senyawa amino berfungsi oksigen.')," & _
                                       "('0','292300','Garam dan hidroksida amonium kuarterner lesitin dan fosfoaminolipid lainny mempunyai rumus kimia tertentu maupun tidak.')," & _
                                       "('0','292400','Senyawa berfungsi karboksiamida senyawa berfungsi amida dari asam karbonat.')," & _
                                       "('0','292500','Senyawa berfungsi karboksiimida (termasuk sakarin dan garamnya) dan senyawa berfungsi imina.')," & _
                                       "('0','292600','Senyawa berfungsi nitril.')," & _
                                       "('0','292700','Senyawa diazo-, azo- atau azoksi.')," & _
                                       "('0','292800','Turunan organik dari hidrazin atau dari hidroksilamina.')," & _
                                       "('0','292900','Senyawa dengan fungsi nitrogen lainnya.')," & _
                                       "('0','293000','Senyawa organo-belerang.')," & _
                                       "('0','293100','Senyawa organo-anorganik lainnya.')," & _
                                       "('0','293200','Senyawa heterosiklik hanya dengan hetero atom oksigen.')," & _
                                       "('0','293300','Senyawa heterosiklik hanya dengan hetero atom nitrogen.')," & _
                                       "('0','293400','Asam nukleat dan garamny mempunyai rumus kimia tertentu maupun tidak senyawa heterosiklik lainnya.')," & _
                                       "('0','293500','Sulfonamida.')," & _
                                       "('0','293600','Provitamin dan vitamin, alami atau direproduksi secara sintesis (termasuk konsentrat alami), turunannya terutama digunakan sebagai vitamin, dan campuran dari bahan yang telah disebutkan, didalam pelarut apapun maupun tidak.')," & _
                                       "('0','293700','Hormon, prostaglandin, tromboksan dan leukotrien, alami atau direproduksi secara sintesis turunan dan analog strukturny termasuk rantai polipeptida yang dimodifikasi, digunakan terutama sebagai hormon.')," & _
                                       "('0','293800','Glikosid alami atau direproduksi secara sintesis, dan garam, eter, ester serta turunan lainnya.')," & _
                                       "('0','293900','Alkaloid, alami atau direproduksi secara sintesis, dan garam, eter, ester serta turunan lainnya.')," & _
                                       "('0','294000','Gul murni secara kimi selain sukros laktos maltos glukosa dan fruktosa gula eter, gula asetal dan gula ester, serta garamny selain produk dari pos 29.37, 29.38 atau 29.39.')," & _
                                       "('0','294100','Antibiotik.')," & _
                                       "('0','294200','Senyawa organik lainnya.')," & _
                                       "('0','300000','Produk farmasi')," & _
                                       "('0','300100','Kelenjar dan bagian tubuh lainnya untuk keperluan terapeutik-organ tubuh, kering, berupa bubuk maupun tidak ekstrak dari kelenjar atau organ tubuh lainnya atau dari sekresinya untuk keperluan terapeutik organ tubuh heparin dan garamnya subtansi lainnya dari manusia atau hewan disiapkan untuk keperluan terapeutik atau profilaktik, tidak dirinci atau termasuk dalam pos lainnya.')," & _
                                       "('0','300200','Darah manusia darah hewan disiapkan untuk keperluan terapeutik, profilaktik atau diagnosis antiserum, bagian darah dan produk imunologi lainny dimodifikasi maupun tidak atau diperoleh dengan proses bioteknologi maupun tidak vaksin, toksin, kultur dari mikro-organisme (tidak termasuk ragi) dan produk yang semacam itu kultur sel, dimodifikasi maupun tidak.')," & _
                                       "('0','300300','Obat (tidak termasuk barang dari pos 30.02, 30.05 atau 30.06) terdiri dari dua atau lebih konstituen yang telah dicampur bersama-sama untuk keperluan terapeutik atau profilaktik, tidak disiapkan dalam dosis tertentu atau tidak dalam bentuk kemasan untuk penjualan eceran.')," & _
                                       "('0','300400','Obat (tidak termasuk barang dari pos 30.02, 30.05 atau 30.06) terdiri dari produk campuran atau tidak untuk keperluan terapeutik atau profilaktik, disiapkan dalam dosis tertentu (termasuk dalam bentuk sistem pemberian transdermal) atau dalam bentuk atau kemasan untuk penjualan eceran.')," & _
                                       "('0','300500','Gumpalan kapas, kas pembalut dan barang semacam itu (misalny perban, plester berperekat, tapal obat), diresapi atau dilapisi dengan zat farmasi atau disiapkan dalam bentuk atau kemasan untuk penjualan eceran untuk keperluan medis, pembedahan, perawatan gigi atau veteriner.')," & _
                                       "('0','300600','Barang farmasi dirinci dalam Catatan 4 pada Bab ini.')," & _
                                       "('0','310000','Pupuk')," & _
                                       "('0','310100','Pupuk hewani atau nabati, dicampur bersama atau diolah secara kimia maupun tidak pupuk yang dihasilkan dari pencampuran atau pengolahan secara kimia produk hewani atau nabati.')," & _
                                       "('0','310200','Pupuk mineral atau kimi mengandung nitrogen.')," & _
                                       "('0','310300','Pupuk mineral atau pupuk kimi mengandung fosfat.')," & _
                                       "('0','310400','Pupuk mineral atau kimi mengandung kalium.')," & _
                                       "('0','310500','Pupuk mineral atau kimia mengandung dua atau tiga unsur penyubur nitrogen, fosfor dan kalium pupuk lainnya barang dari Bab ini dalam bentuk tablet atau bentuk semacam itu atau dalam kemasan dengan berat kotor tidak melebihi 10 kg.')," & _
                                       "('0','320000','Ekstrak penyamak atau pencelup tanin dan turunannya bahan celup, pigmen dan bahan pewarna lainnya cat dan pernis dempul dan mastik lainnya tinta')," & _
                                       "('0','320100','Ekstrak penyamak dari bahan nabati tanin dan garamny eter, ester serta turunan lainnya.')," & _
                                       "('0','320200','Zat penyamak organik sintetik zat penyamak anorganik preparat penyamak, mengandung zat penyamak alami maupun tidak preparat mengandung enzim untuk penyamakan pendahuluan.')," & _
                                       "('0','320300','Bahan pewarna dari nabati atau hewani (termasuk ekstrak pencelupan tetapi tidak termasuk animal black), mempunyai rumus kimia tertentu maupun tidak preparat sebagaimana dirinci dalam Catatan 3 pada Bab ini yang berasal dari bahan pewarna nabati atau hewani.')," & _
                                       "('0','320400','Bahan pewarna organik sintetik, mempunyai rumus kimia tertentu maupun tidak preparat sebagaimana dirinci dalam Catatan 3 dari Bab ini yang berasal dari bahan pewarna organik sintetik produk organik sintetik dari jenis yang digunakan sebagai bahan pencemerlang fluoresen atau sebagai luminofor, mempunyai rumus kimia tertentu maupun tidak.')," & _
                                       "('0','320500','Lak warna preparat yang dibuat dari lak warna sebagaimana dirinci dalam Catatan 3 pada Bab ini.')," & _
                                       "('0','320600','Bahan pewarna lainnya preparat sebagaimana dirinci dalam Catatan 3 pada Bab ini selain yang disebut dalam pos 32.03, 32.04 atau 32.05 produk anorganik dari jenis yang digunakan sebagai luminofor, mempunyai rumus kimia tertentu maupun tidak.')," & _
                                       "('0','320700','Pigmen olahan, bahan opasitas olahan dan pewarna olahan, enamel dan glasir yang dapat divitrifikasi, engobes (slip), pengkilap cair dan preparat semacam itu, dari jenis yang digunakan dalam industri keramik, enamel atau kaca frit kaca dan kaca lainny dalam bentuk bubuk, butiran atau serpih.')," & _
                                       "('0','320800','Cat dan pernis (termasuk enamel dan lak) dibuat dari polimer sintetik atau polimer alam yang dimodifikasi secara kimi didispersikan atau dilarutkan dalam media bukan air larutan sebagaimana dinyatakan dalam Catatan 4 pada Bab ini.')," & _
                                       "('0','320900','Cat dan pernis (termasuk enamel dan lak) dibuat dari polimer sintetik atau polimer alam yang dimodifikasi secara kimi didispersikan atau dilarutkan dalam media air.')," & _
                                       "('0','321000','Cat dan pernis lainnya (termasuk enamel, lak dan distemper) pigmen air olahan dari jenis yang digunakan untuk penyempurnaan kulit.')," & _
                                       "('0','321100','Olahan pengering.')," & _
                                       "('0','321200','Pigmen (termasuk bubuk dan serpih metalik) didispersikan dalam media bukan air, dalam bentuk cair atau past dari jenis yang digunakan dalam pembuatan cat (termasuk enamel) stamping foil bahan celup dan pewarna lainnya disiapkan dalam bentuk atau kemasan untuk penjualan eceran.')," & _
                                       "('0','321300','Pewarna untuk seniman, pelajar atau pelukis papan petunjuk, modifying tints, untuk hiburan dan sejenisny dalam tablet, tabung, jar, botol, bejana atau dalam bentuk atau kemasan semacam itu.')," & _
                                       "('0','321400','Dempul kac dempul pengenten, semen resin, kompon untuk mendempul dan mastik lainnya painters'' filling preparat pelapis tidak tahan panas untuk bagian luar, dinding bagian dalam, lantai, langit-langit atau sejenisnya.')," & _
                                       "('0','321500','Tinta cetak, tinta tulis atau tinta gambar dan tinta lainny konsentrat atau padat maupun tidak.')," & _
                                       "('0','330000','Minyak atsiri dan resinoida preparat wewangian, kosmetika atau rias')," & _
                                       "('0','330100','Minyak atsiri (mengandung terpena atau tidak), termasuk konkrit dan absolut resinoida ekstrak oleoresin konsentrat minyak atsiri dalam lemak, dalam fixed oil, dalam malam atau sejenisny diperoleh melalui enfleurage atau maserasi produk sertaan bersifat terpena pada proses penghilangan terpena dari minyak atsiri hasil sulingan dengan air dan larutan air dari minyak atsiri.')," & _
                                       "('0','330200','Campuran dari berbagai zat bau-bauan dan campuran (termasuk larutan mengandung alkohol) dengan bahan dasar satu atau lebih dari berbagai zat ini, dari jenis yang digunakan sebagai bahan baku industri preparat lain yang dibuat dari berbagai zat bau-bauan, dari jenis yang digunakan untuk pembuatan minuman.')," & _
                                       "('0','330300','Parfum dan cairan pewangi.')," & _
                                       "('0','330400','Preparat kecantikan atau rias dan preparat untuk perawatan kulit (selain obat-obatan), termasuk preparat pelindung kulit terhadap sinar matahari atau pencoklat kulit preparat manikur atau pedikur.')," & _
                                       "('0','330500','Preparat digunakan untuk rambut.')," & _
                                       "('0','330600','Preparat kesehatan mulut atau gigi, termasuk pasta dan bubuk penguat gigi buatan benang untuk pembersih sela gigi (dental floss), dalam kemasan tersendiri untuk penjualan eceran.')," & _
                                       "('0','330700','Preparat yang digunakan sebelum mencukur, sewaktu mencukur, atau sesudah mencukur, deodoran, preparat mandi, preparat perontok bulu dan preparat wewangian, kosmetika atau rias lainny tidak dirinci atau termasuk pos lain preparat penghilang bau ruangan, diberi wewangian atau mengandung desinfektan maupun tidak.')," & _
                                       "('0','340000','Sabun, bahan aktif permukaan organik, preparat pembersih, preparat pelumas, malam artifisial, malam olahan, preparat pemoles dan penggosok, lilin dan barang semacam itu, pasta untuk model, malam untuk gigi dan preparat untuk gigi dengan bahan dasar plester')," & _
                                       "('0','340100','Sabun produk dan preparat aktif-permukaan organik digunakan sebagai sabun, dalam bentuk batangan, cake, potongan atau bentukan yang dicetak, mengandung sabun maupun tidak produk dan preparat aktif permukaan organik untuk membersihkan kulit dalam bentuk cair atau krim dan disiapkan untuk penjualan eceran, mengandung sabun maupun tidak kertas, gumpalan, kain kempa dan bukan tenunan, diresapi, dilapisi atau ditutupi dengan sabun atau deterjen.')," & _
                                       "('0','340200','Bahan aktif permukaan organik (selain sabun) preparat aktif permukaan, preparat pencuci (termasuk preparat pencuci tambahan) dan preparat pembersih, mengandung sabun maupun tidak, selain dari pos 34.01.')," & _
                                       "('0','340300','Preparat pelumas (termasuk preparat minyak pemotong, preparat pelepas baut atau mur, preparat anti-karat atau anti-korosi dan preparat pelepas cetakan, dengan bahan dasar pelumas) dan preparat dari jenis yang digunakan untuk meminyaki atau menggemuki bahan tekstil, kulit samak, kulit berbulu atau material lainny tetapi tidak termasuk preparat yang mengandung minyak petroleum atau minyak yang diperoleh dari mineral mengandung bitumen sebagai unsur dasar, 70 % atau lebih menurut beratnya.')," & _
                                       "('0','340400','Malam artifisial dan malam olahan.')," & _
                                       "('0','340500','Poles dan krim, untuk alas kaki, perabotan, lantai, coachwork, kaca atau logam, pasta dan bubuk penggosok dan preparat semacam itu (dalam bentuk kertas, gumpalan, kain kemp bukan tenunan, plastik seluler atau karet seluler, diresapi, dilapisi atau ditutupi dengan preparat tersebut maupun tidak), tidak termasuk malam dari pos 34.04.')," & _
                                       "('0','340600','Lilin, taper dan sejenisnya.')," & _
                                       "('0','340700','Pasta untuk model, termasuk yang disiapkan untuk hiburan anak-anak preparat yang dikenal sebagai malam untuk gigi atau sebagai kompon untuk membuat cetakan gigi')," & _
                                       "('0','350000','Zat albumina pati dimodifikasi perekat enzim')," & _
                                       "('0','350100','Kasein, kaseinat dan turunan kasein lainnya lem kasein.')," & _
                                       "('0','350200','Albumin (termasuk konsentrat dari dua atau lebih protein whey, yang mengandung protein whey lebih dari 80 % menurut beratny dihitung dari bahan kering), albuminat dan turunan albumin lainnya.')," & _
                                       "('0','350300','Gelatin (termasuk lembaran gelatin dalam bentuk empat persegi panjang (termasuk bujur sangkar), dikerjakan permukaannya atau diwarnai maupun tidak) dan turunan gelatin isinglass lem lain dari bahan hewani, tidak termasuk lem kasein dari pos 35.01.')," & _
                                       "('0','350400','Pepton dan turunannya zat protein lain dan turunanny tidak dirinci atau termasuk dalam pos lainnya bubuk kulit jangat, dikrom maupun tidak.')," & _
                                       "('0','350500','Dekstrin dan pati yang dimodifikasi lainnya (misalny pati yang dipregelatinisasi atau yang diesterifikasi) lem yang dibuat dari pati, atau dekstrin atau pati yang dimodifikasi lainnya.')," & _
                                       "('0','350600','Lem olahan dan perekat olahan lainny tidak dirinci atau termasuk dalam pos lain produk yang cocok digunakan sebagai lem atau perekat, disiapkan untuk penjualan eceran sebagai lem atau perekat, dengan berat bersih tidak melebihi 1 kg.')," & _
                                       "('0','350700','Enzim enzim olahan tidak dirinci atau termasuk dalam pos lainnya.')," & _
                                       "('0','360000','Bahan peledak produk piroteknik korek api paduan piroforik preparat tertentu yang mudah terbakar')," & _
                                       "('0','360100','Bubuk propelan.')," & _
                                       "('0','360200','Bahan peledak olahan, selain bubuk propelan.')," & _
                                       "('0','360300','Sekering pengaman kabel peledak sumbu mesiu atau cap peledak penyala detonator listrik.')," & _
                                       "('0','360400','Kembang api, suar pemberi sinyal, roket hujan, sinyal kabut dan barang piroteknik lainnya.')," & _
                                       "('0','360500','Korek api, selain barang piroteknik dari pos 36.04.')," & _
                                       "('0','360600','Ferro-cerium dan paduan piroforik lainnya dalam segala bentuk barang dari bahan mudah terbakar sebagaimana dirinci dalam Catatan 2 Bab ini.')," & _
                                       "('0','370000','Barang fotografi atau sinematografi')," & _
                                       "('0','370100','Pelat dan film fotografi berbentuk lembaran, peka cahay tidak disinari, dari bahan apapun selain kertas, kertas karton atau tekstil film cetak instan berbentuk lembaran, peka cahay tidak disinari, dalam kemasan maupun tidak.')," & _
                                       "('0','370200','Film fotografi dalam gulungan, peka cahay tidak disinari, dari bahan apapun selain kertas, kertas karton atau tekstil film cetak instan dalam gulungan, peka cahay tidak disinari.')," & _
                                       "('0','370300','Kertas, kertas karton dan tekstil fotografi, peka cahay tidak disinari.')," & _
                                       "('0','370400','Pelat, film, kertas, kertas karton dan tekstil fotografi, disinari tetapi tidak dicuci.')," & _
                                       "('0','370500','Pelat dan film fotografi, disinari dan dicuci, selain film sinematografi.')," & _
                                       "('0','370600','Film sinematografi, disinari dan dicuci, digabung dengan trek suara maupun tidak atau hanya terdiri dari trek suara.')," & _
                                       "('0','370700','Preparat kimia digunakan untuk fotografi (selain pernis, lem, perekat dan preparat semacam itu) produk tidak dicampur digunakan untuk fotografi, disiapkan dalam ukuran tertentu atau disiapkan untuk penjualan eceran dalam bentuk siap pakai.')," & _
                                       "('0','380000','Aneka produk kimia')," & _
                                       "('0','380100','Grafit artifisial grafit koloidal atau grafit semi koloidal preparat dibuat dari grafit atau karbon lainnya dalam bentuk past blok, pelat atau barang semi manufaktur lainnya.')," & _
                                       "('0','380200','Karbon aktif produk mineral alam aktif jelaga hewani, termasuk jelaga hewani habis pakai.')," & _
                                       "('0','380300','Minyak tall, dimurnikan maupun tidak.')," & _
                                       "('0','380400','Residu alkali dari pembuatan pulp kayu, dipekatkan maupun tidak, dihilangkan gulanya atau diolah secara kimi termasuk lignin sulfonat, tetapi tidak termasuk minyak tall dari pos 38.03.')," & _
                                       "('0','380500','Getah, terpentin kayu atau sulfat dan minyak mengandung terpena lainnya yang dihasilkan dari penyulingan atau pengolahan lainnya dari kayu jenis konifera dipentena mentah terpentin sulfit dan para-simen mentah lainnya minyak cemara mengandung alfa-terpineol sebagai bahan utama.')," & _
                                       "('0','380600','Asam rosin dan asam damar serta turunannya spirit rosin dan minyak rosin run gum.')," & _
                                       "('0','380700','Ter kayu minyak ter kayu kreosot kayu nafta kayu pek nabati pek untuk pabrik bir dan preparat semacam itu dibuat dari rosin, asam damar atau pek nabati.')," & _
                                       "('0','380800','Insektisid rodentisid fungisid herbisid produk anti-sprouting dan pengatur pertumbuhan tanaman, desinfektan dan produk semacam itu, disiapkan dalam bentuk atau kemasan untuk penjualan eceran atau sebagai preparat atau barang (misalny pit sumbu dan lilin yang diproses dengan belerang, dan kertas lalat).')," & _
                                       "('0','380900','Bahan untuk penyempurnaan, pembawa bahan celup untuk mempercepat proses pencelupan atau pengolahan bahan celup dan produk serta preparat lainnya (misalnya dressing dan mordant), dari jenis yang digunakan dalam industri tekstil, kertas, kulit samak atau industri sejenis, tidak dirinci atau termasuk dalam pos lainnya.')," & _
                                       "('0','381000','Preparat bersifat asam untuk permukaan logam flux dan preparat tambahan lainnya untuk menyolder, mematri atau mengelas bubuk dan pasta untuk menyolder, mematri atau mengelas terdiri dari logam dan bahan lain preparat dari jenis yang digunakan sebagai inti atau pelapis untuk elektroda las atau batang las.')," & _
                                       "('0','381100','Preparat untuk mencegah ketukan, pencegah oksidasi, mencegah gum, meningkatkan viskositas, preparat anti-korosi dan olahan aditif lainny untuk minyak mineral (termasuk bensin) atau untuk cairan lain yang digunakan untuk tujuan yang sama seperti minyak mineral.')," & _
                                       "('0','381200','Olahan akselerator untuk karet kompon peliat untuk karet atau plastik, tidak dirinci atau termasuk dalam pos lainnya preparat antioksidasi dan kompon stabilisator lainnya untuk karet atau plastik.')," & _
                                       "('0','381300','Preparat dan bahan pengisi alat pemadam kebakaran granat pemadam kebakaran yang sudah diisi.')," & _
                                       "('0','381400','Komposit pelarut dan pengencer organik, tidak dirinci atau termasuk dalam pos lainnya olahan penghapus cat atau pernis.')," & _
                                       "('0','381500','Preparat inisiator reaksi, akselerator reaksi dan katalisator, tidak dirinci atau termasuk dalam pos lain.')," & _
                                       "('0','381600','Semen, mortar, beton dan komposisi semacam itu yang tahan api, termasuk campuran adukan dolomit, selain produk dari pos 38.01.')," & _
                                       "('0','381700','Campuran alkilbenzena dan campuran alkilnaftalen selain yang dimaksud dalam pos 27.07 atau 29.02.')," & _
                                       "('0','381800','Unsur kimia yang diolah untuk digunakan dalam elektronik, dalam bentuk cakram, wafer atau bentuk semacam itu senyawa kimia yang diolah untuk digunakan dalam elektronik.')," & _
                                       "('0','381900','Cairan rem hidrolik dan cairan olahan lainnya untuk transmisi hidrolik, tidak mengandung atau mengandung minyak petroleum atau minyak yang diperoleh dari mineral mengandung bitumen kurang dari 70 % menurut beratnya.')," & _
                                       "('0','382000','Preparat anti beku dan cairan olahan untuk mencairkan es.')," & _
                                       "('0','382100','Media kultur olahan untuk pengembangan atau perawatan mikro organisme (termasuk virus dan semacamnya) atau sel tumbuhan, manusia atau hewan.')," & _
                                       "('0','382200','Reagen diagnosa atau laboratorium pada bahan pendukung, olahan reagen diagnosa atau laboratorium pada bahan pendukung maupun tidak, disiapkan dalam bentuk kit maupun tidak, selain yang dimaksud dalam pos 30.06 bahan referensi bersertifikat.')," & _
                                       "('0','382300','Asam lemak monokarboksilat industri minyak asam dari pemurnian alkohol lemak industri.')," & _
                                       "('0','382400','Olahan pengikat untuk acuan atau inti penuangan logam produk dan preparat kimia dari industri kimia atau industri terkait (termasuk olahan yang terdiri dari campuran produk alami), tidak dirinci atau termasuk dalam pos lainnya.')," & _
                                       "('0','382500','Produk residu dari industri kimia atau industri terkait, tidak dirinci atau termasuk dalam pos lainnya limbah rumah tangga endapan kotoran limbah lainnya yang dirinci dalam Catatan 6 pada bab ini.')," & _
                                       "('0','382600','Biodiesel dan campuranny tidak mengandung atau mengandung kurang dari 70 % menurut beratnya minyak petroleum atau minyak yang diperoleh dari mineral mengandung bitumen.')," & _
                                       "('0','382700','Campuran mengandung turunan halogenasi dari metan etana atau propan tidak dirinci atau termasuk dalam pos lainnya.')," & _
                                       "('0','390000','Plastik dan barang daripadanya')," & _
                                       "('0','390100','Polimer dari etilen dalam bentuk asal.')," & _
                                       "('0','390200','Polimer dari propilena atau dari olefin lainny dalam bentuk asal.')," & _
                                       "('0','390300','Polimer dari stiren dalam bentuk asal.')," & _
                                       "('0','390400','Polimer dari vinil klorida atau dari halogenasi olefin lainny dalam bentuk asal.')," & _
                                       "('0','390500','Polimer dari vinil asetat atau dari vinil ester lainny dalam bentuk asal polimer vinil lainnya dalam bentuk asal.')," & _
                                       "('0','390600','Polimer akrilik dalam bentuk asal.')," & _
                                       "('0','390700','Poliasetal, polieter lainnya dan resin epoksid dalam bentuk asal polikarbonat, resinalkid, polialil ester dan poliester lainny dalam bentuk asal.')," & _
                                       "('0','390800','Poliamida dalam bentuk asal.')," & _
                                       "('0','390900','Amino-resin, resin fenolik dan poliuretan, dalam bentuk asal.')," & _
                                       "('0','391000','Silikon dalam bentuk asal.')," & _
                                       "('0','391100','Resin petroleum, resin kumaron-inden politerpen polisulfid polisulfon dan produk lain dirinci dalam Catatan 3 pada Bab ini, tidak dirinci atau termasuk dalam pos lainny dalam bentuk asal.')," & _
                                       "('0','391200','Selulosa dan turunan kimiany tidak dirinci atau termasuk dalam pos lainny dalam bentuk asal.')," & _
                                       "('0','391300','Polimer alam (misalny asam alginat) dan polimer alam yang dimodifikasi (misalnya protein dikeraskan, turunan kimia dari karet alam), tidak dirinci atau termasuk dalam pos manapun, dalam bentuk asal.')," & _
                                       "('0','391400','Penukar ion dibuat dari polimer pos 39.01 sampai dengan 39.13, dalam bentuk asal.');"

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetUrutan3() As String
        Dim retval As String = ""
        Try
            retval = "Insert Into `tblmaster_taxproduct_reference` (`type`,`Code`,`Description`) Values" & _
"('0','800300','Batang, batang kecil, profil dan kawat timah.')," & _
"('0','800700','Barang lainnya dari timah.')," & _
"('0','810000','Logam tidak mulia lainnya sermet barang')," & _
"('0','810100','Tungsten (wolfram) dan barang dari padany termasuk sisa dan skrap.')," & _
"('0','810200','Molibdenum dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','810300','Tantalum dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','810400','Magnesium dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','810500','Mate kobalt dan produk antara lainnya dari metalurgi kobalt kobalt dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','810600','Bismut dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','810800','Titanium dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','810900','Zirkonium dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','811000','Antimoni dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','811100','Mangan dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','811200','Berilium, kromium, germanium, vanadium, galium, hafnium, indium, niobium (columbium), renium dan talium serta barang dari logam tersebut, termasuk sisa dan skrap.')," & _
"('0','811300','Sermet dan barang daripadany termasuk sisa dan skrap.')," & _
"('0','820000','Perkakas, peralatan, barang tajam, sendok dan garpu, dari logam tidak mulia bagiannya dari logam tidak mulia')," & _
"('0','820100','Perkakas tangan, sebagai berikut : sekop datar, sekop lengkung, beliung, bajak, cangkul, garu dan pengeruk kapak, sabit paruh dan alat potong semacam itu gunting bunga dan pemangkas dari segala jenis sabit tangkai panjang, sabit, pisau jerami, gunting untuk tanaman pagar, pengiris kayu dan perkakas lainnya dari jenis yang digunakan dalam pertanian, perkebunan atau kehutanan.')," & _
"('0','820200','Gergaji tangan bilah untuk segala jenis gergaji (termasuk bilah gergaji untuk membelah, untuk membuat celah atau tanpa gigi).')," & _
"('0','820300','Kikir, parut, tang (termasuk tang potong), penjepit, gunting seng, gunting pemotong logam, pemotong pip catok, tang pembuat lobang dan perkakas tangan semacam itu.')," & _
"('0','820400','Kunci sekrup dan kunci inggris yang digerakkan dengan tangan (termasuk kunci inggris yang mempunyai pengukur putaran tetapi tidak termasuk kunci pembuat alur sekrup) kunci sekrup berlobang yang dapat dipertukarkan, dengan atau tanpa gagang.')," & _
"('0','820500','Perkakas tangan (termasuk intan pemotong kaca), tidak dirinci atau termasuk dalam pos lainnya lampu tiup vice, klem dan sejenisny selain aksesori untuk dan bagian dari, perkakas mesin atau mesin pemotong water-jet paron alat tempa portabel roda gerinda yang digerakkan dengan tangan atau pedal dengan rangkanya.')," & _
"('0','820600','Perkakas terdiri dari dua jenis barang atau lebih dimaksud pada pos 82.02 sampai dengan 82.05, disiapkan dalam set untuk penjualan eceran.')," & _
"('0','820700','Perkakas yang dapat dipertukarkan untuk perkakas tangan, digerakkan dengan tenaga maupun tidak, atau untuk mesin perkakas (misalny untuk pressing, stamping, punching, tapping, threading, drilling, boring, broaching, milling, turning atau screw driving), termasuk acuan untuk menarik atau mengekstrusi logam, dan perkakas untuk mengebor batu karang atau mengebor tanah.')," & _
"('0','820800','Pisau dan bilah pemotong, untuk mesin atau untuk perkakas mekanik.')," & _
"('0','820900','Pelat, tongkat, ujung dan sejenisnya untuk perkakas, tidak terpasang, dari sermet.')," & _
"('0','821000','Perkakas mekanik yang digerakkan dengan tangan, dengan berat 10 kg atau kurang, digunakan dalam pengolahan, pengkondisian atau penyiapan makanan atau minuman.')," & _
"('0','821100','Pisau dengan bilah pemotong, bergigi atau tidak (termasuk pisau pangkas), selain pisau dari pos 82.08, dan mata pisau untuk barang tersebut.')," & _
"('0','821200','Pisau cukur dan silet (termasuk blank silet dalam bentuk strip).')," & _
"('0','821300','Gunting, gunting tukang jahit dan gunting semacam itu, serta bilahnya.')," & _
"('0','821400','Barang tajam lainnya (misalny alat pangkas rambut, parang tukang daging atau parang dapur, pisau pemotong dan pisau cincang, pisau kertas) set dan instrumen untuk manikur atau pedikur (termasuk kikir kuku).')," & _
"('0','821500','Sendok, garpu, sendok sayur, peniris, cake-server, pisau ikan, pisau menteg penjepit gula dan perangkat dapur atau meja semacam itu.')," & _
"('0','830000','Bermacam-macam barang dari logam tidak mulia')," & _
"('0','830100','Gembok dan kunci (dioperasikan dengan anak kunci, kombinasi atau secara listrik), dari logam tidak mulia kunci jepit dan bingkai dengan kunci jepit, kunci terpasang, dari logam tidak mulia anak kunci untuk semua barang tersebut, dari logam tidak mulia.')," & _
"('0','830200','Penyangg alat kelengkapan dan barang semacam itu dari logam tidak mulia yang cocok untuk perabotan, pintu, tangg jendel tirai, coachwork, saddlery, koper, kotak, kotak perhiasan atau sejenisnya rak topi, gantungan topi, kaitan dan barang semacam itu dari logam tidak mulia castor dengan penyangga dari logam tidak mulia penutup pintu otomatis dari logam tidak mulia.')," & _
"('0','830300','Lemari lapis baja atau yang diperkuat, peti pengaman dan pintu serta laci penyimpan untuk ruangan pengaman, peti dokumen atau uang dan sejenisny dari logam tidak mulia.')," & _
"('0','830400','Lemari arsip, lemari kartu indeks, nampan kertas, sandaran kertas, nampan pen penyangga stempel kantor dan perlengkapan kantor atau meja tulis semacam itu, dari logam tidak muli selain perabotan kantor dari pos 94.03.')," & _
"('0','830500','Alat kelengkapan untuk loose-leaf binder atau arsip, jepitan surat, letter corner, jepitan kertas, pemarka indeks dan barang kantor semacam itu dari logam tidak mulia kokot dalam bentuk strip (misalny untuk kantor, upholstery, pengepakan), dari logam tidak mulia.')," & _
"('0','830600','Lonceng, gong dan sejenisny tanpa listrik, dari logam tidak mulia patung dan ornamen lainny dari logam tidak mulia bingkai foto, gambar dan semacam itu, dari logam tidak mulia cermin dari logam tidak mulia.')," & _
"('0','830700','Pembuluh fleksibel dari logam tidak muli dengan atau tanpa alat kelengkapan.')," & _
"('0','830800','Jepit, bingkai dengan jepit, gesper, jepitan gesper, kait, lubang, lubang tali dan sejenisnya dari logam tidak muli dari jenis yang digunakan untuk pakaian atau aksesoris pakaian, alas kaki, perhiasan, jam tangan, buku, tirai, barang dari kulit, barang untuk bepergian atau saddlery atau barang jadi lainnya paku keling berbentuk tabung atau bercabang, dari logam tidak mulia manik-manik dan spangle, dari logam tidak mulia.')," & _
"('0','830900','Sumbat, penutup dan tutup (termasuk sumbat tajuk, penutup sekrup dan sumbat yang dituangkan), kapsul untuk botol, sumbat berulir, penyingkup sumbat, segel dan aksesori pengepakan lainny dari logam tidak mulia.')," & _
"('0','831000','Papan isyarat, papan nam papan alamat dan papan semacam itu, nomor, huruf dan lambang lainny dari logam tidak muli tidak termasuk yang dimaksud dalam pos 94.05.')," & _
"('0','831100','Kawat, batang kecil, pembuluh, pelat, elektroda dan produk semacam itu, dari logam tidak mulia atau dari karbida logam, dilapisi atau diisi dengan fluks dari jenis yang digunakan untuk menyolder, mematri, mengelas atau mengendapkan logam atau karbida logam kawat dan batang kecil, dari bubuk logam tidak mulia diaglomerasi, digunakan untuk penyemprotan logam.')," & _
"('0','840000','Reaktor nuklir, ketel, mesin dan peralatan mekanis bagian daripadanya')," & _
"('0','840100','Reaktor nuklir elemen bahan bakar (cartridges), non radiasi, untuk reaktor nuklir mesin dan aparatus untuk separasi isotopik.')," & _
"('0','840200','Ketel uap air atau ketel uap lainnya (selain ketel air panas untuk pemanasan sentral yang juga dapat menghasilkan uap air tekanan rendah) ketel air super-heated.')," & _
"('0','840300','Ketel pemanasan sentral selain yang dimaksud dari pos 84.02.')," & _
"('0','840400','Instalasi pembantu untuk digunakan dengan ketel yang dimaksud dari pos 84.02 atau 84.03 (misalny economiser, super-heater, soot remover, gas recoverer) kondensor untuk unit tenaga uap air atau unit tenaga uap lainnya.')," & _
"('0','840500','Generator pemroduksi gas atau generator gas air, dengan atau tanpa alat pemurnian generator gas asetilena dan generator gas dengan proses air semacam itu, dengan atau tanpa alat pemurnian.')," & _
"('0','840600','Turbin uap air dan turbin uap lainnya.')," & _
"('0','840700','Mesin piston pembakaran dalam cetus api bolak-balik atau berputar.')," & _
"('0','840800','Mesin piston pembakaran dalam nyala kompresi (mesin diesel atau semi diesel).')," & _
"('0','840900','Bagian yang cocok untuk digunakan semata-mata atau terutama dengan mesin dari pos 84.07 atau 84.08.')," & _
"('0','841000','Turbin hidrolik, kincir air, dan regulatornya.')," & _
"('0','841100','Turbo-jet, turbo-propeller dan turbin gas lainnya.')," & _
"('0','841200','Mesin dan motor lainnya.')," & _
"('0','841300','Pompa untuk cairan, dilengkapi dengan alat ukur maupun tidak elevator cairan.')," & _
"('0','841400','Pompa udara atau pompa vakum, kompresor udara atau kompresor gas lainnya dan kipas angin atau kipas gas lainnya hood ventilasi atau hood daur ulang yang digabung dengan kipas, dilengkapi dengan saringan maupun tidak kabinet pengaman biologis kedap gas, dilengkapi dengan saringan maupun tidak.')," & _
"('0','841500','Mesin pengatur suhu udar terdiri dari kipas yang digerakkan dengan motor dan elemen untuk mengubah suhu dan kelembaban udar termasuk mesin tersebut yang tidak dapat mengatur kelembaban udara secara terpisah.')," & _
"('0','841600','Tanur pembakar untuk bahan bakar cair, untuk bahan bakar padat yang dihancurkan atau untuk gas penyala mekanis, termasuk penggarang mekanisny pembuang abu mekanisnya dan peralatan semacam itu.')," & _
"('0','841700','Tanur dan oven industri atau laboratorium, termasuk incinerator, bukan listrik.')," & _
"('0','841800','Lemari pendingin, lemari pembeku dan perlengkapan pendingin atau pembeku lainny listrik atau lainnya pompa panas selain mesin pengatur suhu udara dari pos 84.15.')," & _
"('0','841900','Perlengkapan mesin, pabrik atau laboratorium, dipanaskan secara elektrik maupun tidak (tidak termasuk tanur, oven dan perlengkapan lain-nya dari pos 85.14), untuk mengolah bahan dengan proses yang memerlukan perubahan suhu seperti memanaskan, memasak, memanggang, menyuling, rektifikasi, mensterilkan, mempasteurisasi, menguapkan, mengeringkan, mengevaporasi, menguapkan, mengkondensasi atau mendinginkan, selain mesin atau instalasi dari jenis yang digunakan untuk keperluan rumah tangga pemanas air instant atau pemanas air dengan tempat penyimpanan, bukan listrik.')," & _
"('0','842000','Mesin calender atau mesin giling lainny selain untuk logam atau kac dan silinder untuk mesin tersebut.')," & _
"('0','842100','Mesin sentrifugal, termasuk pengering sentrifugal mesin dan aparatus penyaring atau pemurni, untuk cairan atau gas.')," & _
"('0','842200','Mesin pencuci piring mesin untuk membersihkan atau mengeringkan botol atau kemasan lainnya mesin untuk mengisi, menutup, menyegel atau memasang label pada botol, kaleng, kotak, kantong atau kemasan lainnya mesin untuk menutup dengan selaput pada botol, guci, tabung dan kemasan semacam itu mesin pengepak atau pembungkus lainnya (termasuk mesin pembungkus heatshrink) mesin untuk mengaerasi minuman.')," & _
"('0','842300','Mesin penimbang (tidak termasuk timbangan dengan kepekaan timbangan sebesar 5 cg atau lebih baik), termasuk mesin penghitung atau mesin pemeriksa yang dioperasikan dengan anak timbangan anak timbangan dari segala jenis mesin timbang.')," & _
"('0','842400','Peralatan mekanis (digerakkan dengan tangan maupun tidak) untuk melemparkan, menyebarkan atau menyemprotkan barang cairan atau bubuk pemadam api, diisi maupun tidak pistol semprot dan peralatan semacam itu mesin penyembur uap air atau pasir dan mesin jet pelempar semacam itu.')," & _
"('0','842500','Katrol dan kerekan, selain kerekan skip derek dan kapstan dongkrak.')," & _
"('0','842600','Derek kapal crane, termasuk crane kabel rangka pengangkat yang dapat berpindah, straddle carrier dan truk kerja yang dilengkapi crane.')," & _
"('0','842700','Truk forklift truk kerja lainnya yang dilengkapi dengan perlengkapan pengangkat atau penanganan.')," & _
"('0','842800','Mesin pengangkat, penanganan, pemuat atau pembongkar lainnya (misalny lift, eskalator, konveyor, teleferic).')," & _
"('0','842900','Buldoser, angledoser, grader, mesin perat mesin pengikis, shovel mekanik, ekskavator, shovel loader, mesin pemadat dan mesin gilas jalan, berdaya gerak sendiri.')," & _
"('0','843000','Mesin pengolah, grading, perat pengikis, penggali, pemadat, perapi, pengaduk atau pengebor lainny untuk tanah, mineral atau bijih pemancang tiang dan pemancang bor bajak salju dan blower salju.')," & _
"('0','843100','Bagian yang cocok untuk digunakan semata-mata atau terutama dengan mesin dari pos 84.25 sampai dengan 84.30.')," & _
"('0','843200','Mesin pertanian, perkebunan atau kehutanan untuk penyiapan atau kultivasi tanah mesin giling lapangan rumput atau lapangan olah raga.')," & _
"('0','843300','Mesin tuai atau mesin tebah, termasuk pengepak jerami atau rumput makanan ternak mesin pemotong rumput atau rumput kering mesin untuk membersihkan, menyortir atau memilih mutu telur, buah atau produk pertanian lainny selain mesin dari pos 84.37.')," & _
"('0','843400','Mesin perah susu dan mesin produk susu.')," & _
"('0','843500','Pengepres, penghancur dan mesin semacam itu yang digunakan dalam pembuatan wine, cider, jus buah atau minuman semacam itu.')," & _
"('0','843600','Mesin pertanian, perkebunan, kehutanan, mesin pemeliharaan unggas atau lebah lainny termasuk mesin tunas dilengkapi dengan perlengkapan mekanis atau termal inkubator dan penetas untuk unggas.')," & _
"('0','843700','Mesin untuk membersihkan, menyortir atau memilih mutu biji, butir atau sayuran polongan kering mesin yang digunakan dalam industri penggilingan atau untuk pembuatan serealia atau sayuran polongan kering, selain mesin tipe pertanian.')," & _
"('0','843800','Mesin, tidak dirinci atau termasuk dalam pos manapun pada Bab ini, untuk industri pengolahan atau pembuatan makanan atau minuman, selain mesin untuk ekstraksi atau pengolahan lemak atau minyak dari hewan, nabati tertentu atau mikroba.')," & _
"('0','843900','Mesin untuk membuat pulp dari bahan serat selulosa atau untuk membuat atau merampungkan kertas atau kertas karton.')," & _
"('0','844000','Mesin penjilid buku, termasuk mesin penjahit buku.')," & _
"('0','844100','Mesin lainnya untuk membuat pulp kertas, kertas atau kertas karton, termasuk mesin pemotong dari semua jenis.')," & _
"('0','844200','Mesin, aparatus dan perlengkapan (selain mesin yang dimaksud dari pos 84.56 sampai dengan 84.65) untuk menyiapkan atau membuat pelat, silinder cetak atau komponen cetak lainnya pelat, silinder cetak dan komponen cetak lainnya pelat, silinder dan batu litograf, disiapkan untuk keperluan pencetakan (misalny diratakan, dibuat tidak licin atau dipoles).')," & _
"('0','844300','Mesin cetak digunakan untuk mencetak dengan memakai pelat, silinder dan komponen cetak lainnya dari pos 84.42 mesin cetak, mesin fotocopy dan mesin faksimili lainny dikombinasi atau tidak bagian dan aksesori lainnya.')," & _
"('0','844400','Mesin untuk mengekstrusi, menarik, mentekstur atau memotong bahan tekstil buatan.')," & _
"('0','844500','Mesin untuk pengolahan serat tekstil mesin pemintal, pengganda atau pemilin dan mesin lainnya untuk memproduksi benang tekstil mesin pengikal atau penggulung tekstil (termasuk penggulung benang pakan) dan mesin untuk menyiapkan benang tekstil untuk digunakan pada mesin dari pos 84.46 atau 84.47.')," & _
"('0','844600','Mesin tenun (loom).')," & _
"('0','844700','Mesin rajut, mesin stitch-bonding dan mesin untuk membuat benang berpalut, tulle, rend bordir, perapih, jalinan atau jaring dan mesin pembuat rumbai.')," & _
"('0','844800','Mesin pembantu untuk digunakan dengan mesin dari pos 84.44, 84.45, 84.46 atau 84.47 (misalny dobi, Jacquard, penghenti gerak otomatis, mekanisme pengubah puntalan) bagian dan aksesori yang cocok digunakan semata-mata atau terutama dengan mesin dalam pos ini atau dari pos 84.44, 84.45, 84.46 atau 84.47 (misalny spindel dan spindel flyer, card clothing, sisir, extruding nipple, puntalan, heald dan heald-frame, jarum rajut).')," & _
"('0','844900','Mesin untuk membuat atau merampungkan bulu kempa atau bukan tenunan dalam lembaran atau bentuk, termasuk mesin untuk membuat topi bulu kempa blok untuk membuat topi.')," & _
"('0','845000','Mesin cuci tipe rumah tangga atau binatu, termasuk mesin yang dapat digunakan untuk mencuci dan mengeringkan.')," & _
"('0','845100','Mesin (selain mesin dari pos 84.50) untuk mencuci, membersihkan, memeras, mengeringkan, menyetrik mengepres (termasuk pengepres fusi), mengelantang, mencelup, menat merampungkan, melapisi atau meresapi benang tekstil, kain atau barang tekstil sudah jadi dan mesin untuk memberi pasta pada kain dasar atau kain dasar lainnya yang digunakan dalam pembuatan penutup lantai seperti linoleum mesin untuk menggulung, melepas gulungan, melipat, memotong atau memotong bergerigi kain tekstil.')," & _
"('0','845200','Mesin jahit, selain dari mesin penjahit buku yang dimaksud dari pos 84.40 perabotan, dasar dan tutup dirancang secara khusus untuk mesin jahit jarum mesin jahit.')," & _
"('0','845300','Mesin untuk mengolah, menyamak atau mengerjakan jangat, kulit atau kulit samak atau untuk membuat atau memperbaiki alas kaki atau barang lain dari jangat, kulit atau kulit samak, selain mesin jahit.')," & _
"('0','845400','Converter, ladle, mesin cetakan ingot dan mesin tuang, dari jenis yang digunakan dalam metalurgi atau dalam pengecoran logam.')," & _
"('0','845500','Kilang pencanai logam dan gilingannya.')," & _
"('0','845600','Mesin perkakas untuk mengerjakan berbagai bahan dengan penghilangan bahan, melalui proses penyinaran laser atau sinar lainnya atau sinar foton, ultrasonik, pelucutan elektro, elektro kimi sinar elektron, sinar ionik atau busur plasma mesin pemotong water-jet.')," & _
"('0','845700','Machining center, mesin konstruksi unit (single station) dan mesin transfer multi-station, untuk mengerjakan logam.')," & _
"('0','845800','Mesin bubut (termasuk turning centre) untuk menghilangkan logam.')," & _
"('0','845900','Mesin perkakas (termasuk mesin way-type unit head) untuk mengebor, menggurdi, menggiling, membuat ulir atau alur dengan menghilangkan logam, selain mesin bubut (termasuk turning centre) dari pos 84.58.')," & _
"('0','846000','Mesin perkakas untuk menghaluskan, menajamkan, menggerind menggosok, mengasah, memoles atau merampungkan logam atau cermet secara lain dengan memakai batu gerind amplas atau produk pemoles, selain mesin pemotong gir, penggerinda gir atau mesin untuk merampungkan gir dari pos 84.61.')," & _
"('0','846100','Mesin perkakas untuk mengetam, membentuk, menyerut, menggerek, memotong gir, menggerinda gir atau merampungkan gir, menggergaji, memotong dan mesin perkakas lainnya yang bekerja dengan menghilangkan logam atau sermet, tidak dirinci atau termasuk dalam pos lainnya.')," & _
"('0','846200','Mesin perkakas (termasuk pengepres) untuk mengerjakan logam dengan menemp memalu atau menempa dengan cetakan (tidak termasuk rolling mill) mesin perkakas (termasuk pengepres, slitting line dan cut-to-length line) untuk mengerjakan logam dengan membengkokkan, melipat, meluruskan, memipihkan, menggunting, melubangi, menakik atau menggigit (tidak termasuk bangku tarik) pengepres untuk mengerjakan logam atau karbida logam, tidak dirinci di atas.')," & _
"('0','846300','Mesin perkakas lainnya untuk mengerjakan logam atau sermet, tanpa menghilangkan bahannya.')," & _
"('0','846400','Mesin perkakas untuk mengerjakan batu, keramik, beton, asbes semen atau bahan mineral sejenisnya atau untuk mengerjakan kaca secara dingin.')," & _
"('0','846500','Mesin perkakas untuk mengerjakan batu, keramik, beton, asbes semen atau bahan mineral sejenisnya atau untuk mengerjakan kaca secara dingin.')," & _
"('0','846600','Bagian dan aksesori yang cocok untuk digunakan semata-mata atau terutama dengan mesin dari pos 84.56 sampai dengan 84.65, termasuk pemegang barang yang dikerjakan atau pemegang perkakas, self-opening die head, dividing head dan kelengkapan khusus lain untuk mesin perkakas pemegang perkakas untuk semua tipe perkakas tangan.')," & _
"('0','846700','Perkakas tangan, digerakkan secara pneumatik, hidraulik atau dengan motor listrik atau bukan listrik terpasang.')," & _
"('0','846800','Mesin dan aparatus untuk menyolder, mematri, mengelas, dapat memotong maupun tidak, selain yang dimaksud dari pos 85.15 mesin dan peralatan untuk mengeraskan permukaan yang digerakkan dengan gas.')," & _
"('0','847000','Mesin hitung dan mesin perekam, pereproduksi dan penayang data ukuran saku dengan fungsi menghitung mesin akunting, mesin cap pengganti perangko, mesin penerbit karcis dan mesin semacam itu, yang dilengkapi dengan alat hitung register kas.')," & _
"('0','847100','Mesin pengolah data otomatis dan unitnya pembaca magnetik atau optik, mesin untuk menyalin data pada media data dalam bentuk kode dan mesin untuk mengolah data tersebut, tidak dirinci atau termasuk dalam pos lainnya.')," & _
"('0','847200','Mesin kantor lainnya (misalny hektograf atau mesin duplikasi stensil, mesin pencetak alamat, mesin penyedia uang kertas otomatis, mesin penyortir koin, mesin penghitung atau pembungkus koin, mesin peruncing pensil, mesin pembuat lubang atau mesin kokot).')," & _
"('0','847300','Bagian dan aksesori (selain penutup, kotak pembawa dan sejenisnya) cocok untuk digunakan semata-mata atau terutama dengan mesin dari pos 84.70 sampai dengan 84.72.')," & _
"('0','847400','Mesin untuk menyortir, mengayak, memisahkan, mencuci, menghancurkan, menggerind mencampur atau mengaduk tanah, batu, bijih atau bahan mineral lainny dalam bentuk padat (termasuk bubuk atau pasta) mesin untuk aglomerasi, membentuk atau mencetak bahan bakar mineral padat, pasta keramik, semen yang tidak dikeraskan, bahan plesteran atau produk mineral lainnya dalam bentuk bubuk atau pasta mesin untuk membentuk cetakan dari pasir.')," & _
"('0','847500','Mesin untuk merakit lampu listrik atau elektronik, tabung atau katup atau bola lampu kilat, dalam glass envelope mesin untuk membuat atau mengerjakan kaca atau barang dari kaca secara panas.')," & _
"('0','847600','Mesin penjual barang otomatis (misalny mesin perangko, rokok, makanan atau minuman), termasuk mesin penukar uang.')," & _
"('0','847700','Mesin untuk mengerjakan karet atau plastik atau untuk pembuatan produk dari bahan tersebut, tidak dirinci atau termasuk dalam pos lainnya dalam Bab ini.')," & _
"('0','847800','Mesin untuk mengolah atau membuat tembakau menjadi barang jadi, tidak dirinci atau termasuk dalam pos lainnya dalam Bab ini.')," & _
"('0','847900','Mesin dan peralatan mekanis, mempunyai fungsi tersendiri, tidak dirinci atau termasuk dalam pos lainnya dalam Bab ini.')," & _
"('0','848000','Kotak cetakan untuk pengecoran logam dasar cetakan pola cetakan cetakan untuk logam (selain cetakan ingot), karbida logam, kac bahan mineral, karet atau plastik.')," & _
"('0','848100','Keran, klep, katup dan peralatan semacam itu untuk pip dinding ketel uap, tangki, tong atau sejenisny termasuk katup pengurang tekanan dan katup yang dikendalikan secara termostatik.')," & _
"('0','848200','Bantalan peluru atau bantalan gulung.')," & _
"('0','848300','Poros transmisi (termasuk poros berputar dan poros engkol) dan engkol rumah bantalan dan bantalan poros polos gir dan gearing ball screw atau roller screw gear box dan pengubah kecepatan lainny termasuk torak konverter roda gaya dan puli, termasuk blok puli kopling dan poros penyambung (termasuk sambungan universal).')," & _
"('0','848400','Gasket dan sambungan semacam itu dari lembaran logam yang dikombinasi dengan bahan lain atau dari dua atau lebih lapisan logam set atau aneka gasket dan sambungan semacam itu, tidak sama komposisiny disiapkan dalam kantong, amplop atau kemasan semacam itu sil mekanis.')," & _
"('0','848500','Mesin untuk additive manufacturing.')," & _
"('0','848600','Mesin dan aparatus dari jenis yang semata-mata atau terutama untuk pembuatan boule atau wafer semikonduktor, peralatan semikonduktor, sirkuit elektronik terpadu atau panel layar datar mesin dan aparatus yang dirinci pada Catatan 11 (C) Bab ini bagian dan aksesori.')," & _
"('0','848700','Bagian mesin, tidak mengandung konektor elektrik, isolator, koil, kontak atau bagian elektrik lainny tidak dirinci atau termasuk pada pos lainnya dalam Bab ini.')," & _
"('0','850000','Mesin dan perlengkapan elektris serta bagiannya perekam dan pereproduksi suar perekam dan pereproduksi gambar dan suara televisi, dan bagian serta aksesori dari barang tersebut')," & _
"('0','850100','Motor dan generator listrik (tidak termasuk perangkat pembangkit tenaga listrik).')," & _
"('0','850200','Perangkat pembangkit tenaga listrik dan konverter berputar.')," & _
"('0','850300','Bagian yang cocok untuk digunakan semata-mata atau terutama dengan mesin dari pos 85.01 atau 85.02.')," & _
"('0','850400','Transformator elektrik, konverter statis (misalny rectifier) dan induktor.')," & _
"('0','850500','Elektro magnet magnet permanen dan barang untuk dijadikan magnet permanen setelah diberi gaya magnet chuck, klem dan peralatan pemegang semacam yang bekerja secara elektro magnetis atau magnet permanen penyambung, kopling dan rem elektro magnetis kepala pengangkat elektro magnetis.')," & _
"('0','850600','Sel primer dan baterai primer.')," & _
"('0','850700','Akumulator listrik, termasuk separatorny empat persegi panjang (termasuk bujur sangkar) maupun tidak.')," & _
"('0','850800','Vacuum cleaner.')," & _
"('0','850900','Peralatan rumah tangga mekanik elektrik dengan motor listrik terpasang, selain vacuum cleaner dari pos 85.08.')," & _
"('0','851000','Alat cukur, alat pangkas rambut dan peralatan penghilang rambut, dengan motor listrik terpasang.')," & _
"('0','851100','Alat penyala atau penghidup elektrik dari jenis yang digunakan untuk mesin pembakaran dalam cetus api atau nyala kompresi (misalny magnet penyal magnet-dinamo, koil penyal busi pencetus dan busi pijar, starter motor) generator (misalny dinamo, alternator) dan sakelar dari jenis yang digunakan dengan mesin tersebut.')," & _
"('0','851200','Perlengkapan penerangan atau perlengkapan pemberi isyarat elektrik (tidak termasuk barang dari pos 85.39), penghapus kac pencair es dan penghilang kabut, dari jenis yang digunakan untuk sepeda atau kendaraan bermotor.')," & _
"('0','851300','Lampu listrik portabel yang dirancang untuk berfungsi dengan sumber energinya sendiri (misalny baterai kering, akumulator, magnet), selain perlengkapan penerangan dari pos 85.12.')," & _
"('0','851400','Tanur dan oven listrik industri atau laboratorium (termasuk yang difungsikan dengan induksi atau dielectric loss) perlengkapan industri atau laboratorium lainnya untuk pengolahan panas bahan dengan induksi atau dielectric loss.')," & _
"('0','851500','Mesin dan aparatus penyolder, pematri atau pengelas listrik (termasuk gas dipanaskan secara elektrik), sinar laser atau sinar lainnya atau sinar foton, ultrasonik, sinar elektron, pulsa magnetis atau busur plasm dapat memotong maupun tidak mesin dan aparatus listrik untuk penyemprotan panas logam atau sermet.')," & _
"('0','851600','Pemanas air instan atau pemanas air dengan tempat penyimpanan dan pemanas celup, listrik aparatus pemanas ruangan dan pemanas tanah, listrik aparatus penata rambut elektro-termal (misalny pengering rambut, pengeriting rambut, pemanas jepit untuk mengeriting rambut) dan pengering tangan setrika listrik peralatan elektro-termal lainnya dari jenis yang digunakan untuk keperluan rumah tangga resistor panas listrik, selain yang dimaksud dari pos 85.45.')," & _
"('0','851700','Perangkat telepon, termasuk smartphone dan telepon lainnya untuk jaringan seluler atau untuk jaringan tanpa kabel lainnya aparatus lainnya untuk mengirimkan atau menerima suar gambar, atau data lainnya termasuk aparatus untuk komunikasi dalam jaringan kabel atau tanpa kabel (seperti local atau wide area network), selain dari aparat transmisi atau penerima dari pos 84.43, 85.25, 85.27 atau 85.28.')," & _
"('0','851800','Mikrofon dan penyangganya pengeras suar dipasang pada rumahnya maupun tidak headphone dan earphone, dikombinasikan dengan mikrofon maupun tidak dan set yang terdiri dari satu mikrofon dan satu atau lebih pengeras suara amplifier listrik audio-frequency set amplifier suara listrik.')," & _
"('0','851900','Aparatus untuk perekam atau reproduksi suara.')," & _
"('0','852100','Aparatus perekam atau pereproduksi video, digabung dengan video tuner maupun tidak.')," & _
"('0','852200','Bagian dan aksesori yang cocok untuk digunakan semata-mata atau terutama dengan aparatus dari pos 85.19 sampai dengan 85.21.')," & _
"('0','852300','Cakram, pit media penyimpan non-volatile berbentuk padat, smart card dan media lain untuk merekam suara atau fenomena lain')," & _
"('0','852400','Modul panel layar datar, dilengkapi dengan layar peka sentuh maupun tidak.')," & _
"('0','852500','Aparatus transmisi untuk penyiaran radio atau televisi digabung dengan aparatus penerima atau dengan aparatus perekam suara maupun tidak kamera televisi, kamera digital dan kamera perekam video.')," & _
"('0','852600','Aparatus radar, aparatus pembantu radio navigasi dan aparatus radio kendali jarak jauh.')," & _
"('0','852700','Aparatus penerima untuk penyiaran radio, dikombinasi maupun tidak, dalam rumah yang sam dengan aparatus perekam atau pereproduksi suara atau penunjuk waktu.')," & _
"('0','852800','Monitor dan proyektor, tidak digabung dengan aparatus penerima televisi aparatus penerima untuk televisi, digabung dengan penerima siaran radio atau aparatus perekam atau pereproduksi suara atau video, maupun tidak.')," & _
"('0','852900','Bagian yang cocok untuk digunakan semata-mata atau terutama dengan aparatus dari pos 85.24 sampai dengan 85.28.')," & _
"('0','853000','Perlengkapan isyarat, pengontrol keselamatan atau lalu lintas elektrik untuk rel kereta api, trem, jalan, lalu lintas air, fasilitas parkir, instalasi pelabuhan atau lapangan udara (selain yang dimaksud dari pos 86.08).')," & _
"('0','853100','Aparatus isyarat suara atau visual elektrik (misalny bel, sirene, panel indikator, alarm pencuri atau kebakaran), selain yang dimaksud dari pos 85.12 atau 85.30.')," & _
"('0','853200','Kapasitor listrik, tetap, variabel atau dapat disesuaikan (pre-set).')," & _
"('0','853300','Resistor listrik (termasuk reostat dan potensiometer), selain resistor panas.')," & _
"('0','853400','Printed circuit.')," & _
"('0','853500','Aparatus listrik untuk memindahkan atau melindungi sirkuit listrik, atau untuk membuat koneksi ke atau pada sirkuit listrik (misalny sakelar, sekering, penangkal petir, pembatas voltase, penekan kenaikan tegangan, steker dan konektor lainny kotak penyambung), untuk voltase melebihi 1.000 volt.')," & _
"('0','853600','Aparatus listrik untuk memindahkan atau melindungi sirkuit listrik, atau untuk membuat koneksi ke atau pada sirkuit listrik (misalny sakelar, relai, sekering, penekan kenaikan tegangan, steker, soket, gagang lampu dan konektor lainny kotak penyambung), untuk voltase tidak melebihi 1.000 volt konektor untuk serat optik, bundel atau kabel serat optik.')," & _
"('0','853700','Papan, panel, konsol, mej kabinet dan landasan lainny dilengkapi dengan dua atau lebih aparatus dari pos 85.35 atau 85.36, untuk pengontrol listrik atau pendistribusi listrik, termasuk yang dilengkapi dengan instrumen atau aparatus dari Bab 90, dan aparatus pengontrol numerik, selain aparatus sakelar dari pos 85.17.')," & _
"('0','853800','Bagian yang cocok untuk digunakan semata-mata atau terutama dengan aparatus dari pos 85.35, 85.36 atau 85.37.')," & _
"('0','853900','Lampu filamen atau lampu discharge listrik, termasuk unit lampu sealed beam dan lampu ultra violet atau infra merah lampu busur sumber cahaya light-emitting diode (LED).')," & _
"('0','854000','Katup dan tabung termionik, katoda dingin atau katup dan tabung katoda foto (misalny katup dan tabung hampa atau berisi uap atau gas, katup dan tabung busur air raksa pengubah arus, tabung sinar katod tabung kamera televisi).')," & _
"('0','854100','Peralatan semi konduktor (sebagai contoh, diod transistor, transduser berbasis semi konduktor) peralatan semi konduktor peka cahay termasuk sel fotovoltaik dirakit menjadi modul atau dibuat menjadi panel maupun tidak light-emitting diode (LED), dirakit dengan light-emitting diode (LED) lainnya maupun tidak kristal piezoelektrik terpasang.')," & _
"('0','854200','Sirkuit elektronik terpadu.')," & _
"('0','854300','Mesin dan aparatus elektrik, mempunyai fungsi tersendiri, tidak dirinci atau termasuk pada pos lainnya dalam Bab ini.')," & _
"('0','854400','Kawat diisolasi (termasuk dienamel atau dianodisasi), kabel (termasuk kabel koaksial) dan konduktor listrik diisolasi lainny dilengkapi dengan konektor maupun tidak kabel serat optik, dibuat dari serat berselubung tersendiri, dirakit dengan konduktor listrik atau dilengkapi dengan konektor maupun tidak.')," & _
"('0','854500','Elektroda karbon, sikat karbon, karbon lampu, karbon baterai dan barang lainnya dari grafit atau karbon lainny dengan atau tanpa logam, jenis yang digunakan untuk keperluan elektrik.')," & _
"('0','854600','Isolator listrik dari berbagai bahan.')," & _
"('0','854700','Alat kelengkapan pengisolasi untuk mesin, peralatan atau perlengkapan elektrik, yang seluruhnya merupakan bahan isolasi yang terpisah dari sejumlah kecil komponen logam (misalny soket berulir) yang semata-mata dipasang selama pencetakan untuk keperluan perakitan, selain isolator yang dimaksud dari pos 85.46 pipa saluran listrik beserta sambungannya dari logam tidak mulia dilapisi bahan pengisolasi.')," & _
"('0','854800','Bagian elektrik dari mesin atau aparatus, tidak dirinci atau termasuk pada pos lainnya dalam Bab ini.')," & _
"('0','854900','Sisa dan skrap elektrik dan elektronik.')," & _
"('0','860000','Lokomotif kereta api atau trem, kendaraan yang bergerak di atas rel dan bagiannya track fixture dan alat kelengkapan rel kereta api atau trem serta bagiannya perlengkapan pemberi isyarat lalu-lintas mekanis (termasuk elektro-mekanis) dari segala jenis.')," & _
"('0','860100','Lokomotif rel digerakkan dengan sumber tenaga listrik dari luar atau dengan akumulator listrik.')," & _
"('0','860200','Lokomotif rel lainnya tender lokomotif.')," & _
"('0','860300','Gerbong tertutup dan gerbong terbuka kereta api atau trem berdaya gerak sendiri, selain yang dimaksud dari pos 86.04.')," & _
"('0','860400','Kendaraan perawatan atau perbaikan kereta api atau trem, berdaya gerak sendiri maupun tidak (misalny kendaraan bengkel, crane, balas penumbuk, pembuat jalur rel, gerbong penguji dan pemeriksa rel).')," & _
"('0','860500','Gerbong kereta api atau trem penumpang tidak berdaya gerak sendiri gerbong bagasi, gerbong barang pos dan gerbong kereta api atau trem untuk keperluan khusus lainny tidak berdaya gerak sendiri (tidak termasuk yang dimaksud dari pos 86.04).')," & _
"('0','860600','Gerbong tertutup dan wagon kereta api atau trem untuk barang, tidak berdaya gerak sendiri.')," & _
"('0','860700','Bagian dari lokomotif atau kendaraan rel kereta api atau trem atau gerbong.')," & _
"('0','860800','Track fixture dan alat kelengkapan rel kereta api atau trem perlengkapan pemberi isyarat, pengontrol keselamatan atau lalu lintas mekanis (termasuk elektro mekanis) untuk rel kereta api, rel trem, jalan, lalu lintas perairan darat, fasilitas parkir, instalasi pelabuhan atau lapangan udara bagiannya.')," & _
"('0','860900','Peti kemas (termasuk peti kemas untuk pengangkutan barang cair) dirancang dan dilengkapi secara khusus untuk dibawa dengan satu jenis atau lebih moda pengangkut.')," & _
"('0','870000','Kendaraan selain yang bergerak di atas rel kereta api atau trem, dan bagian serta aksesorinya')," & _
"('0','870100','Traktor (selain traktor dari pos 87.09).')," & _
"('0','870200','Kendaraan bermotor untuk pengangkutan sepuluh orang atau lebih, termasuk pengemudi.')," & _
"('0','870300','Mobil dan kendaraan bermotor lainnya terutama dirancang untuk pengangkutan orang (selain yang dimaksud dari pos 87.02), termasuk station wagon dan mobil balap.')," & _
"('0','870400','Kendaraan bermotor untuk pengangkutan barang.')," & _
"('0','870500','Kendaraan bermotor untuk keperluan khusus, selain yang terutama dirancang untuk pengangkutan orang atau barang (misalny lori derek, lori crane, kendaraan pemadam kebakaran, lori pencampur beton, lori penyapu jalan, lori penyemprot, mobil bengkel, mobil unit radiologi).')," & _
"('0','870600','Sasis dilengkapi dengan mesin, untuk kendaraan bermotor dari pos 87.01 sampai dengan 87.05.')," & _
"('0','870700','Bodi (termasuk kabin), untuk kendaraan bermotor dari pos 87.01 sampai dengan 87.05.')," & _
"('0','870800','Bagian dan aksesori kendaraan bermotor dari pos 87.01 sampai dengan 87.05.')," & _
"('0','870900','Truk kerj berdaya gerak sendiri, tidak dilengkapi dengan perlengkapan pengangkat atau penanganan, dari tipe yang digunakan di pabrik, gudang, area dok atau bandar udara untuk pengangkutan barang jarak pendek traktor dari tipe yang digunakan di stasiun kereta api bagian dari kendaraan tersebut.')," & _
"('0','871000','Tank dan kendaraan tempur lapis baja lainny bermotor, dilengkapi dengan senjata maupun tidak, dan bagian dari kendaraan tersebut.')," & _
"('0','871100','Sepeda motor (termasuk moped) dan sepeda yang dilengkapi dengan motor bantu, dengan atau tanpa kereta samping kereta samping.')," & _
"('0','871200','Sepeda roda dua dan sepeda lainnya (termasuk sepeda roda tiga untuk pengantar), tidak bermotor.')," & _
"('0','871300','Kendaraan untuk orang cacat, bermotor maupun tidak atau digerakkan secara mekanik lainnya.')," & _
"('0','871400','Bagian dan aksesori kendaraan dari pos 87.11 sampai dengan 87.13.')," & _
"('0','871500','Kereta bayi dan bagiannya.')," & _
"('0','871600','Trailer dan semi trailer kendaraan lainny tidak digerakkan secara mekanik bagiannya.')," & _
"('0','880000','Kendaraan udar kendaraan luar angkas dan bagiannya')," & _
"('0','880100','Balon udara dan balon udara yang dapat dikemudikan pesawat layang, pesawat layang gantung dan kendaraan udara lainnya tanpa tenaga penggerak.')," & _
"('0','880200','Kendaraan udara lainnya (misalny helikopter, pesawat udara), kecuali kendaraan udara tak berawak dari pos 88.06 kendaraan luar angkasa (termasuk satelit) serta kendaraan peluncur luar angkasa dan sub orbital.')," & _
"('0','880400','Parasut (termasuk parasut dan paraglider yang dapat dikemudikan) dan rotochute bagian dan aksesorinya.')," & _
"('0','880500','Gir peluncur kendaraan udara deck-arrestor atau gir semacam itu pesawat latih terbang di darat bagian dari barang tersebut.')," & _
"('0','880600','Kendaraan udara tak berawak.')," & _
"('0','880700','Bagian dari barang pada pos 88.01, 88.02 atau 88.06.')," & _
"('0','890000','Kapal, perahu dan struktur terapung serta bangunan')," & _
"('0','890100','Kapal pesiar, kapal ekskursi, kapal feri, kapal kargo, tongkang dan kendaraan air semacam itu untuk pengangkutan orang atau barang.')," & _
"('0','890200','Kapal penangkap ikan kapal pabrik dan kendaraan air lainnya untuk pemrosesan atau pengawetan produk perikanan.')," & _
"('0','890300','Yacht dan kendaraan air lainnya untuk pelesir atau olah raga sampan dan kano.')," & _
"('0','890400','Kapal penarik dan pendorong.')," & _
"('0','890500','Kapal suar, kapal pemadam kebakaran, kapal keruk, crane terapung dan kendaraan air lainnya yang fungsi berlayarnya bukan merupakan fungsi utama dok terapung platform pengeboran atau produksi terapung atau di bawah air.')," & _
"('0','890600','Kendaraan air lainny termasuk kapal perang dan perahu penyelamat selain sampan.')," & _
"('0','890700','Struktur terapung lainnya (misalny rakit, tangki, coffer-dam, landasan apung, pelampung suar dan rambu laut).')," & _
"('0','890800','Bangunan tempat tinggal')," & _
"('0','890900','Bangunan bukan tempat tinggal')," & _
"('0','891000','Jalan bebas hambatan (kecuali jalan bebas hambatan yang ditinggikan/jalan layang), jalan ray jalan, jalan kereta api dan landasan pesawat terbang')," & _
"('0','891100','Jembatan, jalan layang bebas hambatan dan terowongan')," & _
"('0','891200','Pelabuhan, aliran air/terusan, bendungan, irigasi dan bangunan air lainnya')," & _
"('0','891300','Pipa saluran jarak jauh, saluran komunikasi dan saluran listrik (kabel)')," & _
"('0','891400','Kabel dan pipa saluran lokal dan bangunan')," & _
"('0','891500','Bangunan pertambangan dan industri')," & _
"('0','891600','Fasilitas olah raga dan rekreasi di ruang terbuka')," & _
"('0','891700','Pekerjaan teknik sipil lainnya')," & _
"('0','900000','lnstrumen dan aparatus optis, fotografi, sinematografi, pengukur, pemeriksa presisi medis ctan bedah bagian dan aksesorinya')," & _
"('0','900100','Serat optik dan bundel serat optik, kabel serat optik selain yang dimaksud dari pos 85.44 lembaran dan pelat dari bahan polarisasi lensa (termasuk lensa kontak), prism cermin dan elemen optik lainny dari berbagai bahan, tidak terpasang, selain elemen semacam itu dari kaca tidak dikerjakan secara optik.')," & _
"('0','900200','Lens prism cermin dan elemen optik lain-ny dari berbagai bahan, terpasang, sebagai bagian dari atau alat kelengkapan untuk instrumen atau aparatus, selain elemen semacam itu dari kaca tidak dikerjakan secara optik.')," & _
"('0','900300','Bingkai dan mounting untuk kacamat kacamata pelindung atau sejenisny serta bagiannya.')," & _
"('0','900400','Kacamat kacamata pelindung dan sejenisny korektif, protektif atau lainnya.')," & _
"('0','900500','Teropong gand teropong tunggal, teleskop optik lainny dan mountingnya instrumen astronomi lainnya dan mountingny tetapi tidak termasuk instrumen untuk radio astronomi.')," & _
"('0','900600','Kamera fotografi (selain kamera sinematografi) aparatus lampu kilat fotografi dan bola lampu kilat selain lampu tabung dari pos 85.39.')," & _
"('0','900700','Kamera dan proyektor sinematografi dilengkapi dengan aparatus perekam atau reproduksi suara maupun tidak.')," & _
"('0','900800','Proyektor gambar, selain sinematografi pembesar dan pengecil fotografi (selain sinematografi).')," & _
"('0','901000','Aparatus dan perlengkapan laboratorium fotografi (termasuk sinemetografi), tidak dirinci atau termasuk pos lainnya dari Bab ini negatoscope layar proyeksi.')," & _
"('0','901100','Mikroskop optik gabungan, termasuk yang untuk fotomikrografi, sinefotomikrografi atau mikroproyeksi.')," & _
"('0','901200','Mikroskop selain mikroskop optik aparatus difraksi.')," & _
"('0','901300','Laser, selain dioda laser peralatan dan instrumen optik lainny tidak dirinci atau termasuk pos lainnya dari Bab ini.')," & _
"('0','901400','Kompas pencari arah instrumen dan peralatan navigasi lainnya.')," & _
"('0','901500','Instrumen dan peralatan survei (termasuk survei photogrammetrical), hidrografi, oseanografi, hidrologi, meteorologi atau geofisik tidak termasuk kompas pengukur jarak.')," & _
"('0','901600','Timbangan dengan kepekaan 5 cg atau lebih pek dengan atau tanpa anak timbangannya.')," & _
"('0','901700','Instrumen penggambar, pemberi tanda atau penghitung matematis (misalny mesin perancang, pantograf, protraktor, set gambar, mistar hitung, disc calculator) instrumen untuk mengukur panjang, digunakan dengan tangan (misalny batang dan pita pengukur, mikro meter, kaliper), tidak dirinci atau termasuk dalam pos lain dari Bab ini.')," & _
"('0','901800','Instrumen dan peralatan yang digunakan dalam ilmu medis, bedah, perawatan gigi atau kedokteran hewan, termasuk aparatus scintigraphic, aparatus elektro medis dan instrumen penguji penglihatan lainnya.')," & _
"('0','901900','Peralatan mekano terapi aparatus pijat aparatus penguji kecerdasan psikologis aparatus terapi ozon, terapi oksigen, terapi aerosol, nafas buatan atau pernapasan terapeutik lainnya.')," & _
"('0','902000','Peralatan nafas dan masker gas lainny tidak termasuk masker pelindung yang tidak mempunyai bagian mekanis maupun filter yang dapat diganti.')," & _
"('0','902100','Peralatan ortopedik, termasuk kruk bebat bedah dan bebat hernia bidai dan peralatan patah tulang lainnya anggota badan artifisial alat bantu dengar dan peralatan lainnya yang dipakai atau dibaw atau ditanam dalam tubuh, untuk mengganti yang cacat atau tidak mampu.')," & _
"('0','902200','Aparatus yang didasarkan atas penggunaan sinar X atau radiasi sinar alf bet gamma atau pengion lainnya untuk keperluan medis, pembedahan, perawatan gigi atau kedokteran hewan, maupun tidak, termasuk aparatus radiografi atau radioterapi, tabung sinar X dan generator sinar X lainny generator tegangan tinggi, panel dan meja kontrol, tirai, meja pemeriksaan atau perawatan, kursi dan sejenisnya.')," & _
"('0','902300','Instrumen, aparatus dan model, dirancang untuk keperluan peragaan (misalnya dalam pendidikan atau pameran), tidak cocok untuk keperluan lainnya.')," & _
"('0','902400','Mesin dan peralatan untuk menguji kekerasan, kekuatan, kompresibilitas, elastisitas atau sifat mekanis lainnya dari suatu bahan (misalny logam, kayu, tekstil, kertas, plastik).')," & _
"('0','902500','Hidrometer dan instrumen apung semacam itu, termometer, pirometer, barometer, higrometer dan psikrometer, dengan perekam atau tidak dan berbagai kombinasi dari instrumen tersebut.')," & _
"('0','902600','Instrumen dan aparatus untuk mengukur atau memeriksa aliran, tinggi permukaan, tekanan atau variabel lainnya dari cairan atau gas (misalny pengukur arus, pengukur tinggi permukaan, manometer, pengukur panas), tidak termasuk instrumen dan aparatus dari pos 90.14, 90.15, 90.28 atau 90.32.')," & _
"('0','902700','Instrumen dan aparatus untuk analisa sifat fisika atau kimia (misalnya polarimeter, refraktometer, spektrometer, aparatus analisa gas atau asap) instrumen dan aparatus untuk mengukur atau memeriksa viskositas, porositas, daya muai, tegangan permukaan atau sejenisnya instrumen dan aparatus untuk mengukur atau memeriksa kuantitas panas, suara atau cahaya (termasuk pengukur cahaya) mikrotom.')," & _
"('0','902800','Pengukur pasokan atau produksi gas, cairan atau listrik, termasuk pengukur kalibrasinya.')," & _
"('0','902900','Penghitung putaran, penghitung produksi, taksimeter, pengukur jarak, pedometer dan sejenisnya indikator kecepatan dan takometer, selain barang dari pos 90.14 atau 90.15 stroboskop.')," & _
"('0','903000','Osiloskop, penganalisa spektrum dan instrumen serta aparatus lainnya untuk mengukur atau memeriksa kuantitas elektrik, tidak termasuk pengukur dari pos 90.28 instrumen dan aparatus untuk mengukur atau mendeteksi sinar alf bet gamm sinar X, kosmik atau radiasi pengion lainnya.')," & _
"('0','903100','Instrumen, peralatan dan mesin pengukur atau pemeriks tidak dirinci atau termasuk dalam pos manapun dalam Bab ini proyektor profil.')," & _
"('0','903200','Instrumen dan aparatus pengatur atau pengontrol otomatis.')," & _
"('0','903300','Bagian dan aksesori (tidak dirinci atau termasuk dalam pos lainnya pada Bab ini) untuk mesin, peralatan, instrumen atau aparatus dari Bab 90.')," & _
"('0','910000','Jam dan arloji serta bagiannya')," & _
"('0','910100','Arloji tangan, arloji saku dan arloji lainny termasuk penghitung detik, dengan badan arloji dari logam mulia atau dari logam yang dipalut dengan logam mulia.')," & _
"('0','910200','Arloji tangan, arloji saku dan arloji lainny termasuk penghitung detik, selain yang dimaksud dari pos 91.01.')," & _
"('0','910300','Jam dengan penggerak jam, tidak termasuk jam dari pos 91.04.')," & _
"('0','910400','Jam panel instrumen dan jam tipe semacam untuk kendaraan darat, kendaraan udar kendaraan luar angkasa atau kendaraan air.')," & _
"('0','910500','Jam lainnya.')," & _
"('0','910600','Aparatus perekam waktu dan aparatus untuk mengukur, merekam atau menunjukkan interval waktu secara lain, dengan penggerak jam atau penggerak arloji atau dengan motor sinkron (misalnya pencatat waktu, perekam waktu).')," & _
"('0','910700','Sakelar waktu dengan penggerak jam atau penggerak arloji atau dengan motor sinkron.')," & _
"('0','910800','Penggerak arloji, lengkap dan dirakit.')," & _
"('0','910900','Penggerak jam, lengkap dan dirakit.')," & _
"('0','911000','Penggerak arloji atau penggerak jam, lengkap, tidak dirakit atau dirakit sebagian (set penggerak) penggerak arloji atau penggerak jam, tidak lengkap, dirakit penggerak arloji atau penggerak jam tidak lengkap, tidak dirakit.')," & _
"('0','911100','Badan arloji dan bagiannya.')," & _
"('0','911200','Badan jam dan badan tipe semacam itu untuk barang lain pada Bab ini, dan bagiannya.')," & _
"('0','911300','Tali arloji, ban arloji dan gelang arloji, serta bagiannya.')," & _
"('0','911400','Bagian lainnya dari jam atau arloji.')," & _
"('0','920000','lnstrumen musik bagian dan aksesori barang tersebut')," & _
"('0','920100','Piano, termasuk piano otomatis harpsichord dan instrumen keyboard bersenar lainnya.')," & _
"('0','920200','Instrumen musik bersenar lainnya (misalny gitar, biol harpa).')," & _
"('0','920500','Instrumen musik tiup (misalny organ keyboard berpip akordeon, klarinet, trompet, bag pipe), selain organ fairground dan organ jalanan mekanis.')," & _
"('0','920600','Instrumen musik perkusi (misalny drum, xylophone, simbal, kastanyet, marakas).')," & _
"('0','920700','Instrumen musik, dengan suara yang dihasilkan, atau harus diperkuat, secara elektrik (misalny organ, gitar, akordeon).')," & _
"('0','920800','Kotak musik, fairground organ, organ jalanan mekanis, kicauan burung mekanis, gergaji musik dan instrumen musik lainnya yang tidak termasuk dalam pos manapun dari Bab ini suling pemikat dari segala jenis peluit, terompet panggil dan instrumen isyarat suara dengan tiupan mulut lainnya.')," & _
"('0','920900','Bagian (misalny mekanisme untuk kotak musik) dan aksesori (misalny kartu, cakram, dan gulungan untuk instrumen mekanik) dari instrumen musik metronom, garpu tala dan pipa penala dari segala jenis.')," & _
"('0','930000','Senjata dan amunisi bagian dan aksesorinya')," & _
"('0','930100','Senjata militer, selain revolver, pistol dan senjata dari pos 93.07.')," & _
"('0','930200','Revolver dan pistol, selain yang dimaksud dari pos 93.03 atau 93.04.')," & _
"('0','930300','Senjata api lainnya dan peralatan semacam itu yang dioperasikan dengan penembakan bahan peledak (misalny senapan lantak dan bedil olah rag senjata api yang diisi dari moncongny pistol isyarat dan peralatan lain yang dirancang hanya untuk menembakkan suar isyarat, pistol dan revolver untuk menembakkan peluru kosong, captive-bolt humane killer, senapan pelempar tali).')," & _
"('0','930400','Senjata lainnya (misalny senapan dan pistol per, angin atau gas, serta tongkat polisi), tidak termasuk yang dimaksud dari pos 93.07.')," & _
"('0','930500','Bagian dan aksesori dari barang dari pos 93.01 sampai dengan 93.04.')," & _
"('0','930600','Bom, granat, torpedo, ranjau, misil dan amunisi perang semacam itu serta bagiannya peluru dan amunisi lainnya serta proyektil dan bagianny termasuk sumbat mimis dan sumbat peluru.')," & _
"('0','930700','Pedang, pedang pendek, bayonet, tombak dan senjata semacam itu serta bagiannya dan sarung pedang serta sarung senjata lainnya.')," & _
"('0','940000','Perabotan keperluan tidur, kasur, alas kasur, bantalan kursi dan perabotan yang diisi semacam itu lampu dan alat kelengkapan penerangan, tidak dirinci atau termasuk dalam pos lain tanda iluminasi, papan nama iluminasi dan sejenisnya bangunan prapabrikasi')," & _
"('0','940100','Tempat duduk (selain barang yang dimaksud dari pos 94.02), dapat diubah menjadi tempat tidur maupun tidak, dan bagiannya.')," & _
"('0','940200','Perabotan medis, bedah, perawatan gigi atau kedokteran hewan (misalny meja operasi, meja pemeriksaan, tempat tidur rumah sakit dengan alat kelengkapan mekanis, kursi perawatan gigi) kursi cukur dan kursi semacam itu, yang mempunyai penggerak berputar, rebah dan naik turun bagian dari barang tersebut.')," & _
"('0','940300','Perabotan lain dan bagiannya.')," & _
"('0','940400','Alas kasur barang keperluan tidur dan perabotan semacam itu (misalny kasur, selimut tebal, eiderdown, bantalan kursi, pouffe dan bantal) dilengkapi dengan pegas atau diisi atau dilengkapi bagian dalamnya dengan berbagai bahan atau dengan karet atau plastik seluler, disarungi maupun tidak.')," & _
"('0','940500','Luminer dan alat kelengkapan penerangan termasuk searchlight dan lampu sorot serta bagianny tidak dirinci atau termasuk dalam pos lain tanda iluminasi, papan nama iluminasi dan sejenisny mempunyai sumber cahaya permanen, dan bagiannya yang tidak dirinci atau termasuk dalam pos lain.')," & _
"('0','940600','Bangunan prapabrikasi.')," & _
"('0','950000','Mainan, permainan dan keperluan olah raga bagian dan aksesorinya')," & _
"('0','950300','working or not puzzles of all kinds.')," & _
"('0','950400','Konsol dan mesin video game, permainan meja atau dalam ruangan, termasuk pintable, biliar, meja khusus untuk permainan kasino dan perlengkapan boling otomatis, mesin permainan yang dioperasikan dengan koin, uang kertas, kartu bank, token atau dengan alat pembayaran lainnya.')," & _
"('0','950500','Barang festival, karnaval atau hiburan lainny termasuk alat sulap dan lelucon.')," & _
"('0','950600','Barang dan perlengkapan untuk latihan fisik umum, gimnastik, atletik, olahraga lainnya (termasuk meja tenis) atau permainan luar ruangan, tidak dirinci atau termasuk dalam pos lain pada Bab ini kolam renang dan paddling pool.')," & _
"('0','950700','Joran, mata kail dan perlengkapan pancing lainnya jaring ikan, jaring kupu-kupu dan jaring semacam itu â€œburungâ€ pemikat (selain barang-barang dari pos 92.08 atau 97.05) dan perlengkapan berburu atau menembak semacam itu.')," & _
"('0','950800','Sirkus keliling dan travelling menagerie wahana taman hiburan dan hiburan taman air permainan taman hiburan, termasuk galeri tembak teater keliling.')," & _
"('0','960000','Bermacam-macam barang hasil pabrik')," & _
"('0','960100','Gading, tulang, tempurung kura-kur tanduk, tanduk rus koral, kulit kerang dan bahan ukiran hewani lainnya dikerjakan, serta barang dari bahan tersebut (termasuk barang yang diperoleh melalui pencetakan).')," & _
"('0','960200','Bahan ukiran nabati atau mineral dikerjakan dan barang dari bahan tersebut barang cetakan atau ukiran dari malam, dari stearin, dari getah alam atau resin alam atau dari pasta model, dan barang cetakan atau ukiran lainny yang tidak dirinci atau termasuk dalam pos lain gelatin tidak dikeraskan (kecuali gelatin dari pos 35.03) dikerjakan dan barang dari gelatin tidak dikeraskan.')," & _
"('0','960300','Sapu, sikat (termasuk sikat yang merupakan bagian dari mesin, peralatan atau kendaraan), penyapu lantai mekanis yang digerakkan dengan tangan, tidak bermotor, pengepel dan kemucing simpul dan jumbai yang disiapkan untuk pembuatan sapu atau sikat baki dan roll pengecat penyapu air (selain roller squeegee).')," & _
"('0','960400','Ayakan tangan dan saringan tangan.')," & _
"('0','960500','Set bepergian untuk keperluan rias pribadi, menjahit atau membersihkan sepatu atau pakaian.')," & _
"('0','960600','Kancing, kancing tekan, kancing jepret dan kancing pijat, bentukan kancing dan bagian lainnya dari barang tersebut button blank.')," & _
"('0','960700','Ritsleting dan bagiannya.')," & _
"('0','960800','Bolpoin pena dan marker felt tipped serta pena dan marker porous-tipped lainnya pulpen, pena stilograf dan pena lainnya duplicating stylo pensil putar atau pensil dorong tangkai pen tangkai pensil dan tangkai semacam itu bagian (termasuk tutup dan jepit) dari barang tersebut di atas, selain yang dimaksud dari pos 96.09.')," & _
"('0','960900','Pensil (selain pensil dari pos 96.08), crayon, inti pensil, pastel, arang gambar, kapur tulis atau gambar dan kapur penjahit.')," & _
"('0','961000','Sabak dan papan, permukaannya dapat dipakai untuk menulis atau menggambar, berbingkai maupun tidak.')," & _
"('0','961100','Stempel tanggal, meterai atau nomor, dan sejenisnya (termasuk peralatan untuk mencetak atau menghias timbul label), dirancang untuk digunakan dengan tangan batang penyusun yang digerakkan dengan tangan dan set pencetak dengan tangan yang dilengkapi batang penyusun.')," & _
"('0','961200','Pita mesin ketik atau pita semacam itu, bertinta atau disiapkan secara lain untuk memberi kesan, dalam kumparan atau cartridge maupun tidak bantalan tint bertinta maupun tidak, dengan atau tanpa kotak.')," & _
"('0','961300','Pemantik sigaret dan pemantik lainny mekanis atau elektrik maupun tidak, dan bagiannya selain batu api dan sumbu.')," & _
"('0','961400','Pipa untuk merokok (termasuk cangklong) dan pipa cerutu atau pipa sigaret, dan bagiannya.')," & _
"('0','961500','Sisir, jepitan perapih rambut dan sejenisnya tusuk rambut, jepitan pengikal, pengikal rambut, pengeriting rambut dan sejenisny selain yang dimaksud dari pos 85.16, dan bagiannya.')," & _
"('0','961600','Penyemprot wewangian dan penyemprot rias semacam itu, dan batang serta kepala penyemprot pengoles bedak dan bantalan untuk keperluan kosmetik atau preparat rias.')," & _
"('0','961700','Termos dan bejana hampa udara lainny lengkap bagiannya selain kaca bagian dalam.')," & _
"('0','961800','Model orang-orangan dan bagian tubuh lainnya untuk penjahit otomata dan barang peraga yang menarik lainnya untuk menghias jendela toko.')," & _
"('0','961900','Sanitary towel (pad) dan tampon saniter, serbet (popok), pembebat popok dan barang semacam itu, dari bahan apapun.')," & _
"('0','962000','Monopod, bipod, tripod dan barang semacam itu.')," & _
"('0','970000','Karya seni, barang kolektor dan barang antik')," & _
"('0','970100','Lukisan, gambar dan gambar pastel, dikerjakan seluruhnya dengan tangan, selain gambar dari pos 49.06 dan selain barang jadi yang dicat atau didekorasi dengan tangan kolase, mosaik dan plakat hiasan semacam itu.')," & _
"('0','970200','Ukiran, cetakan, dan litograf asli.')," & _
"('0','970300','Arca dan patung asli, dari berbagai bahan.')," & _
"('0','970400','Perangko atau materai, perangko dibubuhi cap pos, sampul hari pertam benda pos (kertas bermaterai), dan sejenisny bekas atau belum dipakai, selain yang dimaksud dari pos 49.07.')," & _
"('0','970500','Koleksi dan barang kolektor kepentingan arkeologi, etnografi, sejarah, zoologi, botani, mineralogi, anatomi, paleontologi atau numismatika.')," & _
"('0','970600','Barang antik yang umurnya melebihi 100 tahun.')," & _
"('0','980000','Ketentuan khusus untuk industri alat transportasi')," & _
"('0','980100','Kendaraan bermotor dalam keadaan terurai tidak lengkap sasis dengan mesin terpasang dalam keadaan terurai tidak lengkap.')," & _
"('0','980200','Komponen kendaraan bermotor dalam keadaan terurai tidak lengkap.')," & _
"('0','980400','Komponen kapal dari Bab 72.')," & _
"('0','980500','Komponen kapal dari Bab 73.')," & _
"('0','980600','Komponen kapal dari Bab 74.')," & _
"('0','980700','Komponen kapal dari Bab 76.')," & _
"('0','980800','Komponen kapal dari Bab 84.')," & _
"('0','980900','Komponen kapal dari Bab 85.')," & _
"('0','981000','Komponen kapal dari Bab 90.')," & _
"('0','981100','Komponen kapal dari Bab lainnya.')," & _
"('0','990000','Peranti lunak dan barang digital lainnya')," & _
"('0','990100','Peranti lunak dan barang digital lainnya yang ditransmisikan secara elektronik.')," & _
"('1','000000','Jasa')," & _
"('1','010000','Jasa Kontruksi')," & _
"('1','010100','Jasa Konstruksi Umum Untuk Bangunan')," & _
"('1','010101','Jasa konstruksi umum untuk bangunan tempat tinggal')," & _
"('1','010102','Jasa konstruksi umum untuk bangunan bukan tempat tinggal.')," & _
"('1','010200','Jasa Konstruksi Umum Untuk Pekerjaan Teknik Sipil')," & _
"('1','010201','Jasa konstruksi umum untuk jalan bebas hambatan (kecuali jalan layang), jalan raya, jalan, jalan kereta api dan landasan pesawat terbang')," & _
"('1','010202','Jasa konstruksi umum jembatan, jalan layang dan terowongan')," & _
"('1','010203','Jasa konstruksi umum pelabuhan, terusan, bendungan, bangunan air irigasi dan bangunan air lainnya')," & _
"('1','010204','Jasa konstruksi umum pipa saluran jarak jauh, saluran komunikasi dan listrik (kabel)')," & _
"('1','010205','Jasa konstruksi umum pipa saluran dan kabel lokal serta pekerjaan ybdi')," & _
"('1','010206','Jasa konstruksi umum pertambangan dan bangunan industri')," & _
"('1','010207','Jasa konstruksi umum fasilitas olah raga dan rekreasi di ruang terbuka')," & _
"('1','010209','Jasa konstruksi umum pekerjaan teknik sipil lainnya')," & _
"('1','010300','Jasa Penyiapan Tempat')," & _
"('1','010301','Jasa pembongkaran')," & _
"('1','010302','Jasa pembentukan dan pembersihan tempat')," & _
"('1','010303','Jasa penggalian dan pemindahan tanah')," & _
"('1','010304','Jasa pengeboran sumur air bersih dan pemasangan sistem septik')," & _
"('1','010400','Perakitan Dan Pembangunan Konstruksi Prafabrikan')," & _
"('1','010401','Perakitan Dan Pembangunan Konstruksi Prafabrikan')," & _
"('1','010500','Jasa Konstruksi Ketrampilan Khusus')," & _
"('1','010501','Jasa pemancangan tiang ke dalam tanah dan pondasi')," & _
"('1','010502','Jasa pemasangan kerangka konstruksi.')," & _
"('1','010503','Jasa pemasangan atap dan pembuatan tahan air')," & _
"('1','010504','Jasa pemasangan beton')," & _
"('1','010505','Jasa pembangunan struktur baja.')," & _
"('1','010506','Jasa tukang bangunan')," & _
"('1','010507','Jasa Perancah')," & _
"('1','010509','Jasa konstruksi ketrampilan khusus lainnya')," & _
"('1','010600','Jasa Instalasi')," & _
"('1','010601','Jasa instalasi listrik')," & _
"('1','010602','Jasa instalasi pipa air (plumbing) dan pemasangan saluran')," & _
"('1','010603','Jasa instalasi peralatan pemanas ruangan')," & _
"('1','010604','Jasa instalasi perlengkapan gas')," & _
"('1','010605','Jasa penyekatan')," & _
"('1','010609','Jasa instalasi lainnya')," & _
"('1','010700','Jasa Pekerjaan Akhir Dan Penyelesaian Gedung')," & _
"('1','010701','Jasa pemasangan kaca')," & _
"('1','010702','Jasa pelapisan plester')," & _
"('1','010703','Jasa pengecatan')," & _
"('1','010704','Jasa pemasangan ubin lantai dan dinding')," & _
"('1','010705','Jasa pemasangan lantai, penutupan dinding dan pelapisan dinding lainnya')," & _
"('1','010706','Jasa pekerjaan tukang kayu (Carpentry dan Joinery)')," & _
"('1','010707','Jasa pemasangan pagar dan susuran tangga')," & _
"('1','010709','Jasa pekerjaan akhir dan penyelesaian bangunan lainnya')," & _
"('1','020000','Jasa Perdagangan Besar')," & _
"('1','020100','Jasa Perdagangan Besar Kecuali Atas Dasar Balas Jasa Atau Kontrak')," & _
"('1','020101','Jasa perdagangan besar dari bahan mentah pertanian dan binatang hidup, kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020102','Jasa perdagangan besar dari makanan, minuman dan tembakau, kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020103','Jasa perdagangan besar dari tekstil, pakaian dan alas kaki, kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020104','Jasa perdagangan besar dari peralatan, barang-barang dan peralatan rumah tangga, kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020105','Jasa perdagangan besar dari barang-barang konsumen lainnya, kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020106','Jasa perdagangan besar dari bahan-bahan dan perangkat keras konstruksi (bangunan), kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020107','Jasa perdagangan besar dari produk kimia dan farmasi, kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020108','Jasa perdagangan besar mesin-mesin, peralatan dan suku cadang, kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020109','Jasa perdagangan besar dari produk lainnya, kecuali atas dasar balas jasa atau kontrak')," & _
"('1','020200','Jasa Perdagangan Besar Atas Dasar Balas Jasa Atau Kontrak')," & _
"('1','020201','Jasa perdagangan besar dari bahan mentah pertanian dan binatang hidup atas dasar balas jasa atau kontrak')," & _
"('1','020202','Jasa perdagangan besar dari makanan, minuman dan tembakau atas dasar balas jasa atau kontrak')," & _
"('1','020203','Jasa perdagangan besar dari tekstil, pakaian dan alas kaki atas dasar balas jasa atau kontrak')," & _
"('1','020204','Jasa perdagangan besar dari peralatan, barang-barang dan perlengkapan rumah tangga atas dasar balas jasa atau kontrak')," & _
"('1','020205','Jasa perdagangan besar dari barang-barang konsumen lainnya atas dasar balas jasa atau kontrak')," & _
"('1','020206','Jasa perdagangan besar dari bahan-bahan dan perangkat keras konstruksi (bangunan) atas dasar balas jasa atau kontrak')," & _
"('1','020207','Jasa perdagangan besar dari produk kimia dan farmasi atas dasar balas jasa atau kontrak')," & _
"('1','020208','Jasa perdagangan besar dari mesin-mesin, peralatan dan suku cadang atas dasar balas jasa atau kontrak')," & _
"('1','020209','Jasa perdagangan besar dari produk lainnya atas dasar balas jasa atau kontrak')," & _
"('1','030000','Jasa Perdagangan Eceran')," & _
"('1','030100','Jasa Perdagangan Eceran Bukan Di Toko Khusus')," & _
"('1','030101','Jasa Perdagangan Eceran Bukan Di Toko Khusus')," & _
"('1','030200','Jasa Perdagangan Eceran Di Toko Khusus')," & _
"('1','030201','Jasa perdagangan eceran di toko khusus untuk bahan mentah pertanian dan binatang hidup')," & _
"('1','030202','Jasa perdagangan eceran di toko khusus untuk makanan, minuman dan tembakau')," & _
"('1','030203','Jasa perdagangan eceran di toko khusus untuk tekstil, pakaian dan alas kaki')," & _
"('1','030204','Jasa perdagangan eceran di toko khusus untuk peralatan, barang-barang dan perlengkapan rumah tangga')," & _
"('1','030205','Jasa perdagangan eceran di toko khusus untuk barang-barang konsumen lainnya')," & _
"('1','030206','Jasa perdagangan eceran di toko khusus untuk untuk bahan-bahan dan perangkat keras konstruksi (bangunan)')," & _
"('1','030207','Jasa perdagangan eceran di toko khusus untuk produk kimia dan farmasi')," & _
"('1','030208','Jasa perdagangan eceran di toko khusus untuk mesin-mesin, peralatan dan suku cadang')," & _
"('1','030209','Jasa perdagangan eceran di toko khusus untuk produk lainnya')," & _
"('1','030300','Jasa Perdagangan Eceran Dengan Pemesanan Lewat Pos Atau Internet')," & _
"('1','030301','Jasa perdagangan eceran dengan pemesanan lewat pos atau internet untuk bahan mentah pertanian dan binatang hidup')," & _
"('1','030302','Jasa perdagangan eceran dengan pemesanan lewat pos atau internet untuk makanan, minuman dan tembakau')," & _
"('1','030303','Jasa perdagangan eceran dengan pemesanan lewat pos atau internet untuk tekstil, pakaian dan alas kaki')," & _
"('1','030304','Jasa perdagangan eceran dengan pemesanan lewat pos atau internet untuk peralatan, barang-barang dan perlengkapan rumah tangga')," & _
"('1','030305','Jasa perdagangan eceran dengan pemesanan lewat pos atau internet untuk barang_x0002_barang konsumen lainnya')," & _
"('1','030306','Jasa perdagangan eceran dengan pemesanan lewat pos atau internet bahan-bahan dan perangkat keras konstruksi (bangunan)')," & _
"('1','030307','Jasa perdagangan eceran dengan pemesanan lewat pos atau internet untuk produk kimia dan farmasi atas dasar balas jasa atau kontrak')," & _
"('1','030308','Jasa perdagangan eceran dengan pemesanan lewat pos atau internet untuk mesin_x0002_mesin, peralatan dan suku cadang')," & _
"('1','030309','Jasa perdagangan eceran produk lainnya dengan pemesanan lewat pos atau internet')," & _
"('1','030400','Jasa Perdagangan Eceran Bukan Toko Lainnya')," & _
"('1','030401','Jasa perdagangan eceran bahan mentah pertanian dan binatang hidup bukan toko lainnya')," & _
"('1','030402','Jasa perdagangan eceran makanan, minuman dan tembakau bukan toko lainnya')," & _
"('1','030403','Jasa perdagangan eceran tekstil, pakaian dan alas kaki bukan toko lainnya')," & _
"('1','030404','Jasa perdagangan eceran peralatan, barang-barang dan perlengkapan rumah tangga bukan toko lainnya')," & _
"('1','030405','Jasa perdagangan eceran barang-barang konsumen lainnya bukan toko lainnya')," & _
"('1','030406','Jasa perdagangan eceran bahan-bahan dan perangkat keras konstruksi (bangunan) bukan toko lainnya')," & _
"('1','030407','Jasa perdagangan eceran produk kimia dan farmasi bukan toko lainnya')," & _
"('1','030408','Jasa perdagangan eceran mesin-mesin, peralatan dan suku cadang bukan toko lainnya')," & _
"('1','030409','Jasa perdagangan eceran produk lainnya bukan toko lainnya')," & _
"('1','030500','Jasa Perdagangan Eceran Atas Dasar Balas Jasa Atau Kontrak')," & _
"('1','030501','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari bahan mentah pertanian dan binatang hidup')," & _
"('1','030502','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari makanan, minuman dan tembakau')," & _
"('1','030503','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari tekstil, pakaian dan alas kaki')," & _
"('1','030504','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari peralatan, barang_x0002_barang dan perlengkapan rumah tangga')," & _
"('1','030505','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari barang-barang konsumen lainnya')," & _
"('1','030506','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari bahan-bahan dan perangkat keras konstruksi (bangunan)')," & _
"('1','030507','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari produk kimia dan farmasi')," & _
"('1','030508','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari mesin-mesin, peralatan dan suku cadang')," & _
"('1','030509','Jasa perdagangan eceran atas dasar balas jasa atau kontrak dari produk lainnya')," & _
"('1','040000','Jasa Penyediaan Akomodasi, Makanan Dan Minuman')," & _
"('1','040100','Jasa Akomodasi Untuk Wisatawan')," & _
"('1','040101','Jasa akomodasi penyediaan ruangan atau kamar untuk wisatawan')," & _
"('1','040102','Jasa tempat perkemahan')," & _
"('1','040103','Jasa perkemahan dan rekreasi')," & _
"('1','040200','Jasa Akomodasi Penginapan Lainnya Untuk Wisatawan Dan Lainnya')," & _
"('1','040201','Jasa akomodasi penyediaan ruangan atau kamar untuk pelajar di asrama pelajar')," & _
"('1','040202','Jasa akomodasi penyediaan ruangan atau kamar untuk pekerja di asrama atau perkemahan untuk pekerja')," & _
"('1','040209','Jasa akomodasi penyediaan ruangan atau kamar lainnya, ytdl')," & _
"('1','040300','Jasa Penyediaan/Pelayanan Makanan')," & _
"('1','040301','Jasa pelayanan makanan dengan pelayanan penuh restoran')," & _
"('1','040302','Jasa penyiapan/penyediaan makanan dengan pelayanan terbatas')," & _
"('1','040309','Jasa katering untuk acara-acara tertentu')," & _
"('1','040400','Jasa Penyediaan Minuman')," & _
"('1','040401','Jasa penyediaan minuman')," & _
"('1','050000','Jasa Transportasi/Angkutan Penumpang')," & _
"('1','050100','Jasa Transportasi/Angkutan Penumpang Berupa Transportasi Lokal Dan Transportasi Untuk Bertamasya')," & _
"('1','050101','Jasa transportasi/angkutan darat bagi penumpang untuk kota dan daerah sekitarnya (pinggiran kota)')," & _
"('1','050102','Jasa transportasi/angkutan air lokal untuk penumpang')," & _
"('1','050103','Jasa transportasi/angkutan untuk wisata')," & _
"('1','050200','Jasa Transportasi/Angkutan Penumpang Jarak Jauh')," & _
"('1','050201','Jasa transportasi/angkutan penumpang antarkota dengan kereta api')," & _
"('1','050202','Jasa transportasi/angkutan jalan raya antarkota untuk penumpang')," & _
"('1','050203','Jasa angkutan air jarak jauh untuk penumpang')," & _
"('1','050204','Jasa angkutan udara untuk penumpang')," & _
"('1','050205','Jasa angkutan luar angkasa untuk penumpang')," & _
"('1','060000','Jasa Angkutan Barang')," & _
"('1','060100','Jasa Angkutan Barang Melalui Darat')," & _
"('1','060101','Jasa angkutan barang melalui jalan raya')," & _
"('1','060102','Jasa angkutan barang melalui kereta api')," & _
"('1','060103','Jasa angkutan melalui jaringan pipa')," & _
"('1','060200','Jasa Angkutan Air Untuk Barang')," & _
"('1','060201','Jasa angkutan perairan pantai dan lintas samudra untuk barang, dengan kapal berpendingin')," & _
"('1','060202','Jasa angkutan air di perairan dalam untuk barang')," & _
"('1','060300','Jasa Angkutan Barang Melalui Udara Dan Angkasa')," & _
"('1','060301','Jasa angkutan barang melalui udara')," & _
"('1','060302','Jasa angkutan barang melalui luar angkasa')," & _
"('1','070000','Jasa Penyewaan Kendaraan Angkutan Dengan Petugas (Operator)')"

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetUrutan4() As String
        Dim retval As String = ""
        Try
            retval = "Insert Into `tblmaster_taxproduct_reference` (`type`,`Code`,`Description`) Values" & _
            "('1','070100','Jasa Penyewaan Kendaraan Angkutan Dengan Petugas (Operator)')," & _
            "('1','070101','Jasa penyewaan kendaraan angkutan darat dengan petugas (operator)')," & _
            "('1','070102','Jasa penyewaan kapal dengan petugas (operator)')," & _
            "('1','070103','Jasa penyewaan pesawat udara/pesawat terbang dengan petugas (operator)')," & _
            "('1','080000','Jasa Pendukung Angkutan')," & _
            "('1','080100','Jasa Bongkar Muat Barang')," & _
            "('1','080101','Jasa bongkar muat kontainer')," & _
            "('1','080109','Jasa bongkar muat kargo dan bagasi lainnya')," & _
            "('1','080200','Jasa Penyimpanan Dan Pergudangan')," & _
            "('1','080201','Jasa penyimpanan dengan pendingin')," & _
            "('1','080202','Jasa penyimpanan barang berwujud gas dan cair (dalam jumlah banyak)')," & _
            "('1','080209','Jasa penyimpanan dan pergudangan lainnya.')," & _
            "('1','080300','Jasa Pendukung Untuk Angkutan Kereta Api')," & _
            "('1','080301','Jasa pendorong dan penarik angkutan kereta api')," & _
            "('1','080309','Jasa pendukung lainnya untuk angkutan kereta api')," & _
            "('1','080400','Jasa Pendukung Untuk Angkutan/Transportasi Darat')," & _
            "('1','080401','Jasa terminal bus')," & _
            "('1','080402','Jasa pengoperasian jalan raya, jembatan dan terowongan')," & _
            "('1','080403','Jasa parkir')," & _
            "('1','080404','Jasa derek untuk kendaraan pribadi dan komersil')," & _
            "('1','080409','Jasa pendukung lainnya untuk angkutan darat')," & _
            "('1','080500','Jasa Pendukung Untuk Angkutan Air')," & _
            "('1','080501','Jasa pengoperasian pelabuhan dan jalan air (tidak termasuk penanganan kargo) di daerah pantai dannperairan lintas samudra')," & _
            "('1','080502','Jasa pemandu dan penambat kapal')," & _
            "('1','080503','Jasa kapal penyelamat dan pelampung')," & _
            "('1','080509','Jasa pendukung lainnya untuk angkutan air')," & _
            "('1','080600','Jasa Pendukung Untuk Angkutan Udara Atau Angkasa')," & _
            "('1','080601','Jasa operasional bandar udara (kecuali penanganan kargo)')," & _
            "('1','080602','Jasa pengatur lalu lintas udara')," & _
            "('1','080603','Jasa pendukung lainnya untuk angkutan udara')," & _
            "('1','080604','Jasa pendukung untuk angkutan luar angkasa')," & _
            "('1','080900','Jasa Pendukung Angkutan Lainnya')," & _
            "('1','080901','Jasa agen angkutan barang dan jasa angkutan barang lainnya')," & _
            "('1','080909','Jasa pendukung angkutan lainnya')," & _
            "('1','090000','Jasa Pos Dan Kurir')," & _
            "('1','090100','Jasa Pos Dan Kurir')," & _
            "('1','090101','Jasa pos')," & _
            "('1','090102','Jasa kurir')," & _
            "('1','090103','Jasa pengiriman barang lokal')," & _
            "('1','100000','Jasa Distribusi Listrik, Gas Dan Air (Atas Usaha Sendiri)')," & _
            "('1','100100','Jasa Distribusi Listrik Dan Gas (Atas Usaha Sendiri)')," & _
            "('1','100101','Distribusi dan transmisi listrik (atas usaha sendiri)')," & _
            "('1','100102','Distribusi gas melalui saluran (atas usaha sendiri)')," & _
            "('1','100200','Jasa Distribusi Air (Atas Usaha Sendiri)')," & _
            "('1','100201','Distribusi air langsung dari sumber air, kecuali air panas dan uap (atas usaha sendiri)')," & _
            "('1','100202','Distribusi uap, air panas dan air dingin langsung dari sumber (atas usaha sendiri)')," & _
            "('1','100203','Distribusi air, selain melalui saluran (atas usaha sendiri)')," & _
            "('1','110000','Jasa Keuangan Dan Yang Berhubungan Dengannya')," & _
            "('1','110100','Jasa Keuangan Kecuali Investasi Perbankan, Jasa Asuransi Dan Jasa Pensiun')," & _
            "('1','110101','Jasa bank sentral')," & _
            "('1','110102','Jasa penyimpanan uang/deposito')," & _
            "('1','110103','Jasa pemberian pinjaman')," & _
            "('1','110104','Jasa leasing keuangan')," & _
            "('1','110109','Jasa keuangan lainnya, kecuali investasi perbankan, jasa asuransi dan jasa pensiun')," & _
            "('1','110200','Jasa Investasi Perbankan')," & _
            "('1','110201','Jasa Investasi Perbankan')," & _
            "('1','110300','Jasa Asuransi Dan Pensiun (Di Luar Jasa Reasuransi) Kecuali Jasa Jaminan Sosial Wajib')," & _
            "('1','110301','Jasa asuransi dan pensiun (di luar jasa reasuransi)')," & _
            "('1','110302','Jasa asuransi kecelakaan dan kesehatan')," & _
            "('1','110303','Jasa asuransi bukan jiwa lainnya (di luar jasa reasuransi)')," & _
            "('1','110400','Jasa Reasuransi')," & _
            "('1','110401','Jasa reasuransi jiwa')," & _
            "('1','110402','Jasa reasuransi untuk kecelakaan dan kesehatan')," & _
            "('1','110403','Jasa reasuransi bukan jiwa lainnya')," & _
            "('1','110500','Jasa Penunjang Untuk Jasa Keuangan Selain Untuk Asuransi Dan Pensiun')," & _
            "('1','110501','Jasa yang berhubungan dengan investasi perbankan')," & _
            "('1','110502','Jasa perdagangan perantara surat-surat berharga dan komoditas ybdi')," & _
            "('1','110503','Jasa manajemen portofolio kecuali dana pensiun')," & _
            "('1','110504','Jasa wali amanat (trust) dan kustudion (custody)')," & _
            "('1','110505','Jasa yang berhubungan dengan administrasi pasar uang')," & _
            "('1','110509','Jasa pendukung lainnya untuk jasa keuangan')," & _
            "('1','110600','Jasa Penunjang Untuk Asuransi Dan Pensiun')," & _
            "('1','110601','Jasa pedagang perantara dan agen asuransi')," & _
            "('1','110602','Jasa pencocokan klaim asuransi')," & _
            "('1','110603','Jasa aktuaria')," & _
            "('1','110604','Jasa manajemen dana pensiun')," & _
            "('1','110609','Jasa penunjang lainnya untuk asuransi dan pensiun')," & _
            "('1','110700','Jasa Penyimpanan Aset Keuangan')," & _
            "('1','110701','Jasa Penyimpanan Aset Keuangan')," & _
            "('1','120000','Jasa Real Estate')," & _
            "('1','120100','Jasa Real Estate Mencakup Tanah Dan/Atau Bangunan (Properti) Milik Sendiri Atau Yang Disewakan')," & _
            "('1','120101','Jasa persewaan tanah dan/atau bangunan milik sendiri atau yang disewakan')," & _
            "('1','120102','Jasa perdagangan bangunan')," & _
            "('1','120103','Jasa perdagangan tanah kosong dan tanah kavling')," & _
            "('1','120200','Jasa Real Estate Atas Dasar Balas Jasa Atau Kontrak')," & _
            "('1','120201','Jasa manajemen properti atas dasar balas jasa atau kontrak')," & _
            "('1','120202','Penjualan bangunan atas dasar balas jasa atau kontrak')," & _
            "('1','120203','Penjualan tanah atas dasar balas jasa atau kontrak')," & _
            "('1','120204','Jasa penilai real estate atas dasar balas jasa atau kontrak')," & _
            "('1','130000','Jasa Leasing Atau Persewaan Tanpa Petugas (Operator)')," & _
            "('1','130100','Jasa Leasing Atau Persewaan Mesin Dan Perlengkapannya Tanpa Petugas (Operator)')," & _
            "('1','130101','Jasa leasing atau persewaan angkutan/transportasi tanpa petugas (operator)')," & _
            "('1','130102','Jasa leasing atau persewaan mesin dan perlengkapan lainnya tanpa petugas (operator)')," & _
            "('1','130200','Jasa Leasing Atau Persewaan Barang-Barang Lainnya')," & _
            "('1','130201','Jasa leasing atau persewaan televisi, radio, perekam kaset video dan peralatan yang berhubungan dengan aksesorisnya')," & _
            "('1','130202','Jasa leasing atau persewaan pemutar dan keping video')," & _
            "('1','130203','Jasa leasing atau persewaan furnitur dan perlengkapan rumah tangga lainnya')," & _
            "('1','130204','Jasa leasing atau persewaan perlengkapan untuk bersantai dan bersenang_x0002_senang/rekreasi/tamasya')," & _
            "('1','130205','Jasa leasing atau persewaan kain linen rumah tangga')," & _
            "('1','130206','Jasa leasing atau persewaan barang-barang tekstil, pakaian dan alas kaki')," & _
            "('1','130207','Jasa leasing atau persewaan mesin-mesin otomatis dan peralatannya')," & _
            "('1','130209','Jasa leasing atau persewaan barang-barang lainnya ytdl')," & _
            "('1','130300','Jasa Lisensi Untuk Hak Penggunaan Properti Intelektual Dan Produk Sejenis')," & _
            "('1','130301','Jasa lisensi untuk hak penggunaan piranti/perangkat lunak komputer dan basis data')," & _
            "('1','130302','Jasa lisensi untuk hak penggunaan hiburan, literatur dan artistik asli')," & _
            "('1','130303','Jasa lisensi untuk hak penggunaan produk penelitian dan pengembangan')," & _
            "('1','130304','Jasa lisensi untuk hak menggunakan merk dagang dan waralaba')," & _
            "('1','130305','Jasa lisensi untuk hak menggunakan eksplorasi dan evaluasi mineral')," & _
            "('1','130309','Jasa lisensi untuk hak menggunakan produk properti intelektual lainnya')," & _
            "('1','140000','Jasa Penelitian Dan Pengembangan')," & _
            "('1','140100','Jasa Pengembangan Penelitian Dan Percobaan Dalam Ilmu Alam Dan Rekayasa')," & _
            "('1','140101','Jasa pengembangan penelitian dan percobaan dalam ilmu alam')," & _
            "('1','140102','Jasa pengembangan penelitian dan percobaan dalam ilmu rekayasa dan teknologi')," & _
            "('1','140103','Jasa pengembangan penelitian dan percobaan dalam ilmu kedokteran dan farmasi')," & _
            "('1','140104','Jasa pengembangan penelitian dan percobaan dalam ilmu pertanian')," & _
            "('1','140200','Jasa Pengembangan Penelitian Dan Percobaan Dalam Ilmu Sosial Dan Humaniora')," & _
            "('1','140201','Jasa pengembangan penelitian dan percobaan dalam ilmu sosial')," & _
            "('1','140202','Jasa pengembangan penelitian dan percobaan dalam ilmu humaniora')," & _
            "('1','140300','Jasa Pengembangan Penelitian Dan Peecobaan Antar Cabang Ilmu Pengetahuan')," & _
            "('1','140301','Jasa pengembangan penelitian dan percobaan antar cabang ilmu pengetahuan')," & _
            "('1','140400','Penelitian Dan Pengembangan Orisinil')," & _
            "('1','140401','Penelitian Dan Pengembangan Orisinil')," & _
            "('1','150000','Jasa Hukum Dan Akuntansi')," & _
            "('1','150100','Jasa Hukum')," & _
            "('1','150101','Jasa penasehat dan perwakilan hukum dalam hal hukum pidana')," & _
            "('1','150102','Jasa penasehat dan perwakilan hukum dalam hal hukum lainnya')," & _
            "('1','150103','Jasa pembuatan dokumen dan sertifikat hukum')," & _
            "('1','150109','Jasa hukum lainnya')," & _
            "('1','150200','Jasa Akuntansi, Pemeriksaan Keuangan dan Pembukuan')," & _
            "('1','150201','Jasa pemeriksaan keuangan')," & _
            "('1','150202','Jasa akuntansi dan pembukuan')," & _
            "('1','150300','Jasa Konsultan Pajak dan Persiapannya')," & _
            "('1','150301','Jasa konsultan pajak badan usaha dan persiapannya')," & _
            "('1','150302','Jasa persiapan pajak pribadi dan perencanaannya')," & _
            "('1','150400','Jasa Bantuan untuk Usaha yang Bangkrut dan dalam Pengawasan Kurator')," & _
            "('1','150401','Jasa Bantuan untuk Usaha yang Bangkrut dan dalam Pengawasan Kurator')," & _
            "('1','160000','Jasa Ahli, Teknis dan Jasa Bisnis Lainnya')," & _
            "('1','160100','Jasa Konsultasi Manajemen dan Jasa Manajemen, Jasa Teknologi Informasi')," & _
            "('1','160101','Jasa konsultasi manajemen dan jasa manajemen')," & _
            "('1','160102','Jasa konsultasi bisnis')," & _
            "('1','160103','Jasa pendukung dan konsultan teknologi informasi')," & _
            "('1','160104','Jasa desain dan pengembangan teknologi informasi')," & _
            "('1','160105','Jasa penyediaan hosting dan sarana infrastruktur teknologi informasi')," & _
            "('1','160109','Jasa manajemen lainnya kecuali manajemen proyek konstruksi')," & _
            "('1','160200','Jasa Arsitektur, Jasa Perencanaan Perkotaan dan Wilayah, dan Jasa Arsitektur Pertamanan')," & _
            "('1','160201','Jasa arsitektur dan jasa penasehat arsitektur')," & _
            "('1','160202','Jasa perencanaan wilayah dan kota')," & _
            "('1','160203','Jasa arsitektur pertamanan')," & _
            "('1','160300','Jasa Keahlian Teknik')," & _
            "('1','160301','Jasa penasehat keahlian teknik')," & _
            "('1','160302','Jasa keahlian teknik untuk proyek khusus')," & _
            "('1','160303','Jasa manajemen proyek untuk proyek konstruksi')," & _
            "('1','160400','Jasa Ilmu Pengetahuan dan Teknik Lainnya')," & _
            "('1','160401','Jasa ilmu geologi, geofisika, dan pencarian bijih lainnya')," & _
            "('1','160402','Jasa penelitian permukaan tanah dan pembuatan peta')," & _
            "('1','160403','Jasa peramalan cuaca dan meteoorologi cuaca')," & _
            "('1','160404','Jasa analisis dan uji teknis')," & _
            "('1','160500','Jasa Kedokteran Hewan')," & _
            "('1','160501','Jasa kedokteran hewan untuk hewan peliharaan')," & _
            "('1','160502','Jasa kedokteran hewan untuk hewan ternak')," & _
            "('1','160509','Jasa kedokteran hewan lainnya')," & _
            "('1','160600','Jasa Periklanan dan Jasa Penyediaan Tempat atau Waktu untuk Iklan')," & _
            "('1','160601','Jasa periklanan')," & _
            "('1','160602','Pembelian atau penjualan tempat atau waktu untuk iklan')," & _
            "('1','160603','Jasa pembelian dan penjualan tempat atau waktu iklan, kecuali dengan komisi/balas jasa')," & _
            "('1','160700','Jasa Penelitian Pasar dan Jajak Pendapat Masyarakat')," & _
            "('1','160701','Jasa Penelitian Pasar dan Jajak Pendapat Masyarakat')," & _
            "('1','160800','Jasa Fotografi dan Pencetakan Hasilnya')," & _
            "('1','160801','Jasa fotografi dan dokumentasi video')," & _
            "('1','160802','Jasa proses fotografi')," & _
            "('1','160900','Jasa Ahli, Teknik dan Bisnis Lainnya')," & _
            "('1','160901','Jasa desain khusus')," & _
            "('1','160902','Desain asli')," & _
            "('1','160903','Jasa konsultan ilmu pengetahuan dan ilmu teknik, ytdl')," & _
            "('1','160904','Kumpulan informasi asli')," & _
            "('1','160905','Jasa penerjemah dan interpretasi')," & _
            "('1','160906','Jasa merk dagang dan hak perdagangan')," & _
            "('1','160909','Semua jasa ahli, teknis dan jasa bisnis lainnya ytdl')," & _
            "('1','170000','Jasa Telekomunikasi, Penyiaran dan Penyediaan Informasi')," & _
            "('1','170100','Jasa Telepon dan Alat-Alat Telekomunikasi Lainnya')," & _
            "('1','170101','Pusat layanan')," & _
            "('1','170102','Jasa telefoni tidak bergerak (telefoni yang tidak bisa dibawa kemana-mana)')," & _
            "('1','170103','Jasa alat telekomunikasi yang bergerak')," & _
            "('1','170104','Jasa jaringan pribadi')," & _
            "('1','170105','Jasa transmisi data')," & _
            "('1','170109','Jasa telekomunikasi lainnya')," & _
            "('1','170200','Jasa Alat Telekomunikasi Internet')," & _
            "('1','170201','Jasa server internet')," & _
            "('1','170202','Jasa akses internet')," & _
            "('1','170209','Jasa internet lainnya')," & _
            "('1','170300','Jasa Penjualan Pulsa dan Kartu Perdana')," & _
            "('1','170301','Jasa Penjualan Pulsa dan Kartu Perdana')," & _
            "('1','170400','Jasa Agen Berita')," & _
            "('1','170401','Jasa agen berita berupa koran dan majalah.')," & _
            "('1','170402','Jasa agen berita untuk media audiovisual/televisi')," & _
            "('1','170500','Jasa Perpustakaan dan Tempat Penyimpanan Dokumen (Pengarsipan)')," & _
            "('1','170501','Jasa perpustakaan')," & _
            "('1','170502','Jasa tempat penyimpanan dokumen (pengarsipan)')," & _
            "('1','170600','Jasa Penyiaran, Pemrograman Dan Distribusi Program')," & _
            "('1','170601','Siaran radio dan televisi asli')," & _
            "('1','170602','Program saluran siaran radio dan televisi')," & _
            "('1','170603','Jasa penyiaran dan jasa distribusi program dengan banyak saluran')," & _
            "('1','180000','Jasa Pendukung')," & _
            "('1','180100','Jasa Ketenagakerjaan')," & _
            "('1','180101','Jasa pencarian dan penyerahan personil (ke perusahaan)')," & _
            "('1','180102','Jasa penyediaan tenaga kerja')," & _
            "('1','180200','Jasa Penyelidikan Dan Keamanan')," & _
            "('1','180201','Jasa penyelidikan')," & _
            "('1','180202','Jasa konsultan keamanan')," & _
            "('1','180203','Jasa sistem keamanan')," & _
            "('1','180204','Jasa kendaraan lapis baja')," & _
            "('1','180205','Jasa penjagaan')," & _
            "('1','180209','Jasa keamanan lainnya')," & _
            "('1','180300','Jasa Kebersihan')," & _
            "('1','180301','Jasa pembasmian hama dan pemusnahan (serangga)')," & _
            "('1','180302','Jasa kebersihan jendela')," & _
            "('1','180303','Jasa kebersihan umum')," & _
            "('1','180304','Jasa kebersihan khusus')," & _
            "('1','180400','Jasa Pengepakan')," & _
            "('1','180401','Jasa Pengepakan')," & _
            "('1','180500','Jasa Perencanaan Wisata, Petugas Wisata dan Hal-Hal Lain Yang Berhubungan')," & _
            "('1','180501','Jasa pemesanan/reservasi angkutan')," & _
            "('1','180502','Jasa pemesanan/reservasi akomodasi, pelayaran/pesiar dan paket wisata')," & _
            "('1','180503','Jasa pemesanan/reservasi lainnya')," & _
            "('1','180504','Jasa petugas wisata')," & _
            "('1','180505','Jasa pemandu wisata')," & _
            "('1','180506','Jasa promosi wisata dan informasi pengunjung')," & _
            "('1','180900','Jasa Pendukung Lainnya')," & _
            "('1','180901','Jasa laporan kredit')," & _
            "('1','180902','Jasa agen pengumpulan')," & _
            "('1','180903','Jasa pendukung dengan telepon')," & _
            "('1','180904','Jasa administrasi kantor campuran')," & _
            "('1','180905','Jasa pendukung kantor khusus')," & _
            "('1','180906','Jasa bantuan/asistensi dan pengorganisasian untuk rapat/konvensi')," & _
            "('1','180907','Jasa pemeliharaan dan perawatan pemandangan (alam)')," & _
            "('1','180909','Jasa pendukung dan informasi lainnya, ytdl')," & _
            "('1','190000','Jasa Pendukung untuk Pertanian, Perburuan, Kehutanan, Perikanan, Pertambangan dan Keperluan Lain')," & _
            "('1','190100','Jasa Pendukung untuk Pertanian, Perburuan, Kehutanan dan Perikanan')," & _
            "('1','190101','Jasa pendukung untuk produksi hasil panen')," & _
            "('1','190102','Jasa peternakan hewan')," & _
            "('1','190103','Jasa pendukung untuk perburuan')," & _
            "('1','190104','Jasa pendukung untuk kehutanan dan penebangan kayu')," & _
            "('1','190105','Jasa pendukung untuk perikanan')," & _
            "('1','190200','Jasa Pendukung untuk Pertambangan')," & _
            "('1','190201','Jasa Pendukung untuk Pertambangan')," & _
            "('1','190300','Jasa Pendukung untuk Distribusi Listrik, Gas dan Air')," & _
            "('1','190301','Jasa pendukung untuk distribusi dan transmisi listrik')," & _
            "('1','190302','Jasa distribusi gas melalui sumber gas (atas dasar biaya atau kontrak)')," & _
            "('1','190303','Jasa distribusi air melalui sumber air (atas dasar biaya atau kontrak)')," & _
            "('1','190304','Jasa distribusi persediaan uap, air panas dan pendingin ruangan (AC) melalui sumbernya (atas dasar biaya atau kontrak)')," & _
            "('1','190305','Jasa distribusi air, kecuali melalui sumber air (atas dasar biaya atau kontrak)')," & _
            "('1','200000','Jasa Perawatan, Perbaikan dan Instalasi (kecuali Konstruksi)')," & _
            "('1','200100','Jasa Perawatan dan Perbaikan Produk-Produk Logam Pabrikan, kecuali Mesin-Mesin dan Peralatan')," & _
            "('1','200101','Jasa perawatan dan perbaikan produk-produk logam pabrikan, kecuali mesin-mesin dan peralatan')," & _
            "('1','200102','Jasa perawatan dan perbaikan mesin kantor dan mesin hitung')," & _
            "('1','200103','Jasa perawatan dan perbaikan komputer dan suku cadang komputer')," & _
            "('1','200104','Jasa perawatan dan perbaikan mesin transportasi/angkutan dan peralatannya')," & _
            "('1','200105','Jasa perawatan dan perbaikan mesin-mesin dan peralatan lainnya')," & _
            "('1','200200','Jasa Perbaikan Barang-Barang Lainnya')," & _
            "('1','200201','Jasa perbaikan alas kaki dan barang-barang dari kulit')," & _
            "('1','200202','Jasa perbaikan jam tangan, jam dinding dan barang-barang perhiasan')," & _
            "('1','200203','Jasa perbaikan pakaian dan barang tekstil rumah tangga')," & _
            "('1','200204','Jasa perbaikan barang-barang mebel')," & _
            "('1','200209','Jasa perawatan dan perbaikan barang lainnya ytdl')," & _
            "('1','200300','Jasa Instalasi (selain Konstruksi)')," & _
            "('1','200301','Jasa instalasi produk-produk logam pabrikan, kecuali mesin dan peralatan')," & _
            "('1','200302','Jasa instalasi mesin dan peralatan industri, manufaktur dan jasa industri')," & _
            "('1','200303','Jasa instalasi mesin-mesin kantor dan mesin hitung dan komputer')," & _
            "('1','200304','Jasa instalasi peralatan dan perlengkapan radio, televisi, dan alat komunikasi lainnya')," & _
            "('1','200305','Jasa instalasi mesin kedokteran profesional dan peralatannya, alat-alat presisi dan optik')," & _
            "('1','200306','Jasa instalasi mesin dan perlengkapan listrik ytdl')," & _
            "('1','200309','Jasa instalasi barang-barang lain, ytdl')," & _
            "('1','210000','Jasa Pengolahan/Pembuatan dengan Input dari Industri Lain')," & _
            "('1','210100','Jasa Pengolahan Makanan, Minuman Dan Tembakau')," & _
            "('1','210101','Jasa pengolahan daging')," & _
            "('1','210102','Jasa pengolahan ikan')," & _
            "('1','210103','Jasa pengolahan buah dan sayuran')," & _
            "('1','210104','Jasa pengolahan minyak dan lemak dari tumbuhan dan hewan')," & _
            "('1','210105','Jasa pengolahan produk yang terbuat dari susu')," & _
            "('1','210106','Jasa pengolahan produk makanan lainnya')," & _
            "('1','210107','Jasa pengolahan makanan hewan')," & _
            "('1','210108','Jasa pengolahan minuman')," & _
            "('1','210109','Jasa pengolahan tembakau')," & _
            "('1','210200','Jasa Pengolahan Tekstil, Pakaian Jadi Dan Produk Kulit (Maklon)')," & _
            "('1','210201','Jasa pengolahan tekstil')," & _
            "('1','210202','Jasa pembuatan pakaian jadi')," & _
            "('1','210203','Jasa pembuatan kulit dan barang-barang yang terbuat dari kulit')," & _
            "('1','210300','Jasa Pengolahan Kayu dan Kertas')," & _
            "('1','210301','Jasa pengolahan kayu dan produk kayu')," & _
            "('1','210302','Jasa pengolahan kertas dan barang-barang yang terbuat dari kertas')," & _
            "('1','210400','Jasa Pengolahan Produk Minyak, Bahan Kimia dan Barang-Barang Farmasi')," & _
            "('1','210401','Jasa pengolahan produk dari arang dan minyak sulingan')," & _
            "('1','210402','Jasa pembuatan produk-produk dari bahan kimia')," & _
            "('1','210403','Jasa pembuatan produk dari barang-barang farmasi')," & _
            "('1','210500','Jasa Pembuatan Produk Karet Dan Plastik, Dan Produk Barang Tambang Bukan Logam')," & _
            "('1','210501','Jasa pembuatan produk karet dan plastik')," & _
            "('1','210502','Jasa pembuatan produk dari plastik')," & _
            "('1','210503','Jasa pembuatan produk barang tambang bukan logam')," & _
            "('1','210600','Jasa Pembuatan Logam Dasar')," & _
            "('1','210601','Jasa Pembuatan Logam Dasar')," & _
            "('1','210700','Jasa Pembuatan Logam Pabrikan, Mesin Dan Peralatan Lainnya')," & _
            "('1','210701','Jasa pembuatan produk logam berstruktur, tangki, tandon air dan generator uap')," & _
            "('1','210702','Jasa pembuatan senjata dan amunisi')," & _
            "('1','210703','Jasa pembuatan dan perawatan produk logam pabrikan lainnya')," & _
            "('1','210704','Jasa pembuatan komputer, barang elektronik dan barang optik')," & _
            "('1','210705','Jasa pembuatan alat-alat listrik')," & _
            "('1','210706','Jasa pembuatan mesin serbaguna')," & _
            "('1','210707','Jasa pembuatan mesin-mesin khusus')," & _
            "('1','210800','Jasa Pembuatan Alat-Alat Transportasi/Angkutan')," & _
            "('1','210801','Jasa pembuatan kendaraan bermotor dan trailer')," & _
            "('1','210802','Jasa pembuatan alat-alat transportasi/angkutan lainnya')," & _
            "('1','210900','Jasa Pembuatan Barang-Barang Lainnya')," & _
            "('1','210901','Jasa Pembuatan Barang-Barang Lainnya')," & _
            "('1','220000','Jasa Pembuatan Barang-Barang Lainnya； Jasa Penerbitan, Percetakan dan Produksi； Jasa Perbaikan Alat-Alat')," & _
            "('1','220100','Jasa Penerbitan, Percetakan Dan Produksi Ulang')," & _
            "('1','220101','Jasa penerbitan atas dasar balas jasa atau kontrak')," & _
            "('1','220102','Jasa percetakan dan produksi ulang dari media perekam, atas dasar balas jasa atau kontrak')," & _
            "('1','220200','Jasa percetakan, pressing/pemgepresan, pengecapan dan jasa pembuatan plastik sejenis')," & _
            "('1','220201','Jasa percetakan, pressing/pemgepresan, pengecapan dan jasa pembuatan plastik sejenis')," & _
            "('1','220300','Jasa Pemilihan, Penempaan, Pengecapan dan Pembuatan Barang-Barang Logam Sejenis')," & _
            "('1','220301','Jasa pemilihan besi dan baja')," & _
            "('1','220302','Jasa pemilihan logam-logam bukan besi')," & _
            "('1','220303','Jasa metalurgi seperti penempaan, pressing/pengepresan, pembentukan gulungan dan pembuatan bubuk logam')," & _
            "('1','220400','Jasa Daur Ulang Barang-Barang, atas Dasar Balas Jasa atau Kontrak')," & _
            "('1','220401','Jasa daur ulang limbah dan sisaan logam, atas dasar balas jasa atau kontrak')," & _
            "('1','220402','Jasa daur ulang limbah dan sisaan non-logam, atas dasar balas jasa atau kontrak')," & _
            "('1','230000','Jasa Administrasi Publik dan Jasa Lainnya untuk Masyarakat, Jasa Jaminan Sosial Wajib')," & _
            "('1','230100','Jasa Administrasi Pemerintahan')," & _
            "('1','230101','Jasa pemerintahan umum secara keseluruhan')," & _
            "('1','230102','Jasa administrasi publik yang menyediakan layanan pendidikan, kesehatan, kebudayaan dan jasa sosial lainnya, selain jasa jaminan sosial masyarakat')," & _
            "('1','230103','Jasa administrasi umum yang berhubungan dengan cara kerja yang lebih efisien')," & _
            "('1','230109','Jasa administrasi pemerintahan lainnya')," & _
            "('1','230200','Jasa Administrasi Umum yang Disediakan untuk Komunitas Masyarakat Umum Secara Keseluruhan')," & _
            "('1','230201','Jasa administrasi umum yang berhubungan dengan urusan administrasi luar negeri, jasa diplomat dan konsulat lar negeri')," & _
            "('1','230202','Jasa yang berhubungan dengan bantuan ekonomi luar negeri')," & _
            "('1','230203','Jasa yang berhubungan dengan bantuan militer luar negeri')," & _
            "('1','230204','Jasa pertahanan militer')," & _
            "('1','230205','Jasa pertahanan sipil')," & _
            "('1','230206','Jasa polisi dan pemadam kebakaran')," & _
            "('1','230207','Jasa administrasi umum yang berhubungan dengan pengadilan hukum')," & _
            "('1','230208','Jasa administrasi umum yang berhubungan dengan penahanan dan rehabilitasi narapidana/pelaku tindak kejahatan')," & _
            "('1','230209','Jasa administrasi umum yang berhubungan dengan urusan keselamatan dan tatanan masyarakat')," & _
            "('1','230300','Jasa Administrasi yang Berhubungan dengan Skema Jaminan Sosial Wajib')," & _
            "('1','230301','Jasa administrasi yang berhubungan dengan penyakit persalinan atau ganti rugi atas cacat tubuh')," & _
            "('1','230302','Jasa administrasi yang berhubungan dengan skema pensiun pegawai negeri； ketidakmampuan usia lanjut atau skema manfaat ahli waris selain pegawai negeri')," & _
            "('1','230303','Jasa administrasi yang berhubungan dengan pemberian pesangon (dana kompensasi) untuk pemberhentian pekerja')," & _
            "('1','230304','Jasa administrasi yang berhubungan dengan program tunjangan untuk keluarga dan anak')," & _
            "('1','240000','Jasa Pendidikan')," & _
            "('1','240100','Jasa Pendidikan Pra-Sekolah Dasar (SD)')," & _
            "('1','240101','Jasa Pendidikan Pra-Sekolah Dasar (SD)')," & _
            "('1','240200','Jasa Pendidikan Sekolah Dasar (SD)')," & _
            "('1','240201','Jasa Pendidikan Sekolah Dasar (SD)')," & _
            "('1','240300','Jasa Pendidikan Sekolah Menengah')," & _
            "('1','240301','Jasa pendidikan Sekolah Menengah Pertama (SMP), umum')," & _
            "('1','240302','Jasa pendidikan Sekolah Menengah Pertama, teknik dan kejuruan')," & _
            "('1','240303','Jasa pendidikan Sekolah Menengah Atas (SMA), umum')," & _
            "('1','240304','Jasa pendidikan Sekolah Menengah Atas (SMA), teknik dan kejuruan')," & _
            "('1','240400','Jasa Pendidikan Paska-Pendidikan Menengah Non Tersier')," & _
            "('1','240401','Jasa pendidikan paska-pendidikan menengah non tersier, umum')," & _
            "('1','240402','Jasa pendidikan paska-pendidikan menengah non tersier, teknik dan kejuruan')," & _
            "('1','240500','Jasa Pendidikan Tersier')," & _
            "('1','240501','Jasa pendidikan tersier tingkat pertama')," & _
            "('1','240502','Jasa pendidikan tersier tingkat kedua')," & _
            "('1','240900','Jasa Pendidikan Lain (Informal) Dan Pelatihan/Kursus Lainnya Serta Jasa Pendukung Pendidikan')," & _
            "('1','240901','Jasa pendidikan dan pelatihan/kursus lainnya')," & _
            "('1','240902','Jasa pendukung pendidikan')," & _
            "('1','250000','Jasa Kesehatan Masyarakat Dan Organisasi Sosial Kemanusiaan')," & _
            "('1','250100','Jasa Kesehatan Masyarakat')," & _
            "('1','250101','Jasa pasien rawat inap')," & _
            "('1','250102','Jasa kedokteran dan kedokteran gigi')," & _
            "('1','250109','Jasa kesehatan masyarakat lainnya')," & _
            "('1','250200','Jasa Institusi Rawat Inap/Panti Untuk Orang Tua (Panti Jompo) Dan Orang Cacat')," & _
            "('1','250201','Jasa kesehatan selaian rumah sakit')," & _
            "('1','250202','Jasa institusi rawat inap/panti untuk orang tua (panti jompo) dan orang yang memiliki cacat')," & _
            "('1','250300','Jasa Sosial Lainnya Dengan Akomodasi')," & _
            "('1','250301','Jasa Sosial Lainnya Dengan Akomodasi')," & _
            "('1','250400','Jasa Sosial Tanpa Akomodasi Untuk Orang Tua Dan Orang Cacat')," & _
            "('1','250401','Jasa tempat penampungan/rehabilitasi')," & _
            "('1','250409','Jasa sosial lainnya tanpa akomodasi untuk orang tua dan orang cacat')," & _
            "('1','250500','Jasa Sosial Tanpa Akomodasi Lainnya')," & _
            "('1','250501','Jasa peniitipan anak harian')," & _
            "('1','250502','Jasa pengarahan dan konseling yang berhubungan dengan anak-anak, ytdl')," & _
            "('1','250503','Jasa kesejahteraan tanpa akomodasi')," & _
            "('1','250509','Jasa sosial lainnya tanpa akkomodasi ytdl')," & _
            "('1','260000','Jasa Pengumpulan Kotoran Dan Limbah, Pengolahan Dan Pembuanagnnya Serta Jasa Perlindungan Lingkungan Lainnya')," & _
            "('1','260100','Jasa Pembuangan Kotoran, Pengolahan Kotoran Dan Pembersihan Tangki Septik')," & _
            "('1','260101','Jasa pembuangan dan pengolahan kotoran')," & _
            "('1','260102','Jasa pengosongan dan pembersihan tangki septik')," & _
            "('1','260200','Jasa Pengumpulan Limbah')," & _
            "('1','260201','Jasa pengumpulan limbah berbahaya')," & _
            "('1','260202','Jasa pengumpulan zat yang dapat didaur ulang dan tidak berbahaya')," & _
            "('1','260203','Jasa pengumpulan sampah')," & _
            "('1','260300','Jasa Pembuangan Dan Pengolahan Sampah')," & _
            "('1','260301','Jasa penyiapan, pengumpulan dan penyimpanan sampah')," & _
            "('1','260302','Jasa pengolahan dan pembuangan sampah yang mengandung zat berbahaya')," & _
            "('1','260303','Jasa pengolahan dan pembuangan sampah yang tidak berbahaya')," & _
            "('1','260400','Jasa Perbaikan')," & _
            "('1','260401','Jasa perbaikan dan pembersihan tempat/ruangan')," & _
            "('1','260402','Jasa penahanan, pengawasan dan pemantauan/monitoring dan jasa perbaikan tempat/ruangan lainnya ytdl')," & _
            "('1','260403','Jasa perbaikan gedung')," & _
            "('1','260409','Jasa perbaikan lainnya, ytdl')," & _
            "('1','260500','Jasa Sanitasi Dan Hal Sejenisnya')," & _
            "('1','260501','Jasa penyapuan dan pembersihan salju')," & _
            "('1','260509','Jasa sanitasi lainnya')," & _
            "('1','260900','Jasa Perlindungan Lingkungan Lainnya')," & _
            "('1','260901','Jasa Perlindungan Lingkungan Lainnya')," & _
            "('1','270000','Jasa Keanggotaan Organisasi')," & _
            "('1','270100','Jasa Yang Disediakan Oleh Perusahaan, Organisasi Pekerja Dan Organisasi Profesional')," & _
            "('1','270101','Jasa yang disediakan oleh perusahaan dan organisasi pekerja')," & _
            "('1','270102','Jasa yang disediakan oleh organisasi profesional')," & _
            "('1','270200','Jasa Yang Disediakan Oleh Serikat Perdagangan')," & _
            "('1','270201','Jasa Yang Disediakan Oleh Serikat Perdagangan')," & _
            "('1','270900','Jasa Yang Disediakan Oleh Keanggotaan Organisasi Lainnya')," & _
            "('1','270901','Jasa keagamaan')," & _
            "('1','270902','Jasa yang disediakan oleh organisasi politik')," & _
            "('1','270909','Jasa yang disediakan keanggotaan organisasi lainnya, ytdl')," & _
            "('1','280000','Jasa Pariwisata, Kebudayaan Dan Olahraga')," & _
            "('1','280100','Jasa Audiovisual')," & _
            "('1','280101','Jasa perekaman suara')," & _
            "('1','280102','Jasa produksi gambar bergerak, video tape, program televisi dan radio')," & _
            "('1','280103','Jasa paska produksi audiovisual')," & _
            "('1','280104','Jasa distribusi gambar bergerak, video tape dan program televisi')," & _
            "('1','280105','Jasa proyeksi gambar bergerak')," & _
            "('1','280200','Jasa Presentasi Dan Promosi Acara Pertunjukan Kesenian Dan Hiburan Langsung Lainnya')," & _
            "('1','280201','Jasa promosi dan pengorganisasian acara pertunjukan kesenian')," & _
            "('1','280202','Jasa produksi dan presentasi acara pertunjukan kesenian')," & _
            "('1','280203','Jasa fasilitas untuk operasi acara kesenian')," & _
            "('1','280209','Jasa kesenian dan pertunjukan hiburan lainnya')," & _
            "('1','280300','Jasa Seniman Pertunjukan Dan Seniman Lainnya')," & _
            "('1','280301','Jasa seniman pertunjukan')," & _
            "('1','280302','Jasa penulis, pengubah, pemahat dan seniman lain kecuali seniman pertunjukan')," & _
            "('1','280303','Karya asli dari penulis, pengubah, dan seniman lain kecuali seniman pertunjukan, pelukis dan pemahat')," & _
            "('1','280400','Jasa Museum Dan Pemeliharaannya')," & _
            "('1','280401','Jasa museum dan pemeliharaan tempat serta bangunan bersejarah')," & _
            "('1','280402','Jasa kebun raya, kebun binatang dan cagar alam')," & _
            "('1','280500','Jasa Olahraga Dan Olahraga Rekreasi')," & _
            "('1','280501','Jasa promosi dan pengorganisasian kegiatan olah raga dan olah raga rekreasi')," & _
            "('1','280502','Jasa pengoperasian fasilitas olah raga dan olah raga rekreasi')," & _
            "('1','280509','Jasa olah raga dan olah raga rekreasi lainnya')," & _
            "('1','280600','Jasa Atlet')," & _
            "('1','280601','Jasa atlet')," & _
            "('1','280602','Jasa pendukung yang berhubungan dengan olah raga dan rekreasi')," & _
            "('1','280900','Jasa Hiburan Dan Rekreasi Lainnya')," & _
            "('1','280901','Jasa taman hiburan dan atraksi sejenisnya')," & _
            "('1','280902','Jasa perjudian dan pertaruhan')," & _
            "('1','280903','Jasa mesin hiburan yang menggunakan koin/uang logam')," & _
            "('1','280909','Jasa rekreasi dan hiburan lainnya ytdl')," & _
            "('1','290000','Jasa Lainnya')," & _
            "('1','290100','Jasa Pencucian, Pembersihan Dan Pewarnaan')," & _
            "('1','290101','Jasa loundry/mesin cuci dengan koin')," & _
            "('1','290102','Jasa dry-cleaning/pencucian kering (termasuk jasa membersihkan produk kulit berbulu)')," & _
            "('1','290103','Jasa pembersihan barang tekstil lainnya')," & _
            "('1','290104','Jasa pengepresan')," & _
            "('1','290105','Jasa pencelupan dan pewarnaan')," & _
            "('1','290200','Jasa Kecantikan Dan Perawatan Tubuh')," & _
            "('1','290201','Jasa penataan dan pemotongan rambut')," & _
            "('1','290202','Jasa perawatan kosmetik, perawatan tangan dan kuku/manikur dan perawatan kaki/pedikur')," & _
            "('1','290203','Jasa perawatan tubuh')," & _
            "('1','290209','Jasa perawatan kecantikan lainnya')," & _
            "('1','290300','Jasa Penguburan, Pembakaran Mayat Dan Pengurusan Pemakaman')," & _
            "('1','290301','Jasa kuburan dan pembakaran mayat')," & _
            "('1','290302','Jasa pengurusan pemakaman')," & _
            "('1','290900','Aneka Jasa Lainnya')," & _
            "('1','290901','Jasa pendamping (escort)')," & _
            "('1','290909','Aneka jasa lainny, ytdl')," & _
            "('1','300000','Jasa Pembantu Rumah Tangga')," & _
            "('1','300100','Jasa Pembantu Rumah Tangga')," & _
            "('1','300101','Jasa Pembantu Rumah Tangga')," & _
            "('1','310000','Jasa Yang Disediakan Oleh Badan Dan Organisasi Di Luar Wilayah Dalam Negeri (Luar Negeri)')," & _
            "('1','310100','Jasa Yang Disediakan Oleh Badan Dan Organisasi Di Luar Wilayah Dalam Negeri (Luar Negeri)')," & _
            "('1','310101','Jasa Yang Disediakan Oleh Badan Dan Organisasi Di Luar Wilayah Dalam Negeri (Luar Negeri)');"

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetInsertString(ByVal vUrutan As Integer) As String
        Dim retval As String = ""
        Try
            If vUrutan = 1 Then
                retval = Me.GetUrutan1()
            ElseIf vUrutan = 2 Then
                retval = Me.GetUrutan2()
            ElseIf vUrutan = 3 Then
                retval = Me.GetUrutan3()
            ElseIf vUrutan = 4 Then
                retval = Me.GetUrutan4()
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

End Class
