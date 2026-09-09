// ============================================
// Fiyat Takip - Çeviri Sözlüğü (TR / EN)
// ============================================
const TRANSLATIONS = {
  tr: {
    // --- Ortak ---
    common: {
      save: 'Kaydet', cancel: 'İptal', delete: 'Sil', edit: 'Düzenle',
      refresh: 'Yenile', clear: 'Temizle', search: 'Ara...', loading: 'Yükleniyor...',
      yes: 'Evet', no: 'Hayır', close: 'Kapat', back: 'Geri',
      total: 'Toplam', export: 'Dışa Aktar', import: 'İçe Aktar',
      confirm: 'Onayla', add: 'Ekle', remove: 'Kaldır', update: 'Güncelle',
      filter: 'Filtrele', all: 'Tümü', none: 'Hiçbiri', actions: 'İşlemler',
      upload: 'Yükle', download: 'İndir', select: 'Seçiniz',
      image: 'Görsel', date: 'Tarih', amount: 'Tutar', quantity: 'Adet',
      price: 'Fiyat', note: 'Not', status: 'Durum', customer: 'Müşteri',
      company: 'Firma', model: 'Model', color: 'Renk', currency: 'Para Birimi'
    },

    // --- Giriş ---
    login: {
      title: 'Fiyat Takip Sistemi',
      username: 'Kullanıcı Adı',
      password: 'Parola',
      loginBtn: 'Giriş Yap',
      errorCredentials: 'Kullanıcı adı veya parola hatalı.',
      errorEmpty: 'Kullanıcı adı ve parola gerekli.',
      errorConnection: 'Bağlantı hatası. Lütfen tekrar deneyin.'
    },

    // --- Üst Bar ---
    topbar: {
      backup: 'Yedekle',
      backupTitle: 'Tüm verileri yedekle',
      logout: 'Çıkış',
      logoutTitle: 'Çıkış Yap',
      userManagement: 'Kullanıcı Yönetimi',
      langTitle: 'Dil Değiştir / Switch Language',
      themeMavi: 'Mavi Tema',
      themeYesil: 'Yeşil Tema',
      themeMor: 'Mor Tema',
      themeKoyu: 'Koyu Tema'
    },

    // --- Sekmeler ---
    tabs: {
      fiyatlar: 'Fiyatlar',
      siparisler: 'Siparişler',
      gerceklesen: 'Gerçekleşen',
      odemeler: 'Ödemeler',
      gelirTakip: 'Gelir Takip',
      cari: 'Cari'
    },

    // --- Fiyatlar ---
    fiyatlar: {
      title: 'Fiyat Listesi',
      addTitle: 'Yeni Fiyat Ekle',
      editTitle: 'Fiyat Düzenle',
      emptyMsg: 'Henüz kayıtlı fiyat yok.',
      imageRemove: 'Görseli Kaldır',
      searchPlaceholder: 'Müşteri, firma veya ürün ara...',
      fieldMusteri: 'Müşteri Adı',
      fieldFirma: 'Firma Adı',
      fieldUrun: 'Ürün Açıklaması',
      fieldFiyat: 'Fiyat',
      fieldMiktar: 'Miktar',
      fieldTarih: 'Tarih',
      fieldNotlar: 'Notlar',
      fieldGorsel: 'Görsel',
      fieldExcel: 'Maliyet Excel\'i',
      placeholderMusteri: 'Örn: Ahmet Yılmaz',
      placeholderFirma: 'Seçin veya yazın...',
      placeholderFiyat: '0.00',
      placeholderNotlar: 'Opsiyonel notlar...',
      excelInfo: 'Boş bırakırsan mevcut Excel korunur.',
      existingImageRemove: 'Mevcut Görseli Kaldır',
      newImageRemove: 'Yeni Görseli Kaldır',
      toastSaved: 'Fiyat kaydedildi!',
      toastUpdated: 'Değişiklikler kaydedildi!',
      toastDeleted: 'Kayıt silindi',
      confirmDelete: 'Bu kaydı silmek istediğine emin misin?',
      confirmSiparisAktar: 'Bu fiyat kaydını siparişe aktarmak istediğine emin misin? Kayıt fiyat listesinden kaldırılıp siparişler sayfasına taşınacak.',
      toastAktarildi: 'Siparişe aktarıldı!',
      errorLoad: 'Kayıt yüklenemedi',
      colImage: 'Görsel', colDate: 'Tarih', colCustomer: 'Müşteri', colCompany: 'Firma',
      colModel: 'Model', colPrice: 'Fiyat', colQuantity: 'Adet', colTotal: 'Toplam',
      colNote: 'Not',
      downloadExcel: 'Maliyet Excel\'i İndir',
      transferSiparis: 'Siparişe Aktar',
      noImage: 'Görsel yok'
    },

    // --- Siparişler ---
    siparisler: {
      title: 'Sipariş Listesi',
      emptyMsg: 'Henüz sipariş yok.',
      completedTitle: 'Tamamlanan Sipariş Listesi',
      editTitle: 'Sipariş Düzenle',
      bantView: 'Bant Görünümü',
      returnToList: 'Listeye Dön',
      importBant: 'Bant Excel\'i Yükle',
      exportBant: 'Bant Planını İndir',
      confirmDelete: 'Bu siparişi silmek istediğine emin misin?',
      confirmGerceklesenAktar: 'Bu siparişi gerçekleşen bölümüne aktarmak istediğine emin misin? Sipariş buradan kaldırılıp Gerçekleşen sayfasına taşınacak.',
      toastSaved: 'Sipariş güncellendi!',
      toastDeleted: 'Sipariş silindi',
      toastStatus: 'Durum güncellendi',
      toastQuickUpdate: 'Güncellendi',
      errorLoad: 'Sipariş yüklenemedi',
      colTermin: 'Termin', colFit: 'FİT', colLdTest: 'LD TEST', colBulkTest: 'BULK TEST',
      colGoldTest: 'GOLD TEST', colAksesuar: 'AKSESUAR', colKumas: 'KUMAŞ',
      aksesuarTitle: 'Aksesuar Detayları', aksesuarTamam: 'Aksesuar Tamam',
      kumasTitle: 'Kumaş Detayları', kumasTamam: 'Kumaş Tamam',
      markComplete: 'Tamamlandı olarak işaretle', toastCompleted: 'Tamamlandı olarak işaretlendi',
      siparisReverted: 'Sipariş geri alındı',
      summaryActive: 'Aktif Sipariş', summaryQuantity: 'Toplam Adet',
      summaryTotal: 'Toplam Tutar', summaryHazirlaniyor: 'Hazırlanıyor',
      summaryUretimde: 'Üretimde', summaryTamamlandi: 'Tamamlandı',
      statusHazirlaniyor: 'Hazırlanıyor', statusUretimde: 'Üretimde',
      statusTamamlandi: 'Tamamlandı', statusIptal: 'İptal',
      bantUploadEmpty: 'Bant Excel\'i yüklenmemiş. Lütfen önce Yükle butonuyla bant planını yükleyin.',
      orderSelectTitle: 'Sipariş Seç',
      searchBantPlaceholder: 'Model ara...',
      searchBantOrderPlaceholder: 'Model veya müşteri ara...',
      addRow: 'Satır Ekle', deleteRow: 'Satır Sil', deleteSelection: 'Seçimi Sil',
      rowHint: 'Sadece seçili sütunda çalışır',
      toastRowAdded: 'satır eklendi', toastRowDeleted: 'satır silindi',
      bantSearchResult: 'eşleşme', bantSearchNotFound: 'Sonuç bulunamadı',
      completedTotal: 'Toplam'
    },

    // --- Gerçekleşen ---
    gerceklesen: {
      title: 'Gerçekleşen Siparişler & Maliyet',
      emptyMsg: 'Henüz gerçekleşen kaydı yok.',
      completedTitle: 'Tamamlanan Gerçekleşen',
      editTitle: 'Gerçekleşen Düzenle',
      confirmDelete: 'Bu kaydı silmek istediğine emin misin?',
      confirmDeleteSiparis: 'Bu siparişi silmek istediğine emin misin?',
      toastSaved: 'Gerçekleşen güncellendi!',
      toastDeleted: 'Gerçekleşen kaydı silindi',
      toastSiparisDeleted: 'Sipariş silindi',
      toastTransferred: 'Gerçekleşene aktarıldı!',
      toastStatus: 'Durum güncellendi',
      errorLoad: 'Kayıt yüklenemedi',
      colImage: 'Görsel', colCustomer: 'Müşteri', colCompany: 'Firma',
      colModel: 'Model', colTermin: 'Termin', colPrice: 'Fiyat', colQuantity: 'Adet',
      colKesim: 'Kesim', colYuklenen: 'Yüklenen', colIkinciKalite: '2.Kalite',
      colSatisToplam: 'Satış Top.', colKumas: 'Kumaş', colAksesuar: 'Aksesuar',
      colIscilik: 'İşçilik', colBaski: 'Baskı/Nakış/Yıkama', colTopMaliyet: 'Top.Maliyet',
      colKarZarar: 'Kâr/Zarar', colKarYuzde: 'Kâr %',
      labelKumasBedeli: 'Kumaş Bedeli',
      labelAksesuarBedeli: 'Aksesuar Bedeli',
      labelIscilikBedeli: 'İşçilik Bedeli',
      labelBaskiNakisYikama: 'Baskı/Nakış/Yıkama',
      labelDigerGider: 'Diğer Giderler',
      addDetail: 'Ekle',
      kumasFaturaTitle: 'Kumaş Faturaları',
      aksesuarFaturaTitle: 'Aksesuar Faturaları',
      excelInfo: 'Boş bırakırsan mevcut Excel korunur. Yeni Excel seçersen güncellenir.',
      openExcel: 'Maliyet Excel\'ini Aç',
      summaryCompleted: 'Tamamlanan', summaryTotalSales: 'Toplam Satış',
      summaryTotalCost: 'Toplam Maliyet', summaryNetProfit: 'Net Kâr / Zarar',
      statusDevamEdiyor: 'Devam Ediyor',
      monthlyReport: 'Aylık Kâr / Zarar Raporu',
      monthCol: 'Ay', salesTotalCol: 'Satış Top.', costTotalCol: 'Maliyet Top.',
      profitLossCol: 'Kâr / Zarar', profitPctCol: 'Kâr %', recordCol: 'Kayıt',
      grandTotal: 'GENEL TOPLAM',
      yearlyReport: 'Yıllık Kâr / Zarar Raporu', yearCol: 'Yıl',
      detailAciklama: 'Açıklama', detailDeger: 'Değer'
    },

    // --- Ödemeler ---
    odemeler: {
      title: 'Ödeme Listesi',
      emptyMsg: 'Henüz ödeme kaydı yok.',
      collectedTitle: 'Tahsil Edilenler',
      weeklyTitle: 'Haftalık Ödeme Planı',
      newOdeme: 'Yeni Ödeme',
      editOdeme: 'Ödeme Düzenle',
      emptyWeekly: 'Henüz bekleyen ödeme yok.',
      emptyCollected: 'Henüz tahsil edilen ödeme yok.',
      confirmDelete: 'Bu ödeme kaydını silmek istediğine emin misin?',
      toastSaved: 'Ödeme kaydedildi!',
      toastUpdated: 'Ödeme güncellendi!',
      toastStatus: 'Durum güncellendi',
      toastInstallmentSaved: 'taksit kaydedildi!',
      toastInstallmentUpdated: 'taksit güncellendi',
      toastDeleted: 'Silindi',
      errorLoad: 'Ödeme listesi yüklenemedi',
      errorWeeklyLoad: 'Haftalık görünüm yüklenemedi',
      errorValidation: 'Tutar ve vade tarihi zorunludur.',
      errorVade: 'Vade en az 1 ay olmalıdır.',
      errorNoInstallment: 'Güncellenecek taksit bulunamadı.',
      colFirma: 'Firma / Cari', colTur: 'Tür', colTutar: 'Tutar',
      colVade: 'Vade', colCekBanka: 'Çek/Banka', colStatus: 'Durum',
      colTahsilTarihi: 'Tahsil Tarihi',
      filterAllStatus: 'Tüm Durumlar', filterAllTypes: 'Tüm Türler',
      searchPlaceholder: 'Firma, çek no veya banka ara...',
      statusBekliyor: 'Bekliyor', statusTahsilEdildi: 'Tahsil Edildi',
      statusGecikti: 'Gecikti', statusIptalO: 'İptal',
      typeNakit: 'Nakit', typeCek: 'Çek',
      typeKrediKarti: 'Kredi Kartı',
      typeBankaKredisi: 'Banka Kredisi', typeSpotKredi: 'Spot Kredi', typeRotatifKredi: 'Rotatif Kredi', typeVadeliOdeme: 'Vadeli Ödeme',
      typeKira: 'Kira', typeMaas: 'Maaş', typeVergi: 'Vergi',
      fieldFirma: 'Firma / Cari', fieldOdemeTuru: 'Ödeme Türü', fieldTutar: 'Tutar',
      fieldVade: 'Vade Tarihi', fieldAciklama: 'Açıklama',
      fieldCekNo: 'Çek No', fieldBanka: 'Banka Adı', fieldCekSahibi: 'Çeki Veren',
      fieldKrediSure: 'Kaç ay?', fieldKrediTutar: 'Aylık Taksit Tutarı',
      krediHint: 'İlk ödeme tarihinden itibaren her ay otomatik ödeme kaydı oluşturulur.',
      krediUpdateHint: 'Kalan tutar, ödenmemiş taksitlere eşit bölünecek.',
      installmentLabel: 'Taksit', midPayment: 'Ara Ödeme',
      updateRemaining: 'Kalanı Güncelle',
      creditGroupPaid: 'Tamamı Ödendi', creditGroupRemaining: 'Kalan',
      creditGroupPaidAmount: 'Ödenen',
      weekOverdue: 'Gecikmiş', weekThisWeek: 'Bu Hafta', weekThisMonth: 'Bu Ay', weekPast: 'Geçti',
      weekPrefix: 'Hafta',
      summaryBekleyen: 'Bekleyen', summaryThisWeek: 'Bu Hafta',
      summaryGecikmis: 'Gecikmiş', summaryToplamBekleyen: 'Toplam Bekleyen',
      summaryTahsilEdilen: 'Tahsil Edilen', summaryTotalTahsilat: 'Toplam Tahsilat',
      summaryBuAyFaiz: 'Bu Ay Faiz',
      combinedTitle: 'Gelir-Gider Raporu',
      weekGelir: 'Gelir',
      weekGider: 'Gider',
      weekBakiye: 'Net',
      weekGelirler: 'Gelirler',
      weekGiderler: 'Giderler',
      emptyCombined: 'Henüz gelir veya gider kaydı yok.'
    },

    // --- Aksesuar Takip ---
    aksesuar: {
      title: 'Aksesuar Takip',
      emptyMsg: 'Henüz aksesuar eklenmedi.',
      newAksesuar: 'Yeni Aksesuar Ekle',
      colName: 'Aksesuar', colSiparisAdet: 'Sip.Miktar',
      colGelenAdet: 'Gelen', colBirimFiyat: 'Birim Fiyat', colNot: 'Not',
      confirmDelete: 'Bu aksesuarı silmek istediğine emin misin?',
      toastDeleted: 'Aksesuar silindi',
      errorAdd: 'Eklenemedi', errorDelete: 'Silinemedi',
      sevkiyatTitle: 'Açıklama', sevkiyatAdet: 'Miktar',
      sevkiyatFatura: 'Fatura No', sevkiyatTutar: 'Tutar',
      addSevkiyat: 'Sevkiyat Ekle'
    },

    // --- Kumaş Takip ---
    kumas: {
      title: 'Kumaş Takip',
      emptyMsg: 'Henüz kumaş eklenmedi.',
      newKumas: 'Yeni Kumaş Ekle',
      confirmDelete: 'Bu kumaşı silmek istediğine emin misin?',
      toastDeleted: 'Kumaş silindi',
      toastCalc: 'Birim',
      errorAdd: 'Eklenemedi', errorDelete: 'Silinemedi', errorLoad: 'Kumaş listesi yüklenemedi',
      colName: 'Kumaş Adı', colColor: 'Renk', colAdet: 'Adet',
      colPastalBoyu: 'Pastal B. (m)', colEn: 'En (m)', colGramaj: 'Gramaj (g/m²)',
      colFire: 'Fire %', colAsorti: 'Asorti', colBiye: 'Biye (g)',
      colBirimGider: 'Birim Gider (GR/MT)', colKumasSiparis: 'Kumaş Siparişi',
      colToplamGelen: 'Toplam Gelen', colKumasFiyat: 'Kumaş Fiyatı',
      colKesimAdet: 'Kesim Ad.', colKesimYuzde: 'Kesim %',
      colGerceklesenBirim: 'Gerç. Birim',
      sevkiyatMiktar: 'Miktar', sevkiyatFatura: 'Fatura No',
      sevkiyatTutar: 'Tutar', addSevkiyat: 'Sevkiyat Ekle'
    },

    // --- Kullanıcı Yönetimi ---
    kullanicilar: {
      title: 'Kullanıcı Yönetimi',
      newUser: 'Yeni Kullanıcı',
      colUser: 'Kullanıcı', colRol: 'Rol', colDate: 'Tarih',
      fieldUsername: 'Kullanıcı Adı', fieldPassword: 'Parola',
      fieldPasswordLeave: 'Boş bırakılırsa değişmez',
      fieldRole: 'Rol', fieldPermissions: 'Yetkiler',
      errorUsername: 'Kullanıcı adı zorunludur.',
      errorPermission: 'En az bir yetki seçilmelidir.',
      toastSaved: 'Kullanıcı eklendi', toastUpdated: 'Kullanıcı güncellendi',
      toastDeleted: 'Kullanıcı silindi',
      confirmDeletePrefix: 'kullanıcısını silmek istediğine emin misin?',
      permLabel: 'yetki',
      roles: { admin: 'Admin', muhasebe: 'Muhasebe', operasyon: 'Operasyon', satis: 'Satış' }
    },

    // --- Bant Görünümü ---
    bant: {
      title: 'Bant Programı',
      uploadBtn: 'Yükle', downloadBtn: 'İndir',
      searchPlaceholder: 'Model ara...',
      emptyMsg: 'Bant Excel\'i yüklenmemiş.',
      toastUploadOk: 'Yüklendi!',
      toastUploadError: 'Hata'
    },

    // --- Gelir Takip ---
    gelirTakip: {
      title: 'Gelir Değerlendirme',
      emptyMsg: 'Henüz gelir takip kaydı yok.',
      editTitle: 'Gelir Değerlendirme',
      confirmDelete: 'Bu gelir takip kaydını silmek istediğinize emin misiniz?',
      toastSaved: 'Gelir takip güncellendi!',
      toastDeleted: 'Gelir takip silindi',
      errorLoad: 'Gelir takip listesi yüklenemedi',
      colCustomer: 'Müşteri', colCompany: 'Firma', colModel: 'Model',
      colToplamTutar: 'Toplam Tutar', colNakit: 'Nakit', colCek: 'Çek',
      colKalan: 'Kalan', colNot: 'Not',
      summaryToplamGelir: 'Toplam Gelir', summaryNakit: 'Nakit',
      summaryCekToplam: 'Çek Toplam', summaryTeminatCek: 'Teminata Verilen Çek',
      summaryKirdirilanCek: 'Kırdırılan Çek', summaryNetNakit: 'Net Nakit',
      fieldNakitTutar: 'Nakit Tutar',
      cekEkle: '+ Çek Ekle', cekSil: 'Çek Sil',
      cekTutar: 'Çek Tutarı', cekVade: 'Vade Tarihi', cekBanka: 'Banka',
      cekNo: 'Çek No',
      cekDurumBekliyor: 'Bekliyor', cekDurumTeminata: 'Teminata Verildi',
      cekDurumKirdirildi: 'Kırdırıldı',
      cekKomisyon: 'Komisyon / Kesinti', cekNetTutar: 'Net Tutar',
      cekKirdirmaTarihi: 'Kırdırma Tarihi', cekNotlar: 'Notlar',
      cekBaslik: 'Çek Detayları',
      noCek: 'Çek bulunmuyor'
    },

    // --- Cari (Tedarikçi Hesapları) ---
    cari: {
      title: 'Cari (Tedarikçi) Hesapları',
      emptyMsg: 'Henüz cari kartı yok.',
      newCari: '+ Yeni Cari',
      editCari: 'Cari Kartı Düzenle',
      confirmDelete: 'Bu cari kartını ve tüm hareketlerini silmek istediğinize emin misiniz?',
      colKod: 'Cari Kodu',
      colAd: 'Cari Adı',
      colVergiNo: 'Vergi No',
      colTelefon: 'Telefon',
      colBorc: 'Borç',
      colAlacak: 'Alacak',
      colBakiye: 'Bakiye',
      ekstre: 'Cari Ekstre',
      hareketler: 'Hareketler',
      yeniHareket: '+ Hareket Ekle',
      hareketDuzenle: 'Hareket Düzenle',
      turBorc: 'Borç (Alış Faturası)',
      turAlacak: 'Alacak (Ödeme)',
      fieldTarih: 'Tarih',
      fieldTur: 'Tür',
      fieldTutar: 'Tutar',
      fieldParaBirimi: 'Para Birimi',
      fieldAciklama: 'Açıklama',
      fieldBelgeNo: 'Belge No',
      fieldOdemeSekli: 'Ödeme Şekli',
      odemeNakit: 'Nakit',
      odemeCek: 'Çek',
      odemeHavale: 'Havale/EFT',
      toplamBorc: 'Toplam Borç',
      toplamAlacak: 'Toplam Alacak',
      kapanisBakiyesi: 'Kapanış Bakiyesi',
      noHareket: 'Bu cariye ait hareket yok.',
      confirmDeleteHareket: 'Bu hareketi silmek istediğinize emin misiniz?',
      toastSaved: 'Cari kartı kaydedildi!',
      toastDeleted: 'Cari kartı silindi',
      toastHareketSaved: 'Hareket kaydedildi!',
      toastHareketDeleted: 'Hareket silindi',
      errorLoad: 'Cari listesi yüklenemedi'
    },

    // --- Aylar ---
    months: ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
             'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık'],
    monthsShort: ['Oca','Şub','Mar','Nis','May','Haz',
                  'Tem','Ağu','Eyl','Eki','Kas','Ara'],

    // --- Renkler ---
    colors: {
      red: 'Kırmızı', blue: 'Lacivert', black: 'Siyah', white: 'Beyaz',
      grey: 'Gri', green: 'Yeşil', yellow: 'Sarı', purple: 'Mor',
      placeholder: 'Örn: Kırmızı, Lacivert', name: 'Renk adı', qty: 'Adet'
    }
  },

  // ============================================
  // ENGLISH
  // ============================================
  en: {
    common: {
      save: 'Save', cancel: 'Cancel', delete: 'Delete', edit: 'Edit',
      refresh: 'Refresh', clear: 'Clear', search: 'Search...', loading: 'Loading...',
      yes: 'Yes', no: 'No', close: 'Close', back: 'Back',
      total: 'Total', export: 'Export', import: 'Import',
      confirm: 'Confirm', add: 'Add', remove: 'Remove', update: 'Update',
      filter: 'Filter', all: 'All', none: 'None', actions: 'Actions',
      upload: 'Upload', download: 'Download', select: 'Select',
      image: 'Image', date: 'Date', amount: 'Amount', quantity: 'Qty',
      price: 'Price', note: 'Note', status: 'Status', customer: 'Customer',
      company: 'Company', model: 'Model', color: 'Color', currency: 'Currency'
    },

    login: {
      title: 'Price Tracking System',
      username: 'Username',
      password: 'Password',
      loginBtn: 'Sign In',
      errorCredentials: 'Invalid username or password.',
      errorEmpty: 'Username and password are required.',
      errorConnection: 'Connection error. Please try again.'
    },

    topbar: {
      backup: 'Backup',
      backupTitle: 'Backup all data',
      logout: 'Logout',
      logoutTitle: 'Sign Out',
      userManagement: 'User Management',
      langTitle: 'Switch Language / Dil Değiştir',
      themeMavi: 'Blue Theme',
      themeYesil: 'Green Theme',
      themeMor: 'Purple Theme',
      themeKoyu: 'Dark Theme'
    },

    tabs: {
      fiyatlar: 'Prices',
      siparisler: 'Orders',
      gerceklesen: 'Actual Costs',
      odemeler: 'Payments',
      gelirTakip: 'Revenue',
      cari: 'Accounts'
    },

    fiyatlar: {
      title: 'Price List',
      addTitle: 'Add New Price',
      editTitle: 'Edit Price',
      emptyMsg: 'No prices recorded yet.',
      imageRemove: 'Remove Image',
      searchPlaceholder: 'Search customer, company or product...',
      fieldMusteri: 'Customer Name',
      fieldFirma: 'Company Name',
      fieldUrun: 'Product Description',
      fieldFiyat: 'Price',
      fieldMiktar: 'Quantity',
      fieldTarih: 'Date',
      fieldNotlar: 'Notes',
      fieldGorsel: 'Image',
      fieldExcel: 'Cost Excel',
      placeholderMusteri: 'e.g. John Doe',
      placeholderFirma: 'Select or type...',
      placeholderFiyat: '0.00',
      placeholderNotlar: 'Optional notes...',
      excelInfo: 'Leave empty to keep current Excel.',
      existingImageRemove: 'Remove Current Image',
      newImageRemove: 'Remove New Image',
      toastSaved: 'Price saved!',
      toastUpdated: 'Changes saved!',
      toastDeleted: 'Record deleted',
      confirmDelete: 'Are you sure you want to delete this record?',
      confirmSiparisAktar: 'Are you sure you want to transfer this price to orders? It will be moved to the Orders page.',
      toastAktarildi: 'Transferred to orders!',
      errorLoad: 'Could not load record',
      colImage: 'Image', colDate: 'Date', colCustomer: 'Customer', colCompany: 'Company',
      colModel: 'Model', colPrice: 'Price', colQuantity: 'Qty', colTotal: 'Total',
      colNote: 'Note',
      downloadExcel: 'Download Cost Excel',
      transferSiparis: 'Transfer to Order',
      noImage: 'No image'
    },

    siparisler: {
      title: 'Order List',
      emptyMsg: 'No orders yet.',
      completedTitle: 'Completed Orders',
      editTitle: 'Edit Order',
      bantView: 'Ribbon View',
      returnToList: 'Return to List',
      importBant: 'Upload Ribbon Excel',
      exportBant: 'Download Ribbon Plan',
      confirmDelete: 'Are you sure you want to delete this order?',
      confirmGerceklesenAktar: 'Are you sure you want to transfer this order to actual costs? It will be moved from here.',
      toastSaved: 'Order updated!',
      toastDeleted: 'Order deleted',
      toastStatus: 'Status updated',
      toastQuickUpdate: 'Updated',
      errorLoad: 'Could not load order',
      colTermin: 'Deadline', colFit: 'FIT', colLdTest: 'LD TEST', colBulkTest: 'BULK TEST',
      colGoldTest: 'GOLD TEST', colAksesuar: 'ACCESSORIES', colKumas: 'FABRIC',
      aksesuarTitle: 'Accessory Details', aksesuarTamam: 'Accessories Done',
      kumasTitle: 'Fabric Details', kumasTamam: 'Fabric Done',
      markComplete: 'Mark as Completed', toastCompleted: 'Marked as completed',
      siparisReverted: 'Order reverted',
      summaryActive: 'Active Orders', summaryQuantity: 'Total Qty',
      summaryTotal: 'Total Amount', summaryHazirlaniyor: 'Preparing',
      summaryUretimde: 'In Production', summaryTamamlandi: 'Completed',
      statusHazirlaniyor: 'Preparing', statusUretimde: 'In Production',
      statusTamamlandi: 'Completed', statusIptal: 'Cancelled',
      bantUploadEmpty: 'Ribbon Excel not loaded. Please upload the ribbon plan first.',
      orderSelectTitle: 'Select Order',
      searchBantPlaceholder: 'Search model...',
      searchBantOrderPlaceholder: 'Search model or customer...',
      addRow: 'Add Row', deleteRow: 'Delete Row', deleteSelection: 'Delete Selection',
      rowHint: 'Works only in selected column',
      toastRowAdded: 'row(s) added', toastRowDeleted: 'row(s) deleted',
      bantSearchResult: 'matches', bantSearchNotFound: 'No results found',
      completedTotal: 'Total'
    },

    gerceklesen: {
      title: 'Actual Orders & Cost',
      emptyMsg: 'No actual cost records yet.',
      completedTitle: 'Completed Actual Costs',
      editTitle: 'Edit Actual Cost',
      confirmDelete: 'Are you sure you want to delete this record?',
      confirmDeleteSiparis: 'Are you sure you want to delete this order?',
      toastSaved: 'Actual cost updated!',
      toastDeleted: 'Record deleted',
      toastSiparisDeleted: 'Order deleted',
      toastTransferred: 'Transferred to actual costs!',
      toastStatus: 'Status updated',
      errorLoad: 'Could not load record',
      colImage: 'Image', colCustomer: 'Customer', colCompany: 'Company',
      colModel: 'Model', colTermin: 'Deadline', colPrice: 'Price', colQuantity: 'Qty',
      colKesim: 'Cut', colYuklenen: 'Loaded', colIkinciKalite: '2nd Quality',
      colSatisToplam: 'Sales Total', colKumas: 'Fabric', colAksesuar: 'Accessories',
      colIscilik: 'Labor', colBaski: 'Print/Embroidery/Wash', colTopMaliyet: 'Total Cost',
      colKarZarar: 'Profit/Loss', colKarYuzde: 'Profit %',
      labelKumasBedeli: 'Fabric Cost',
      labelAksesuarBedeli: 'Accessories Cost',
      labelIscilikBedeli: 'Labor Cost',
      labelBaskiNakisYikama: 'Print/Embroidery/Wash',
      labelDigerGider: 'Other Expenses',
      addDetail: 'Add',
      kumasFaturaTitle: 'Fabric Invoices',
      aksesuarFaturaTitle: 'Accessory Invoices',
      excelInfo: 'Leave empty to keep current Excel. Select new Excel to update.',
      openExcel: 'Open Cost Excel',
      summaryCompleted: 'Completed', summaryTotalSales: 'Total Sales',
      summaryTotalCost: 'Total Cost', summaryNetProfit: 'Net Profit / Loss',
      statusDevamEdiyor: 'Ongoing',
      monthlyReport: 'Monthly Profit / Loss Report',
      monthCol: 'Month', salesTotalCol: 'Sales Total', costTotalCol: 'Cost Total',
      profitLossCol: 'Profit / Loss', profitPctCol: 'Profit %', recordCol: 'Records',
      grandTotal: 'GRAND TOTAL',
      yearlyReport: 'Yearly Profit / Loss Report', yearCol: 'Year',
      detailAciklama: 'Description', detailDeger: 'Value'
    },

    odemeler: {
      title: 'Payment List',
      emptyMsg: 'No payments recorded yet.',
      collectedTitle: 'Collected',
      weeklyTitle: 'Weekly Payment Schedule',
      newOdeme: 'New Payment',
      editOdeme: 'Edit Payment',
      emptyWeekly: 'No pending payments yet.',
      emptyCollected: 'No collected payments yet.',
      confirmDelete: 'Are you sure you want to delete this payment?',
      toastSaved: 'Payment saved!',
      toastUpdated: 'Payment updated!',
      toastStatus: 'Status updated',
      toastInstallmentSaved: 'installment(s) saved!',
      toastInstallmentUpdated: 'installment updated',
      toastDeleted: 'Deleted',
      errorLoad: 'Could not load payment list',
      errorWeeklyLoad: 'Could not load weekly view',
      errorValidation: 'Amount and due date are required.',
      errorVade: 'Term must be at least 1 month.',
      errorNoInstallment: 'No installment found to update.',
      colFirma: 'Company / Customer', colTur: 'Type', colTutar: 'Amount',
      colVade: 'Due Date', colCekBanka: 'Check/Bank', colStatus: 'Status',
      colTahsilTarihi: 'Collection Date',
      filterAllStatus: 'All Statuses', filterAllTypes: 'All Types',
      searchPlaceholder: 'Search company, check no or bank...',
      statusBekliyor: 'Pending', statusTahsilEdildi: 'Collected',
      statusGecikti: 'Overdue', statusIptalO: 'Cancelled',
      typeNakit: 'Cash', typeCek: 'Check',
      typeKrediKarti: 'Credit Card',
      typeBankaKredisi: 'Bank Loan', typeSpotKredi: 'Spot Credit', typeRotatifKredi: 'Revolving Credit', typeVadeliOdeme: 'Term Payment',
      typeKira: 'Rent', typeMaas: 'Salary', typeVergi: 'Tax',
      fieldFirma: 'Company / Customer', fieldOdemeTuru: 'Payment Type', fieldTutar: 'Amount',
      fieldVade: 'Due Date', fieldAciklama: 'Description',
      fieldCekNo: 'Check No', fieldBanka: 'Bank Name', fieldCekSahibi: 'Check Issuer',
      fieldKrediSure: 'Duration (months)', fieldKrediTutar: 'Monthly Installment',
      krediHint: 'Automatic payment entries will be created each month from the first due date.',
      krediUpdateHint: 'Remaining amount will be split equally among unpaid installments.',
      installmentLabel: 'Installment', midPayment: 'Mid Payment',
      updateRemaining: 'Update Remaining',
      creditGroupPaid: 'Fully Paid', creditGroupRemaining: 'Remaining',
      creditGroupPaidAmount: 'Paid',
      weekOverdue: 'Overdue', weekThisWeek: 'This Week', weekThisMonth: 'This Month', weekPast: 'Past',
      weekPrefix: 'Week',
      summaryBekleyen: 'Pending', summaryThisWeek: 'This Week',
      summaryGecikmis: 'Overdue', summaryToplamBekleyen: 'Total Pending',
      summaryTahsilEdilen: 'Collected', summaryTotalTahsilat: 'Total Collected',
      summaryBuAyFaiz: 'This Month Interest',
      combinedTitle: 'Income-Expense Report',
      weekGelir: 'Income',
      weekGider: 'Expense',
      weekBakiye: 'Net',
      weekGelirler: 'Income',
      weekGiderler: 'Expenses',
      emptyCombined: 'No income or expense records yet.'
    },

    aksesuar: {
      title: 'Accessory Tracking',
      emptyMsg: 'No accessories added yet.',
      newAksesuar: 'Add New Accessory',
      colName: 'Accessory', colSiparisAdet: 'Order Qty',
      colGelenAdet: 'Received', colBirimFiyat: 'Unit Price', colNot: 'Note',
      confirmDelete: 'Are you sure you want to delete this accessory?',
      toastDeleted: 'Accessory deleted',
      errorAdd: 'Could not add', errorDelete: 'Could not delete',
      sevkiyatTitle: 'Description', sevkiyatAdet: 'Quantity',
      sevkiyatFatura: 'Invoice No', sevkiyatTutar: 'Amount',
      addSevkiyat: 'Add Shipment'
    },

    kumas: {
      title: 'Fabric Tracking',
      emptyMsg: 'No fabrics added yet.',
      newKumas: 'Add New Fabric',
      confirmDelete: 'Are you sure you want to delete this fabric?',
      toastDeleted: 'Fabric deleted',
      toastCalc: 'Unit',
      errorAdd: 'Could not add', errorDelete: 'Could not delete', errorLoad: 'Could not load fabrics',
      colName: 'Fabric Name', colColor: 'Color', colAdet: 'Qty',
      colPastalBoyu: 'Marker L. (m)', colEn: 'Width (m)', colGramaj: 'Weight (g/m²)',
      colFire: 'Waste %', colAsorti: 'Assortment', colBiye: 'Bias (g)',
      colBirimGider: 'Unit Cost (GR/MT)', colKumasSiparis: 'Fabric Order',
      colToplamGelen: 'Total Received', colKumasFiyat: 'Fabric Price',
      colKesimAdet: 'Cut Qty', colKesimYuzde: 'Cut %',
      colGerceklesenBirim: 'Actual Unit',
      sevkiyatMiktar: 'Quantity', sevkiyatFatura: 'Invoice No',
      sevkiyatTutar: 'Amount', addSevkiyat: 'Add Shipment'
    },

    kullanicilar: {
      title: 'User Management',
      newUser: 'New User',
      colUser: 'User', colRol: 'Role', colDate: 'Date',
      fieldUsername: 'Username', fieldPassword: 'Password',
      fieldPasswordLeave: 'Leave blank to keep current',
      fieldRole: 'Role', fieldPermissions: 'Permissions',
      errorUsername: 'Username is required.',
      errorPermission: 'At least one permission must be selected.',
      toastSaved: 'User added', toastUpdated: 'User updated',
      toastDeleted: 'User deleted',
      confirmDeletePrefix: 'Are you sure you want to delete user',
      permLabel: 'permission(s)',
      roles: { admin: 'Admin', muhasebe: 'Accounting', operasyon: 'Operations', satis: 'Sales' }
    },

    bant: {
      title: 'Ribbon Schedule',
      uploadBtn: 'Upload', downloadBtn: 'Download',
      searchPlaceholder: 'Search model...',
      emptyMsg: 'Ribbon Excel not loaded.',
      toastUploadOk: 'Uploaded!',
      toastUploadError: 'Error'
    },

    // --- Revenue Tracking ---
    gelirTakip: {
      title: 'Revenue Evaluation',
      emptyMsg: 'No revenue tracking records yet.',
      editTitle: 'Revenue Evaluation',
      confirmDelete: 'Are you sure you want to delete this revenue record?',
      toastSaved: 'Revenue tracking updated!',
      toastDeleted: 'Revenue tracking deleted',
      errorLoad: 'Could not load revenue tracking list',
      colCustomer: 'Customer', colCompany: 'Company', colModel: 'Model',
      colToplamTutar: 'Total Amount', colNakit: 'Cash', colCek: 'Check',
      colKalan: 'Remaining', colNot: 'Note',
      summaryToplamGelir: 'Total Revenue', summaryNakit: 'Cash',
      summaryCekToplam: 'Check Total', summaryTeminatCek: 'Check as Collateral',
      summaryKirdirilanCek: 'Discounted Check', summaryNetNakit: 'Net Cash',
      fieldNakitTutar: 'Cash Amount',
      cekEkle: '+ Add Check', cekSil: 'Remove Check',
      cekTutar: 'Check Amount', cekVade: 'Due Date', cekBanka: 'Bank',
      cekNo: 'Check No',
      cekDurumBekliyor: 'Pending', cekDurumTeminata: 'As Collateral',
      cekDurumKirdirildi: 'Discounted',
      cekKomisyon: 'Commission / Fee', cekNetTutar: 'Net Amount',
      cekKirdirmaTarihi: 'Discount Date', cekNotlar: 'Notes',
      cekBaslik: 'Check Details',
      noCek: 'No checks'
    },

    cari: {
      title: 'Supplier Accounts',
      emptyMsg: 'No cari cards yet.',
      newCari: '+ New Cari',
      editCari: 'Edit Cari Card',
      confirmDelete: 'Are you sure you want to delete this cari card and all its transactions?',
      colKod: 'Code',
      colAd: 'Name',
      colVergiNo: 'Tax No',
      colTelefon: 'Phone',
      colBorc: 'Debit',
      colAlacak: 'Credit',
      colBakiye: 'Balance',
      ekstre: 'Account Statement',
      hareketler: 'Transactions',
      yeniHareket: '+ Add Transaction',
      hareketDuzenle: 'Edit Transaction',
      turBorc: 'Debit (Purchase)',
      turAlacak: 'Credit (Payment)',
      fieldTarih: 'Date',
      fieldTur: 'Type',
      fieldTutar: 'Amount',
      fieldParaBirimi: 'Currency',
      fieldAciklama: 'Description',
      fieldBelgeNo: 'Document No',
      fieldOdemeSekli: 'Payment Method',
      odemeNakit: 'Cash',
      odemeCek: 'Check',
      odemeHavale: 'Wire/EFT',
      toplamBorc: 'Total Debit',
      toplamAlacak: 'Total Credit',
      kapanisBakiyesi: 'Closing Balance',
      noHareket: 'No transactions for this cari.',
      confirmDeleteHareket: 'Are you sure you want to delete this transaction?',
      toastSaved: 'Cari card saved!',
      toastDeleted: 'Cari card deleted',
      toastHareketSaved: 'Transaction saved!',
      toastHareketDeleted: 'Transaction deleted',
      errorLoad: 'Could not load cari list'
    },

    months: ['January','February','March','April','May','June',
             'July','August','September','October','November','December'],
    monthsShort: ['Jan','Feb','Mar','Apr','May','Jun',
                  'Jul','Aug','Sep','Oct','Nov','Dec'],

    colors: {
      red: 'Red', blue: 'Navy', black: 'Black', white: 'White',
      grey: 'Grey', green: 'Green', yellow: 'Yellow', purple: 'Purple',
      placeholder: 'e.g. Red, Navy', name: 'Color name', qty: 'Qty'
    }
  }
};

// ============================================
// i18n Fonksiyonları
// ============================================
const LANG_KEY = 'ft_lang';

function getLang() {
  return localStorage.getItem(LANG_KEY) || 'tr';
}

function setLang(lang) {
  localStorage.setItem(LANG_KEY, lang);
  location.reload();
}

function toggleLang() {
  setLang(getLang() === 'tr' ? 'en' : 'tr');
}

function t(key, ...args) {
  const lang = getLang();
  const dict = TRANSLATIONS[lang] || TRANSLATIONS['tr'];
  const keys = key.split('.');
  let val = dict;
  for (const k of keys) {
    if (val == null) break;
    val = val[k];
  }
  if (typeof val === 'string') {
    return val.replace(/\{(\d+)\}/g, (_, i) => args[i] !== undefined ? args[i] : '?');
  }
  if (Array.isArray(val)) {
    return args[0] >= 0 && args[0] < val.length ? val[args[0]] : val.join(', ');
  }
  return key; // fallback: show key if missing
}
