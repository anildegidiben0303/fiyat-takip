// ============================================
// Fiyat Takip - Server-side Çeviri Sözlüğü (TR/EN)
// ============================================

const TR = {
  // Auth
  sessionInvalid: 'Oturum geçersiz. Lütfen tekrar giriş yapın.',
  accessDenied: 'Bu sayfaya erişim izniniz yok.',
  imageOnly: 'Sadece resim dosyaları (jpg, png, gif, webp, bmp) yüklenebilir.',
  excelOnly: 'Sadece Excel dosyaları (.xlsx, .xls) yüklenebilir.',
  fileTypeError: 'Sadece resim (jpg,png,gif,webp,bmp) ve Excel (.xlsx,.xls) dosyaları yüklenebilir.',

  // Fiyatlar
  fiyatRequired: 'Müşteri adı, ürün açıklaması ve fiyat zorunludur',
  fiyatSaved: 'Fiyat kaydedildi',
  fiyatUpdated: 'Güncellendi',
  fiyatDeleted: 'Silindi',
  fiyatNotFound: 'Kayıt bulunamadı',

  // Siparişler
  siparisRequired: 'Müşteri adı, ürün açıklaması ve fiyat zorunludur',
  siparisUpdated: 'Sipariş güncellendi',
  siparisDeleted: 'Sipariş silindi',
  siparisNotFound: 'Sipariş bulunamadı',
  siparisStatusUpdated: 'Durum güncellendi',
  siparisReverted: 'Sipariş geri alındı',
  fiyatIdRequired: 'fiyat_id zorunludur',
  fiyatNotFoundById: 'Fiyat kaydı bulunamadı',

  // Gerçekleşen
  gerceklesenRequired: 'Müşteri adı, ürün açıklaması ve fiyat zorunludur',
  gerceklesenUpdated: 'Gerçekleşen güncellendi',
  gerceklesenDeleted: 'Silindi',
  gerceklesenNotFound: 'Kayıt bulunamadı',
  gerceklesenCreated: 'Gerçekleşen kaydı oluşturuldu',
  gerceklesenExists: 'Zaten gerçekleşende mevcut',
  siparisIdRequired: 'siparis_id zorunludur',

  // Ödemeler
  odemeRequired: 'Tutar ve vade tarihi zorunludur',
  odemeSaved: 'Ödeme kaydedildi',
  odemeUpdated: 'Ödeme güncellendi',
  odemeStatusUpdated: 'Durum güncellendi',

  // Aksesuar
  aksesuarRequired: 'Aksesuar adı zorunludur',
  aksesuarCreated: 'Aksesuar eklendi',
  aksesuarUpdated: 'Aksesuar güncellendi',
  aksesuarDeleted: 'Aksesuar silindi',
  aksesuarNotFound: 'Aksesuar bulunamadı',

  // Kumaş
  kumasCreated: 'Kumaş eklendi',
  kumasUpdated: 'Kumaş güncellendi',
  kumasDeleted: 'Kumaş silindi',
  kumasNotFound: 'Kumaş bulunamadı',

  // Kullanıcılar
  userRequired: 'Kullanıcı adı ve parola zorunludur.',
  userExists: 'Bu kullanıcı adı zaten kullanılıyor.',
  userCreated: 'Kullanıcı eklendi',
  userUpdated: 'Kullanıcı güncellendi',
  userDeleted: 'Kullanıcı silindi',
  userNotFound: 'Kullanıcı bulunamadı',
  adminCantDelete: 'Admin kullanıcısı silinemez.',

  // Excel
  excelNotFound: 'Excel dosyası zorunludur',
  excelInvalidSheet: 'Excel dosyasında hiç sayfa bulunamadı',
  excelNotSystem: 'Bu Excel sistem tarafından oluşturulmamış.',
  excelUploaded: 'Excel yüklendi, gerçekleşene aktarıldı',
  excelProcessError: 'Excel işlenemedi',
  excelNoFile: 'Dosya yüklenemedi',
  excelNotSelected: 'Excel seçilmedi.',
  excelUploadedOk: 'Yüklendi!',
  excelDeleted: 'Silindi',

  // Bant
  bantRequired: 'key zorunludur',
  bantSaved: 'Atama kaydedildi',
  bantRowAdded: 'satır eklendi',
  bantRowDeleted: 'satır silindi',
  bantInvalid: 'Geçersiz',
  bantExcelNotFound: 'Bant Excel\'i bulunamadı.',
  bantSheetNotFound: 'Sayfa bulunamadı',
  bantReadError: 'Excel okunamadı',

  // Backup
  backupFailed: 'Yedekleme başarısız',

  // Logout
  logoutOk: 'Çıkış yapıldı',

  // Gelir Takip
  gelirTakipCreated: 'Gelir takip kaydı oluşturuldu',
  gelirTakipUpdated: 'Gelir takip güncellendi',
  gelirTakipDeleted: 'Silindi',
  gelirTakipNotFound: 'Gelir takip kaydı bulunamadı',
  gelirTakipExists: 'Bu sipariş için zaten gelir takip kaydı mevcut',

  // Cari (Tedarikçi Hesapları)
  cariCreated: 'Cari kartı oluşturuldu',
  cariUpdated: 'Cari kartı güncellendi',
  cariDeleted: 'Cari kartı silindi',
  cariNotFound: 'Cari kartı bulunamadı',
  cariAdiRequired: 'Cari adı zorunludur',
  cariHareketCreated: 'Hareket eklendi',
  cariHareketUpdated: 'Hareket güncellendi',
  cariHareketDeleted: 'Hareket silindi',
  cariHareketNotFound: 'Hareket bulunamadı',
  cariHareketRequired: 'Tarih, tür ve tutar zorunludur',

  // Ek
  siparisAktarildi: 'Siparişe aktarıldı',
  bulunamadi: 'Bulunamadı',
  durumZorunlu: 'durum zorunludur',
  ok: 'ok'
};

const EN = {
  sessionInvalid: 'Session invalid. Please log in again.',
  accessDenied: 'You do not have permission to access this page.',
  imageOnly: 'Only image files (jpg, png, gif, webp, bmp) can be uploaded.',
  excelOnly: 'Only Excel files (.xlsx, .xls) can be uploaded.',
  fileTypeError: 'Only image (jpg,png,gif,webp,bmp) and Excel (.xlsx,.xls) files can be uploaded.',

  fiyatRequired: 'Customer name, product description and price are required',
  fiyatSaved: 'Price saved',
  fiyatUpdated: 'Updated',
  fiyatDeleted: 'Deleted',
  fiyatNotFound: 'Record not found',

  siparisRequired: 'Customer name, product description and price are required',
  siparisUpdated: 'Order updated',
  siparisDeleted: 'Order deleted',
  siparisNotFound: 'Order not found',
  siparisStatusUpdated: 'Status updated',
  siparisReverted: 'Order reverted',
  fiyatIdRequired: 'fiyat_id is required',
  fiyatNotFoundById: 'Price record not found',

  gerceklesenRequired: 'Customer name, product description and price are required',
  gerceklesenUpdated: 'Actual cost updated',
  gerceklesenDeleted: 'Deleted',
  gerceklesenNotFound: 'Record not found',
  gerceklesenCreated: 'Actual cost record created',
  gerceklesenExists: 'Already exists in actual costs',
  siparisIdRequired: 'siparis_id is required',

  odemeRequired: 'Amount and due date are required',
  odemeSaved: 'Payment saved',
  odemeUpdated: 'Payment updated',
  odemeStatusUpdated: 'Status updated',

  aksesuarRequired: 'Accessory name is required',
  aksesuarCreated: 'Accessory added',
  aksesuarUpdated: 'Accessory updated',
  aksesuarDeleted: 'Accessory deleted',
  aksesuarNotFound: 'Accessory not found',

  kumasCreated: 'Fabric added',
  kumasUpdated: 'Fabric updated',
  kumasDeleted: 'Fabric deleted',
  kumasNotFound: 'Fabric not found',

  userRequired: 'Username and password are required.',
  userExists: 'This username is already taken.',
  userCreated: 'User added',
  userUpdated: 'User updated',
  userDeleted: 'User deleted',
  userNotFound: 'User not found',
  adminCantDelete: 'Admin user cannot be deleted.',

  excelNotFound: 'Excel file is required',
  excelInvalidSheet: 'No sheet found in Excel file',
  excelNotSystem: 'This Excel was not created by the system.',
  excelUploaded: 'Excel uploaded, transferred to actual costs',
  excelProcessError: 'Could not process Excel',
  excelNoFile: 'File could not be uploaded',
  excelNotSelected: 'No Excel selected.',
  excelUploadedOk: 'Uploaded!',
  excelDeleted: 'Deleted',

  bantRequired: 'key is required',
  bantSaved: 'Assignment saved',
  bantRowAdded: 'row(s) added',
  bantRowDeleted: 'row(s) deleted',
  bantInvalid: 'Invalid',
  bantExcelNotFound: 'Ribbon Excel not found.',
  bantSheetNotFound: 'Sheet not found',
  bantReadError: 'Could not read Excel',

  backupFailed: 'Backup failed',

  logoutOk: 'Logged out',

  // Revenue Tracking
  gelirTakipCreated: 'Revenue tracking record created',
  gelirTakipUpdated: 'Revenue tracking updated',
  gelirTakipDeleted: 'Deleted',
  gelirTakipNotFound: 'Revenue tracking record not found',
  gelirTakipExists: 'Revenue tracking record already exists for this order',

  // Cari (Supplier Accounts)
  cariCreated: 'Cari card created',
  cariUpdated: 'Cari card updated',
  cariDeleted: 'Cari card deleted',
  cariNotFound: 'Cari card not found',
  cariAdiRequired: 'Cari name is required',
  cariHareketCreated: 'Transaction added',
  cariHareketUpdated: 'Transaction updated',
  cariHareketDeleted: 'Transaction deleted',
  cariHareketNotFound: 'Transaction not found',
  cariHareketRequired: 'Date, type and amount are required',

  // Extra
  siparisAktarildi: 'Transferred to orders',
  bulunamadi: 'Not found',
  durumZorunlu: 'status is required',
  ok: 'ok'
};

function t(lang, key) {
  const dict = lang === 'en' ? EN : TR;
  return dict[key] || TR[key] || key;
}

module.exports = { t };
