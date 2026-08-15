/**
 * =========================================================
 * FILE: AdminAuth.gs
 * FUNGSI: Validasi Password Admin untuk fitur "Mode Admin" --
 *         dijalankan 100% di server (Apps Script), BUKAN di
 *         HTML/JavaScript, sehingga tidak bisa dibuka paksa
 *         hanya dengan memodifikasi browser/console.
 *
 * CARA GANTI PASSWORD:
 *  Opsi 1 (PALING MUDAH): ubah nilai ADMIN_PASSWORD_DEFAULT di
 *  bawah ini langsung dari Apps Script Editor, lalu simpan &
 *  deploy ulang Web App.
 *
 *  Opsi 2 (tanpa perlu edit kode / redeploy): buka menu
 *  Project Settings (ikon gerigi) di Apps Script Editor -> scroll
 *  ke "Script Properties" -> tambahkan properti baru dengan
 *  Key = ADMIN_PASSWORD dan Value = password baru Anda.
 *  Jika Script Property ini diisi, nilainya akan dipakai dan
 *  MENGGANTIKAN ADMIN_PASSWORD_DEFAULT di bawah -- tidak perlu
 *  edit kode maupun deploy ulang setiap kali ganti password.
 *
 * CARA PASANG:
 *  1. Buat file script baru di Apps Script Editor bernama "AdminAuth.gs".
 *  2. Salin seluruh isi file ini ke dalamnya.
 *  3. (Opsional) Ganti ADMIN_PASSWORD_DEFAULT di bawah ini.
 *  4. Deploy ulang / refresh Web App.
 * =========================================================
 */

// GANTI PASSWORD INI sesuai kebutuhan Anda (atau pakai Script Properties, lihat catatan di atas).
const ADMIN_PASSWORD_DEFAULT = "ovin123";

// Dipanggil dari frontend (google.script.run.cekPasswordAdmin(password)) saat
// user mengklik "Masuk" di popup Mode Admin. Mengembalikan true/false saja --
// password ASLI tidak pernah dikirim balik ke browser.
function cekPasswordAdmin(inputPassword) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const configuredPassword = scriptProps.getProperty('ADMIN_PASSWORD') || ADMIN_PASSWORD_DEFAULT;

    if (!inputPassword) return false;

    return String(inputPassword) === String(configuredPassword);
  } catch (e) {
    // Kalau terjadi error tak terduga, JANGAN pernah anggap valid -- gagal aman (fail-safe).
    return false;
  }
}