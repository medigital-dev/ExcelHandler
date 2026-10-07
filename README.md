# ExcelHandler

Library PHP untuk CodeIgniter 4 yang menyediakan dua cara bekerja dengan file
Excel:

- `ExcelHandler` untuk membaca file, membuat workbook, mengisi template, dan
  menambahkan validasi/styling. Library ini menggunakan PhpSpreadsheet dan
  memuat workbook ke memori.
- `ExcelStreamWriter` untuk membuat file `.xlsx` berukuran besar dengan
  penulisan baris secara streaming menggunakan OpenSpout.

## Kebutuhan

- PHP `^8.4`
- CodeIgniter 4 `^4.4`
- PhpSpreadsheet `^5.10`
- OpenSpout `^5.12`

## Instalasi

Tambahkan package melalui Composer:

```bash
composer require medigital-dev/excel-handler
```

Jika source library digunakan langsung dari repository, jalankan:

```bash
composer install
```

Namespace class yang tersedia:

```php
use MedigitalDev\ExcelHandler\Libraries\ExcelHandler;
use MedigitalDev\ExcelHandler\Libraries\ExcelStreamWriter;
```

## Memilih library

| Kebutuhan | Library |
| --- | --- |
| Membaca `.xls` atau `.xlsx` | `ExcelHandler` |
| Memproses file input dalam beberapa batch | `ExcelHandler::chunk()` |
| Membuat workbook biasa, mengatur cell, sheet, style, atau dropdown | `ExcelHandler` |
| Mengisi dan menyimpan template workbook yang sudah ada | `ExcelHandler::openForEdit()` |
| Export dataset besar dengan penggunaan memori rendah | `ExcelStreamWriter` |

`ExcelHandler` menyimpan workbook untuk operasi tulis/edit di memori. Untuk
export dengan banyak baris, gunakan `ExcelStreamWriter`; class tersebut hanya
membuat file `.xlsx` baru dan tidak mendukung membaca, mengedit template,
dropdown, atau styling per-cell.

## ExcelHandler

### Membaca file dari path

`load()` membuka file untuk dibaca. Gunakan `toArray()` untuk membaca seluruh
sheet:

```php
$excel = new ExcelHandler();
$rows = $excel
    ->load(WRITEPATH . 'uploads/siswa.xlsx')
    ->toArray(); // Sheet pertama; baris pertama dianggap header
```

Secara default, header dinormalisasi menjadi key huruf kecil dengan spasi
sebagai underscore. Contohnya, header `Nama Siswa` menjadi `nama_siswa`.
Header kosong menjadi `kolom_N`, header duplikat mendapat akhiran seperti
`_2`, dan baris yang seluruh nilainya kosong dilewati. Nilai string di-trim
secara default.

Opsi `toArray($sheet, $hasHeader, $trim)`:

- `$sheet`: index sheet mulai dari `0` atau nama sheet.
- `$hasHeader`: `true` untuk header di baris pertama (default), `false` untuk
  hasil tanpa header, atau angka mulai dari `0` untuk menentukan baris header.
  Baris sebelum dan termasuk baris header dilewati.
- `$trim`: trim spasi pada nilai string (default `true`).

```php
$sheetNames = $excel->getSheetNames();
$rawRows = $excel->toArray('Data', hasHeader: false);
$rowsWithHeaderOnThirdRow = $excel->toArray('Data', hasHeader: 2);
```

### Membaca upload CodeIgniter 4

`read()` menerima `UploadedFile` CodeIgniter. Ekstensi dan isi file diperiksa;
hanya file Excel `.xls` dan `.xlsx` yang diterima. File diproses dari lokasi
sementara upload dan tidak dipindahkan.

```php
$file = $this->request->getFile('spreadsheet');

$rows = (new ExcelHandler())
    ->read($file)
    ->toArray();
```

File sementara hanya tersedia selama request. Jika file perlu diproses oleh
queue atau cron, pindahkan/simpan upload terlebih dahulu lalu gunakan `load()`
dengan path file tersebut.

### Memproses file input per batch

Untuk input besar, `chunk()` menghasilkan batch agar seluruh data tidak perlu
ditampung sekaligus. Ukuran default adalah 500 baris; ukuran sheet dibaca ulang
untuk setiap batch, sehingga batch yang lebih besar (misalnya 1.000–2.000)
dapat lebih sesuai untuk file yang sangat besar.

```php
$excel = (new ExcelHandler())->load(WRITEPATH . 'uploads/siswa.xlsx');

foreach ($excel->chunk(1000, 'Sheet1') as $batch) {
    // Contoh: validasi atau simpan batch ke database.
    $model->insertBatch($batch);
}
```

Parameter sheet, header, dan trim mengikuti aturan `toArray()`:

```php
foreach ($excel->chunk(500, 0, hasHeader: 1, trim: true) as $batch) {
    // Baris kedua file digunakan sebagai header.
}
```

### Membuat workbook baru

`setHeaders()` menulis header pada baris pertama. `writeRows()` menambahkan
baris mulai dari baris kosong berikutnya. Untuk setiap baris, urutan value yang
dipakai; key array asosiatif tidak dicocokkan dengan header.

```php
$excel = new ExcelHandler();
$excel->setHeaders(['NIS', 'Nama', 'Kelas'])
    ->writeRows([
        ['1001', 'Ayu', 'X-A'],
        ['1002', 'Bima', 'X-B'],
    ])
    ->styleHeaderRow()
    ->autoSizeColumns();

$path = $excel->save(WRITEPATH . 'exports/siswa.xlsx');
```

Untuk membuat beberapa sheet, pilih sheet sebelum menulis atau berikan nama
sheet sebagai argumen:

```php
$excel = new ExcelHandler();
$excel->setHeaders(['Bulan', 'Total'], 'Ringkasan')
    ->writeRows([['Januari', 120]], 'Ringkasan')
    ->setHeaders(['Tanggal', 'Keterangan'], 'Detail')
    ->writeRows([['2026-01-01', 'Contoh']], 'Detail');
```

Metode tulis lainnya:

```php
$excel->writeRow(['1003', 'Citra', 'X-C']); // Satu baris
$excel->setCell('B3', 'Nama');             // Koordinat Excel
$excel->setCell([2, 3], 'Nama');            // Koordinat [kolom, baris], 1-based
$excel->setCells(['B3' => 'Nama', 'C3' => 'X-A']);
$value = $excel->getCell('B3');
```

String angka berisi lebih dari 15 digit (misalnya NIK atau nomor rekening)
yang ditulis melalui `writeRows()` disimpan sebagai teks agar digitnya tidak
kehilangan presisi di Excel.

### Mengisi template yang sudah ada

Gunakan `openForEdit()` untuk memuat seluruh workbook template ke memori.
Perubahan cell dan penyimpanan berikutnya diterapkan ke workbook tersebut.

```php
$excel = new ExcelHandler();
$excel->openForEdit(APPPATH . 'templates/rekap.xlsx')
    ->setCell('C4', 'Yogyakarta')
    ->setCell('C5', date('d-m-Y'))
    ->save(WRITEPATH . 'exports/rekap-final.xlsx');
```

Jangan gunakan `load()` untuk mengedit file. `load()` hanya untuk operasi baca;
untuk perubahan pada file yang sudah ada, gunakan `openForEdit()`.

### Dropdown dan validasi data

`setDropdown()` menerima satu cell, beberapa cell, atau range. Untuk opsi
panjang atau opsi yang mengandung koma/tanda kutip, opsi disimpan pada sheet
bantuan tersembunyi. Berikan `sourceKey` berbeda untuk setiap daftar yang
berbeda dalam workbook yang sama.

```php
$excel->setDropdown(
    'D2:D200',
    ['Aktif', 'Tidak Aktif'],
    sourceKey: 'status'
);
```

`setDataValidation()` mendukung tipe `list`, `whole`, `decimal`, `date`,
`textLength`, dan `custom`. Contoh membatasi nilai bulat:

```php
$excel->setDataValidation('E2:E200', [
    'type'  => 'whole',
    'min'   => 0,
    'max'   => 100,
    'error' => 'Nilai harus antara 0 dan 100.',
]);
```

Validasi tanggal menerima string tanggal atau `DateTimeInterface` sebagai
batas. Tipe `custom` menggunakan key `formula`, misalnya `'=B2>0'`.

### Download dan format simpan

`download()` menghasilkan response download CodeIgniter 4. Gunakan dari
controller:

```php
public function export()
{
    $excel = new ExcelHandler();
    $excel->setHeaders(['Kode', 'Nama'])
        ->writeRows([['A01', 'Contoh']]);

    return $excel->download('data.xlsx');
}
```

Format simpan yang didukung oleh implementasi adalah `Xlsx`, `Xls`, `Csv`, dan
`Ods`:

```php
$excel->save(WRITEPATH . 'exports/data.csv', 'Csv');
```

`download()` menggunakan file sementara dan menghapusnya setelah diproses.
Untuk export yang sangat besar, jangan gunakan `download()` karena workbook
akan dimuat ke memori sebagai body response.

## ExcelStreamWriter

`ExcelStreamWriter` menulis baris langsung ke file `.xlsx`. Folder tujuan
dibuat otomatis. Jika diberikan, header ditulis tebal dengan latar biru dan
teks putih.

```php
$writer = new ExcelStreamWriter();
$writer->open(WRITEPATH . 'exports/transaksi.xlsx', [
    'ID',
    'Tanggal',
    'Nominal',
]);

foreach ($batches as $batch) {
    $writer->writeRows($batch);
}

$path = $writer->close();
```

`writeRow()` menulis satu baris dan `writeRows()` menulis sekumpulan baris.
Keduanya memakai urutan value pada array; key asosiatif diabaikan.

```php
$writer = new ExcelStreamWriter();
$writer->open(WRITEPATH . 'exports/sederhana.xlsx')
    ->writeRow(['Kode', 'Nama'])
    ->writeRows([
        ['A01', 'Ayu'],
        ['A02', 'Bima'],
    ]);

$path = $writer->close();
```

`close()` wajib dipanggil setelah penulisan selesai agar file diselesaikan
dengan benar. Method tersebut mengembalikan path file. `isOpen()` dapat
digunakan untuk memeriksa status writer. Setelah `close()`, instance dapat
digunakan kembali dengan `open()` untuk file lain.

Writer tidak boleh dibuka ulang sebelum stream sebelumnya ditutup. Menulis
sebelum `open()` atau memanggil `close()` saat tidak ada stream terbuka akan
menghasilkan `RuntimeException`. Destructor mencoba menutup stream yang masih
terbuka, tetapi pemanggilan `close()` eksplisit tetap disarankan.

File yang dibuat dapat dikirim sebagai download melalui response CodeIgniter:

```php
$writer = new ExcelStreamWriter();
$writer->open(WRITEPATH . 'exports/transaksi.xlsx', ['ID', 'Nama']);
$writer->writeRows($rows);
$path = $writer->close();

return $this->response->download($path, null)->setFileName('transaksi.xlsx');
```

## Catatan

- `ExcelHandler::toArray()` membaca seluruh sheet ke memori; gunakan `chunk()`
  untuk input yang besar.
- `ExcelHandler::openForEdit()` memuat workbook ke memori.
- `ExcelStreamWriter` hanya untuk membuat file `.xlsx` baru secara streaming.
- Pastikan direktori `WRITEPATH` dapat ditulis oleh proses PHP.
- Operasi file dapat melempar `RuntimeException` jika file tidak tersedia,
  stream belum dibuka, atau konfigurasi validasi tidak sesuai.
