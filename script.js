function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('RT Smart - Portal Warga RT 04 RW 13')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getNamaBulanDariTanggal(tanggalObj) {
  var daftarBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var d = new Date(tanggalObj);
  return isNaN(d.getTime()) ? null : daftarBulan[d.getMonth()];
}

// -------------------------------------------------------------
// KEUANGAN & CEK TAGIHAN MANDIRI
// -------------------------------------------------------------
function getDashboardData(bulanPilihan) {
  var sheetKas = getSheet('Pemasukan') ? getSheet('Pemasukan').getDataRange().getValues() : [];
  var sheetKeluar = getSheet('Pengeluaran') ? getSheet('Pengeluaran').getDataRange().getValues() : [];
  var sheetLain = getSheet('PemasukanLain') ? getSheet('PemasukanLain').getDataRange().getValues() : [];

  var daftarBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var targetBulanIndex = daftarBulan.indexOf(bulanPilihan);

  var totalKasRT = 0;
  var totalRonda = 0;
  var totalKeluarKas = 0;
  var totalKeluarRonda = 0;
  var saldoAwal = 0;

  var rekapGang = {};
  var totalPemasukanBulanIni = 0;
  var listPengeluaranBulan = [];
  var totalPengeluaranBulanIni = 0;

  for (var i = 1; i < sheetKas.length; i++) {
    var tglBayar = sheetKas[i][0];
    var namaBulanBayar = getNamaBulanDariTanggal(tglBayar);
    var gang = sheetKas[i][2];
    var kas = Number(sheetKas[i][4]) || 0;
    var ronda = Number(sheetKas[i][5]) || 0;

    totalKasRT += kas;
    totalRonda += ronda;

    var bIndex = daftarBulan.indexOf(namaBulanBayar);
    if (bIndex !== -1 && bIndex < targetBulanIndex) {
      saldoAwal += (kas + ronda);
    }

    if (namaBulanBayar === bulanPilihan) {
      totalPemasukanBulanIni += (kas + ronda);
      if (!rekapGang[gang]) rekapGang[gang] = { kas: 0, ronda: 0 };
      rekapGang[gang].kas += kas;
      rekapGang[gang].ronda += ronda;
    }
  }

  for (var k = 1; k < sheetLain.length; k++) {
    var tglLain = sheetLain[k][0];
    var namaBulanLain = getNamaBulanDariTanggal(tglLain);
    var nomLain = Number(sheetLain[k][2]) || 0;
    
    totalKasRT += nomLain;
    var bIndexLain = daftarBulan.indexOf(namaBulanLain);
    if (bIndexLain !== -1 && bIndexLain < targetBulanIndex) {
      saldoAwal += nomLain;
    }
    if (namaBulanLain === bulanPilihan) {
      totalPemasukanBulanIni += nomLain;
    }
  }

  for (var j = 1; j < sheetKeluar.length; j++) {
    var tglKeluar = sheetKeluar[j][0];
    var namaBulanKeluar = getNamaBulanDariTanggal(tglKeluar);
    var sumber = sheetKeluar[j][2];
    var nomKeluar = Number(sheetKeluar[j][3]) || 0;
    var ket = sheetKeluar[j][4];

    if (sumber === "Kas RT") totalKeluarKas += nomKeluar;
    if (sumber === "Uang Ronda") totalKeluarRonda += nomKeluar;

    var bIndexKeluar = daftarBulan.indexOf(namaBulanKeluar);
    if (bIndexKeluar !== -1 && bIndexKeluar < targetBulanIndex) {
      saldoAwal -= nomKeluar;
    }

    if (namaBulanKeluar === bulanPilihan) {
      totalPengeluaranBulanIni += nomKeluar;
      listPengeluaranBulan.push({ bulan: namaBulanKeluar, sumber: sumber, nominal: nomKeluar, keterangan: ket });
    }
  }

  var rekapGangArray = [];
  for (var g in rekapGang) {
    rekapGangArray.push({ gang: g, kas: rekapGang[g].kas, ronda: rekapGang[g].ronda, total: rekapGang[g].kas + rekapGang[g].ronda });
  }

  return {
    saldoKasRT: totalKasRT - totalKeluarKas,
    saldoRonda: totalRonda - totalKeluarRonda,
    saldoAwalBulan: saldoAwal,
    pemasukanBulanIni: totalPemasukanBulanIni,
    pengeluaranBulanIni: totalPengeluaranBulanIni,
    saldoAkhirBulan: saldoAwal + totalPemasukanBulanIni - totalPengeluaranBulanIni,
    rekapGang: rekapGangArray,
    pengeluaran: listPengeluaranBulan
  };
}

function getFinancialChartData() {
  var sheetKas = getSheet('Pemasukan') ? getSheet('Pemasukan').getDataRange().getValues() : [];
  var sheetKeluar = getSheet('Pengeluaran') ? getSheet('Pengeluaran').getDataRange().getValues() : [];
  var sheetLain = getSheet('PemasukanLain') ? getSheet('PemasukanLain').getDataRange().getValues() : [];
  
  var daftarBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var kasPerBulan = new Array(12).fill(0);
  var rondaPerBulan = new Array(12).fill(0);
  
  for (var i = 1; i < sheetKas.length; i++) {
    var tglBayar = sheetKas[i][0];
    var namaBulanBayar = getNamaBulanDariTanggal(tglBayar);
    var kas = Number(sheetKas[i][4]) || 0;
    var ronda = Number(sheetKas[i][5]) || 0;
    var idx = daftarBulan.indexOf(namaBulanBayar);
    if (idx !== -1) {
      kasPerBulan[idx] += kas;
      rondaPerBulan[idx] += ronda;
    }
  }

  for (var k = 1; k < sheetLain.length; k++) {
    var tglLain = sheetLain[k][0];
    var namaBulanLain = getNamaBulanDariTanggal(tglLain);
    var nomLain = Number(sheetLain[k][2]) || 0;
    var idxLain = daftarBulan.indexOf(namaBulanLain);
    if (idxLain !== -1) {
      kasPerBulan[idxLain] += nomLain;
    }
  }

  for (var j = 1; j < sheetKeluar.length; j++) {
    var tglKeluar = sheetKeluar[j][0];
    var namaBulanKeluar = getNamaBulanDariTanggal(tglKeluar);
    var sumber = sheetKeluar[j][2];
    var nomKeluar = Number(sheetKeluar[j][3]) || 0;
    var idxKeluar = daftarBulan.indexOf(namaBulanKeluar);
    if (idxKeluar !== -1) {
      if (sumber === "Kas RT") kasPerBulan[idxKeluar] -= nomKeluar;
      else if (sumber === "Uang Ronda") rondaPerBulan[idxKeluar] -= nomKeluar;
    }
  }

  return {
    labels: ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'],
    kasRT: kasPerBulan,
    kompRonda: rondaPerBulan
  };
}

function cekTagihanMandiriByBlok(queryBlok, bulanTarget) {
  var sheetWarga = getSheet('Warga');
  var dataWarga = sheetWarga ? sheetWarga.getDataRange().getValues() : [];
  var foundNama = "";
  var foundBlok = "";

  var q = queryBlok.toString().trim().toLowerCase();
  for (var i = 1; i < dataWarga.length; i++) {
    var blokWarga = dataWarga[i][2] ? dataWarga[i][2].toString().trim().toLowerCase() : "";
    if (blokWarga.includes(q)) {
      foundBlok = dataWarga[i][2];
      foundNama = dataWarga[i][3];
      break;
    }
  }

  if (!foundNama) {
    return { nama: null };
  }

  var sheetKas = getSheet('Pemasukan');
  var dataKas = sheetKas ? sheetKas.getDataRange().getValues() : [];
  var kasSudahBayar = false;
  var nominalKasSudahBayar = 0;

  for (var j = 1; j < dataKas.length; j++) {
    var b = dataKas[j][1];
    var n = dataKas[j][3] ? dataKas[j][3].toString().trim().toLowerCase() : "";
    if (b === bulanTarget && n === foundNama.trim().toLowerCase()) {
      var kVal = Number(dataKas[j][4]) || 0;
      if (kVal > 0) {
        kasSudahBayar = true;
        nominalKasSudahBayar = kVal;
      }
    }
  }

  var rekapRonda = getRekapRondaAdmin(bulanTarget, "2026");
  var nominalKompensasi = 30000;
  var statusRondaInfo = "Tidak Ronda (Kompensasi Penuh)";
  var jumlahRondaWarga = 0;

  for (var k = 0; k < rekapRonda.length; k++) {
    if (rekapRonda[k].nama.trim().toLowerCase() === foundNama.trim().toLowerCase()) {
      nominalKompensasi = rekapRonda[k].tagihanKompensasi;
      statusRondaInfo = rekapRonda[k].statusKompensasi;
      jumlahRondaWarga = rekapRonda[k].jumlahRonda;
    }
  }

  var tagihanKasNominal = kasSudahBayar ? 0 : 12000;
  var totalTagihanFinal = tagihanKasNominal + nominalKompensasi;

  return {
    nama: foundNama,
    blokNo: foundBlok,
    kasSudahBayar: kasSudahBayar,
    nominalKas: tagihanKasNominal,
    jumlahRonda: jumlahRondaWarga,
    nominalRonda: nominalKompensasi,
    statusRondaInfo: statusRondaInfo,
    totalTagihan: totalTagihanFinal
  };
}

function kirimBuktiBayar(payload) {
  var sheet = getSheet('KonfirmasiBayar');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('KonfirmasiBayar');
    sheet.appendRow(['Timestamp', 'NamaWarga', 'JenisTagihan', 'Keterangan', 'BuktiFoto', 'Status']);
  }
  sheet.appendRow([new Date(), payload.nama, payload.jenis, payload.keterangan, payload.fotoUrl, 'Belum Dikonfirmasi']);
  return { status: "success", message: "Bukti pembayaran berhasil dikirim ke Admin!" };
}

function getNotifikasiAdmin() {
  var sheet = getSheet('KonfirmasiBayar');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] === 'Belum Dikonfirmasi') {
      list.push({
        row: i + 1,
        tanggal: new Date(data[i][0]).toLocaleString(),
        nama: data[i][1],
        jenis: data[i][2],
        keterangan: data[i][3],
        foto: data[i][4]
      });
    }
  }
  return list;
}

function konfirmasiBayarSelesai(rowIdx) {
  var sheet = getSheet('KonfirmasiBayar');
  sheet.getRange(rowIdx, 6).setValue('Selesai');
  return "Konfirmasi selesai!";
}

// -------------------------------------------------------------
// KEAMANAN RONDA & VERIFIKASI ADMIN
// -------------------------------------------------------------
function getDaftarJadwalRonda() {
  var sheet = getSheet('JadwalRonda');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({ id: data[i][0], pos: data[i][1], hari: data[i][2], blokNo: data[i][3], nama: data[i][4] });
  }
  return list;
}

function getSemuaMasterWarga() {
  var sheet = getSheet('Warga') || getSheet('JadwalRonda');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  var mapCek = {};
  for (var i = 1; i < data.length; i++) {
    var nama = data[i][3] || data[i][4];
    var blok = data[i][2] || data[i][3];
    var key = nama + " (" + blok + ")";
    if (nama && !mapCek[key]) {
      mapCek[key] = true;
      list.push({ nama: nama, blokNo: blok });
    }
  }
  return list;
}

function simpanAbsenRondaHarian(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('KonfirmasiAbsenRonda');
  if (!sheet) {
    sheet = ss.insertSheet('KonfirmasiAbsenRonda');
    sheet.appendRow(['Timestamp', 'Hari', 'Pos', 'BlokNo', 'NamaWarga', 'StatusAbsen', 'FotoSelfie', 'StatusVerifikasi']);
  }

  var now = new Date();
  sheet.appendRow([now, payload.hari, payload.pos, payload.blokNo, payload.nama, payload.statusAbsen, payload.fotoUrl, 'Belum Dikonfirmasi']);

  return {
    status: "success",
    message: "Absen ronda dan foto selfie berhasil dikirim ke Admin untuk diverifikasi!"
  };
}

function getNotifikasiAbsenAdmin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('KonfirmasiAbsenRonda');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === 'Belum Dikonfirmasi') {
      list.push({
        row: i + 1,
        tanggal: new Date(data[i][0]).toLocaleString(),
        hari: data[i][1],
        pos: data[i][2],
        blokNo: data[i][3],
        nama: data[i][4],
        statusAbsen: data[i][5],
        foto: data[i][6]
      });
    }
  }
  return list;
}

function prosesVerifikasiAbsenRonda(rowIdx, aksi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKonfirm = ss.getSheetByName('KonfirmasiAbsenRonda');
  if (!sheetKonfirm) return "Error";

  var rowValues = sheetKonfirm.getRange(rowIdx, 1, 1, 8).getValues()[0];
  var timestamp = rowValues[0];
  var hari = rowValues[1];
  var pos = rowValues[2];
  var blokNo = rowValues[3];
  var nama = rowValues[4];
  var statusAbsen = rowValues[5];
  var fotoUrl = rowValues[6];

  if (aksi === 'Terima') {
    sheetKonfirm.getRange(rowIdx, 8).setValue('Diterima');
    var sheetRonda = ss.getSheetByName('Ronda');
    if (!sheetRonda) {
      sheetRonda = ss.insertSheet('Ronda');
      sheetRonda.appendRow(['Timestamp', 'Hari', 'Pos', 'BlokNo', 'NamaWarga', 'Status', 'FotoSelfie']);
    }
    sheetRonda.appendRow([timestamp, hari, pos, blokNo, nama, statusAbsen, fotoUrl]);
    return "Absen berhasil diterima dan dicatat!";
  } else {
    sheetKonfirm.getRange(rowIdx, 8).setValue('Ditolak');
    return "Absen telah ditolak.";
  }
}

// -------------------------------------------------------------
// ADMIN & REKAP
// -------------------------------------------------------------
function getDaftarGang() {
  var data = getSheet('Warga').getDataRange().getValues();
  var listGang = [];
  for (var i = 1; i < data.length; i++) {
    var g = data[i][1];
    if (g && listGang.indexOf(g) === -1) listGang.push(g);
  }
  return listGang;
}

function getWargaByGang(gang) {
  var data = getSheet('Warga').getDataRange().getValues();
  var listWarga = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === gang) {
      listWarga.push({ id: data[i][0], blokNo: data[i][2], nama: data[i][3] });
    }
  }
  return listWarga;
}

function simpanIuranWarga(payload) {
  var sheet = getSheet('Pemasukan');
  var now = new Date();
  payload.data.forEach(function(item) {
    if (item.kas > 0 || item.ronda > 0) {
      sheet.appendRow([now, payload.bulan, payload.gang, item.nama, item.kas, item.ronda, 'Setoran Iuran']);
    }
  });
  return "Berhasil menyimpan iuran!";
}

function simpanPemasukanLain(bulan, nominal, keterangan) {
  getSheet('PemasukanLain').appendRow([new Date(), bulan, nominal, keterangan]);
  return "Berhasil!";
}

function simpanPengeluaran(bulan, sumber, nominal, keterangan) {
  getSheet('Pengeluaran').appendRow([new Date(), bulan, sumber, nominal, keterangan]);
  return "Berhasil!";
}

function getRekapRondaAdmin(bulanTarget, tahunTarget) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRonda = ss.getSheetByName('Ronda');
  var sheetWarga = ss.getSheetByName('Warga');
  var daftarBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  
  var dataWarga = sheetWarga ? sheetWarga.getDataRange().getValues() : [];
  var wargaMap = {};
  for (var i = 1; i < dataWarga.length; i++) {
    var gang = dataWarga[i][1];
    var blokNo = dataWarga[i][2];
    var nama = dataWarga[i][3];
    if (nama) {
      wargaMap[nama.toString().trim().toLowerCase()] = { nama: nama, blokNo: blokNo, gang: gang || "Umum", jumlahRonda: 0, statusBayar: "Belum Bayar" };
    }
  }

  if (sheetRonda) {
    var dataRonda = sheetRonda.getDataRange().getValues();
    for (var j = 1; j < dataRonda.length; j++) {
      var tgl = new Date(dataRonda[j][0]);
      var bRonda = daftarBulan[tgl.getMonth()];
      var thnRonda = tgl.getFullYear().toString();
      var namaRonda = dataRonda[j][4] ? dataRonda[j][4].toString().trim().toLowerCase() : "";
      if (bRonda === bulanTarget && thnRonda === tahunTarget.toString() && wargaMap[namaRonda]) {
        wargaMap[namaRonda].jumlahRonda++;
      }
    }
  }

  var hasilRekap = [];
  for (var key in wargaMap) {
    var item = wargaMap[key];
    var tagihan = 30000;
    var statusKompensasi = "Tidak Ronda (Kompensasi Penuh)";
    if (item.jumlahRonda === 1) { tagihan = 15000; statusKompensasi = "Kurang 1x (Kompensasi 50%)"; }
    else if (item.jumlahRonda >= 2) { tagihan = 0; statusKompensasi = "Aman (Capai Target)"; }

    hasilRekap.push({ nama: item.nama, blokNo: item.blokNo, gang: item.gang, jumlahRonda: item.jumlahRonda, tagihanKompensasi: tagihan, statusKompensasi: statusKompensasi, statusBayar: item.statusBayar });
  }
  return hasilRekap;
}