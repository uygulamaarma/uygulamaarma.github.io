/* ArmaLife Stok Sorgulama — Excel → veri dönüştürücü
   Hem telefonda (tarayıcı) hem de bilgisayarda (Node) aynı kodla çalışır. */
(function (root) {
  'use strict';

  // Başlık adlarını karşılaştırılabilir hale getir: büyük harf, Türkçe karakterleri sadeleştir, boşluk/noktalama sil
  function norm(s) {
    return String(s == null ? '' : s)
      .toLocaleUpperCase('tr-TR')
      .replace(/İ/g, 'I').replace(/I/g, 'I').replace(/Ş/g, 'S').replace(/Ğ/g, 'G')
      .replace(/Ü/g, 'U').replace(/Ö/g, 'O').replace(/Ç/g, 'C').replace(/Â/g, 'A')
      .replace(/[^A-Z0-9]/g, '');
  }

  function temizKod(v) {
    if (v == null) return '';
    var s = (typeof v === 'number') ? String(Math.round(v)) : String(v);
    s = s.trim();
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    return s;
  }
  function kodAnahtari(v) { return temizKod(v).toLocaleUpperCase('tr-TR').replace(/\s+/g, ''); }
  function sayi(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v).trim().replace(/\s/g, '');
    if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.'); // 1.234,56
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  function metin(v) { return v == null ? '' : String(v).trim(); }

  // Başlık satırını bul ve sütunları eşle
  function sutunBul(satirlar, gerekenler, adaylar) {
    var limit = Math.min(satirlar.length, 15);
    for (var i = 0; i < limit; i++) {
      var r = satirlar[i] || [];
      var h = r.map(norm);
      var eslesme = {};
      Object.keys(adaylar).forEach(function (anahtar) {
        var liste = adaylar[anahtar];
        for (var k = 0; k < liste.length; k++) {
          var idx = h.indexOf(liste[k]);
          if (idx >= 0) { eslesme[anahtar] = idx; break; }
        }
        if (eslesme[anahtar] == null) {
          // kısmi eşleşme (ör. "KAMPANYA TANIMI " / "KAMPANYA ADI")
          for (var k2 = 0; k2 < liste.length; k2++) {
            if (liste[k2].charAt(0) !== '~') continue;
            var parca = liste[k2].slice(1);
            for (var j = 0; j < h.length; j++) if (h[j] && h[j].indexOf(parca) >= 0) { eslesme[anahtar] = j; break; }
            if (eslesme[anahtar] != null) break;
          }
        }
      });
      if (gerekenler.every(function (g) { return eslesme[g] != null; })) return { baslikSatiri: i, sutun: eslesme };
    }
    return null;
  }

  var STOK_ADAYLARI = {
    barkod: ['BARKOD', 'BARKODNO', 'BARCODE'],
    kod: ['STOKKODU', 'URUNKODU', 'MODELKODU'],
    ad: ['STOKADI', 'URUNADI', 'STOKACIKLAMASI'],
    renk: ['RENKADI', 'RENK', 'RENKACIKLAMASI'],
    beden: ['BEDEN', 'BEDENADI'],
    fiili: ['FIILISTOKADET', 'FIILISTOK'],
    stokAdet: ['STOKADET', 'STOKMIKTARI', 'STOK', '~STOKADET'],
    transfer: ['TRANSFERSTOKADET', 'TRANSFERSTOK', '~TRANSFER'],
    fiyatKdvli: ['PERAKENDESATISFIYATIKDVLI', 'SATISFIYATIKDVLI', 'PERAKENDEFIYATKDVLI'],
    fiyatKdvsiz: ['PERAKENDESATISFIYATIKDVSIZ', 'SATISFIYATIKDVSIZ', 'PERAKENDESATISFIYATI', '~PERAKENDESATISFIYATI'],
    magaza: ['STOKYERIUNVAN', 'STOKYERIUNVANI', 'STOKYERI', 'MAGAZA', 'MAGAZAADI', 'DEPO'],
    cinsiyet: ['CINSIYETADI', 'CINSIYET'],
    klasman: ['KLASMANADI', 'KLASMAN']
  };
  var KAMPANYA_ADAYLARI = {
    kod: ['STOKKODU', 'URUNKODU', 'MODELKODU'],
    tanim: ['KAMPANYATANIMI', 'KAMPANYA', 'KAMPANYAADI', '~KAMPANYA']
  };

  function yuvarla2(n) { return Math.round(n * 100) / 100; }

  /* satirListeleri: [{ad: 'dosya.xlsx', satirlar: [[...],[...]]}, ...]
     secenekler: { kdv: 10, varsayilanMagaza: 'MAĞAZA' } */
  function stokOlustur(satirListeleri, secenekler) {
    secenekler = secenekler || {};
    var kdv = secenekler.kdv == null ? 10 : Number(secenekler.kdv);
    var magazalar = [], magazaIdx = {};
    var modeller = {};
    var urunIdx = {}; // barkod -> urunler indeksi
    var urunler = [];
    var uyarilar = [];
    var toplamSatir = 0;

    satirListeleri.forEach(function (dosya) {
      var bulunan = sutunBul(dosya.satirlar, ['barkod', 'kod'], STOK_ADAYLARI);
      if (!bulunan) throw new Error('"' + dosya.ad + '" dosyasında "Barkod" ve "Stok Kodu" sütunları bulunamadı.');
      var c = bulunan.sutun;
      if (c.fiili == null && c.stokAdet == null) throw new Error('"' + dosya.ad + '" dosyasında stok adedi sütunu ("Fiili Stok Adet" veya "Stok Adet") bulunamadı.');
      if (c.fiyatKdvli == null && c.fiyatKdvsiz == null) uyarilar.push('"' + dosya.ad + '": fiyat sütunu bulunamadı, fiyatlar boş görünecek.');
      var vMagaza = secenekler.varsayilanMagaza || dosya.ad.replace(/\.[^.]+$/, '').replace(/\s*STOK\s*$/i, '').trim() || 'MAĞAZA';

      for (var i = bulunan.baslikSatiri + 1; i < dosya.satirlar.length; i++) {
        var r = dosya.satirlar[i];
        if (!r) continue;
        var barkod = temizKod(r[c.barkod]);
        var kod = metin(r[c.kod]);
        if (!barkod && !kod) continue;
        if (!barkod) continue;
        toplamSatir++;
        var mAd = c.magaza != null ? metin(r[c.magaza]) : '';
        if (!mAd) mAd = vMagaza;
        if (magazaIdx[mAd] == null) { magazaIdx[mAd] = magazalar.length; magazalar.push(mAd); }
        var mi = magazaIdx[mAd];

        var fiili = c.fiili != null ? sayi(r[c.fiili]) : sayi(r[c.stokAdet]);
        var transfer = c.transfer != null ? sayi(r[c.transfer]) : 0;
        if (c.fiili == null && c.stokAdet != null && c.transfer != null) fiili = sayi(r[c.stokAdet]) - transfer;

        var kodK = kodAnahtari(kod);
        if (!modeller[kodK]) {
          var fiyat = null;
          if (c.fiyatKdvli != null && sayi(r[c.fiyatKdvli]) > 0) fiyat = yuvarla2(sayi(r[c.fiyatKdvli]));
          else if (c.fiyatKdvsiz != null && sayi(r[c.fiyatKdvsiz]) > 0) fiyat = yuvarla2(sayi(r[c.fiyatKdvsiz]) * (1 + kdv / 100));
          modeller[kodK] = [
            metin(r[c.ad]) || kod,
            fiyat,
            kod,
            c.cinsiyet != null ? metin(r[c.cinsiyet]) : '',
            c.klasman != null ? metin(r[c.klasman]) : ''
          ];
        }

        var ui = urunIdx[barkod];
        if (ui == null) {
          ui = urunIdx[barkod] = urunler.length;
          urunler.push([barkod, kodK, c.renk != null ? metin(r[c.renk]) : '', c.beden != null ? metin(r[c.beden]) : '', []]);
        }
        var stoklar = urunler[ui][4];
        var var_ = null;
        for (var s = 0; s < stoklar.length; s++) if (stoklar[s][0] === mi) { var_ = stoklar[s]; break; }
        if (var_) { var_[1] += fiili; var_[2] += transfer; }
        else stoklar.push([mi, fiili, transfer]);
      }
    });

    if (!urunler.length) throw new Error('Stok dosyasında okunabilir ürün satırı bulunamadı.');
    return {
      tarih: new Date().toISOString(),
      kaynak: satirListeleri.map(function (d) { return d.ad; }).join(', '),
      kdv: kdv,
      magazalar: magazalar,
      modeller: modeller,
      urunler: urunler,
      satirSayisi: toplamSatir,
      uyarilar: uyarilar
    };
  }

  function kampanyaOlustur(satirListeleri) {
    var liste = {}, adet = 0;
    satirListeleri.forEach(function (dosya) {
      var bulunan = sutunBul(dosya.satirlar, ['kod'], KAMPANYA_ADAYLARI);
      if (!bulunan) throw new Error('"' + dosya.ad + '" dosyasında "Stok Kodu" sütunu bulunamadı.');
      var c = bulunan.sutun;
      for (var i = bulunan.baslikSatiri + 1; i < dosya.satirlar.length; i++) {
        var r = dosya.satirlar[i];
        if (!r) continue;
        var kod = kodAnahtari(r[c.kod]);
        if (!kod) continue;
        var tanim = c.tanim != null ? metin(r[c.tanim]).replace(/\s+/g, ' ') : '';
        if (!tanim) tanim = '1 ALANA 1 BEDAVA';
        if (!liste[kod]) adet++;
        liste[kod] = tanim.toLocaleUpperCase('tr-TR');
      }
    });
    if (!adet) throw new Error('Kampanya dosyasında stok kodu bulunamadı.');
    return {
      tarih: new Date().toISOString(),
      kaynak: satirListeleri.map(function (d) { return d.ad; }).join(', '),
      liste: liste,
      adet: adet
    };
  }

  var api = { stokOlustur: stokOlustur, kampanyaOlustur: kampanyaOlustur, kodAnahtari: kodAnahtari, temizKod: temizKod, norm: norm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ArmaVeri = api;
})(typeof self !== 'undefined' ? self : this);
