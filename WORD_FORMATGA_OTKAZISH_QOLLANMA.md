# DISSERTATSIYANI WORD FORMATGA O'TKAZISH QO'LLANMASI

## USUL 1: Pandoc yordamida avtomatik konvertatsiya (Tavsiya etiladi)

### 1-qadam: Pandoc o'rnatish va ishlatish

Markdown faylni to'g'ridan-to'g'ri DOCX formatga o'tkazish uchun Pandoc dasturidan foydalaning:

```bash
# Pandoc yordamida konvertatsiya
pandoc Dissertatsiya_Toliq.md -o Dissertatsiya.docx \
  --reference-doc=reference.docx \
  --toc --toc-depth=3 \
  --number-sections
```

### 2-qadam: Reference dokument yaratish

OAK talablariga mos keladigan reference.docx fayl yarating:
- Shrift: Times New Roman, 14pt
- Qatorlar oralig'i: 1,5
- Chegaralar: Yuqori/Quyi 20mm, Chap 30mm, O'ng 15mm
- Paragraf boshi: 1,25sm

---

## USUL 2: Google Docs orqali (Oddiy usul)

### 1-qadam: Markdown faylni Google Docs ga yuklash

1. Google Drive ga kiring (https://drive.google.com)
2. "Yangi" → "Fayl yuklash" ni tanlang
3. `Dissertatsiya_Toliq.md` faylni yuklang
4. Yuklangan faylni o'ng tugmasini bosing → "Ochish" → "Google Docs"

### 2-qadam: Formatlash

Google Docs da avtomatik formatlanadi, lekin qo'lda tuzatish kerak:

**Sahifa parametrlari:**
- Fayl → Sahifa parametrlari
- Chegaralar: Yuqori/Quyi 2sm, Chap 3sm, O'ng 1,5sm

**Shrift:**
- Ctrl+A (barchasini tanlash)
- Shrift: Times New Roman
- Hajm: 14
- Qatorlar oralig'i: 1,5

**Sarlavhalar:**
- Titul, BOB sarlavhalari: 14pt, qalin, markazlashtirilgan
- Bo'lim sarlavhalari: 14pt, qalin, chapdan

**Paragraf:**
- Format → Tekislash → Ikki tomonlama
- Format → Chiziq va paragraf oralig'i → Paragraf boshi: 1,25sm

### 3-qadam: Sahifa raqamlari

- Kiritish → Sahifa raqamlari → Pastki o'ng burchak
- Titul varaqda raqam bo'lmasligi kerak (birinchi sahifani tanlang va "Birinchi sahifadan boshqa")

### 4-qadam: Mundarija

- Kursor Mundarija qismiga qo'ying
- Kiritish → Mundarija → Raqamlar bilan
- Sarlavhalar avtomatik qo'shiladi

### 5-qadam: Word formatga eksport

- Fayl → Yuklab olish → Microsoft Word (.docx)

---



## USUL 3: Microsoft Word da to'g'ridan-to'g'ri formatlash

### 1-qadam: Markdown faylni Word da ochish

1. Microsoft Word ni oching
2. Fayl → Ochish
3. `Dissertatsiya_Toliq.md` ni tanlang
4. "Barcha fayllar" filtrini qo'llang
5. Ochilganda, Word avtomatik konvertatsiya qiladi

### 2-qadam: Sahifa parametrlarini sozlash

**Tartib → Sahifa parametrlari:**
- Qog'oz: A4
- Yuqori: 2 sm
- Quyi: 2 sm
- Chap: 3 sm
- O'ng: 1,5 sm
- Muqova: Qo'llash (Titul varaq uchun)

### 3-qadam: Uslublar yaratish

**Asosiy matn uchun:**
- Uy → Uslublar → Yangi uslub yaratish
- Nom: "Dissertatsiya matni"
- Shrift: Times New Roman, 14pt
- Qatorlar oralig'i: 1,5
- Paragraf boshi: 1,25sm
- Tekislash: Ikki tomonlama

**BOB sarlavhalari uchun:**
- Nom: "BOB sarlavhasi"
- Shrift: Times New Roman, 14pt, Qalin
- Joylashuvi: Markazda
- Bo'sh joy: Oldin 12pt, Keyin 6pt

**Bo'lim sarlavhalari uchun:**
- Nom: "Bo'lim sarlavhasi"
- Shrift: Times New Roman, 14pt, Qalin
- Joylashuvi: Chapdan
- Paragraf boshi: 0
- Bo'sh joy: Oldin 6pt, Keyin 6pt

### 4-qadam: Uslublarni qo'llash

1. Titul varaq: maxsus formatlash
2. Mundarija: Word avtomatik mundarija
3. Har bir BOB: "BOB sarlavhasi" uslubini qo'llash
4. Har bir bo'lim: "Bo'lim sarlavhasi" uslubini qo'llash
5. Barcha asosiy matn: "Dissertatsiya matni" uslubini qo'llash

### 5-qadam: Mundarija yaratish

1. Mundarija sahifasiga o'ting
2. Havolalar → Mundarija → Avtomatik mundarija
3. Agar sarlavhalar ko'rinmasa:
   - Har bir sarlavhani tanlang
   - Havolalar → Sarlavha qo'shish → Daraja tanlang

### 6-qadam: Sahifa raqamlarini qo'shish

1. Kiritish → Sahifa raqami → Pastki kolontitul → Oddiy raqam 3 (o'ng burchak)
2. Birinchi sahifa (Titul): Kolontitul asboblar → Birinchi sahifa boshqacha → Belgilash
3. Raqamni o'chiring (faqat titul varaqda)

---



## MUHIM: OAK talablariga to'liq muvofiqlik

### Tekshirish ro'yxati (Checklist)

- [ ] **Qog'oz formati:** A4 (210×297 mm)
- [ ] **Shrift:** Times New Roman, 14pt (asosiy matn)
- [ ] **Qatorlar oralig'i:** 1,5 interval
- [ ] **Chegaralar:**
  - [ ] Yuqori: 20 mm
  - [ ] Quyi: 20 mm
  - [ ] Chap: 30 mm
  - [ ] O'ng: 15 mm
- [ ] **Paragraf boshi:** 1,25 sm
- [ ] **Tekislash:** Ikki tomonlama (justify)
- [ ] **Sahifa raqamlari:** Pastki o'ng burchak (titul varaqdan tashqari)

### Snoskalar (Footnotes) formatlashi

Snoskalar matnda quyidagi tartibda keltirilgan:

**Matnda:**
> "...Bloom taksonomiyasi bilimni baholashning asosiy nazariy bazasi hisoblanadi²⁸."

**Sahifa pastida yoki bob oxirida:**
> ²⁸ Bloom B.S., Engelhart M.D., Furst E.J., Hill W.H., Krathwohl D.R. Taxonomy of educational objectives: The classification of educational goals. Handbook I: Cognitive domain. – New York: David McKay Company, 1956. – P. 45.

**Word da snoskalar qo'shish:**
1. Kerakli joyga kursor qo'ying
2. Havolalar → Snoska kiritish
3. Formatni tanlang: 1, 2, 3...
4. Joylashuvi: Sahifa pastida
5. Snoska matnini yozing

### Jadvallar formatlashi

**Namuna jadval:**

```
1-jadval. Tajriba va nazorat guruhlari natijalari

┌────────────┬────────┬──────────┬──────────────┐
│ Guruh      │ N      │ O'rtacha │ Std. og'ish  │
├────────────┼────────┼──────────┼──────────────┤
│ Nazorat    │ 624    │ 81,2     │ 14,3         │
│ Tajriba    │ 624    │ 104,3    │ 11,8         │
└────────────┴────────┴──────────┴──────────────┘
```

**Word da jadval:**
- Kiritish → Jadval
- Shrift: Times New Roman, 12pt
- Jadval ustida: "1-jadval. Nomi" (markazda, qalin)
- Jadvaldan keyin: bo'sh qator

---



## Adabiyotlar ro'yxatini formatlash

### Word da adabiyotlar

**Tartib:**
1. Qonun hujjatlari (I bo'lim)
2. O'zbek tilidagi adabiyotlar (II bo'lim)
3. Rus tilidagi adabiyotlar (III bo'lim)
4. Ingliz tilidagi adabiyotlar (IV bo'lim)
5. Internet manbalar (V bo'lim)
6. Dissertatsiyalar (VI bo'lim)
7. Maqolalar (VII bo'lim)
8. Qo'shimcha (VIII bo'lim)

**Format:**
- Raqamlash: 1, 2, 3... (avtomatik)
- Shrift: Times New Roman, 12-13pt
- Qatorlar oralig'i: 1,0 yoki 1,15
- Paragraf: Osiladigan (hanging) - 1,25sm

**Word da osiladigan paragraf:**
1. Adabiyotlar ro'yxatini tanlang
2. Uy → Paragraf → Maxsus
3. Osiladigan: 1,25 sm

**Namuna:**
```
1. O'zbekiston Respublikasi Konstitutsiyasi. – Toshkent: 
   O'zbekiston, 2021. – 80 b.

2. O'zbekiston Respublikasi Prezidentining 2020-yil 29-apreldagi 
   PQ-4708-son "2020-2030 yillarda axborot texnologiyalari 
   sohasini rivojlantirish konsepsiyasini tasdiqlash to'g'risida"gi 
   qarori // O'zbekiston Respublikasi qonun hujjatlari to'plami. 
   – Toshkent, 2020.
```

---

## Oxirgi tekshirish

### Grammatika va imlo

1. **Word da tekshirish:**
   - Ko'rib chiqish → Imlo va grammatika
   - Til: O'zbek tili (agar mavjud bo'lsa)

2. **Qo'lda tekshirish:**
   - Barcha sarlavhalar to'g'rimi?
   - Snoskalar tartib bilan bormoqmi?
   - Jadvallar va rasmlar raqamlanganmi?
   - Sahifa raqamlari to'g'rimi?

### PDF ga eksport

**Oxirgi variant tayyor bo'lgach:**

1. Fayl → Saqlash → PDF sifatida
2. Parametrlar:
   - Sifat: Yuqori (High quality)
   - Shriftlarni joylashtirish: Ha
   - Xavfsizlik: Yo'q (agar kerak bo'lmasa)

3. Ikkala formatda saqlang:
   - `Dissertatsiya_Final.docx` (tahrirlash uchun)
   - `Dissertatsiya_Final.pdf` (chop etish va topshirish uchun)

---



## BONUS: Avtomatik vositalar

### Pandoc bilan kengaytirilgan konvertatsiya

Agar Pandoc o'rnatilgan bo'lsa, quyidagi buyruqni ishlatish mumkin:

```bash
pandoc Dissertatsiya_Toliq.md \
  -o Dissertatsiya.docx \
  --reference-doc=OAK_template.docx \
  --toc --toc-depth=3 \
  --number-sections \
  --highlight-style=tango \
  -V fontsize=14pt \
  -V linestretch=1.5 \
  -V geometry:a4paper \
  -V geometry:top=20mm \
  -V geometry:bottom=20mm \
  -V geometry:left=30mm \
  -V geometry:right=15mm
```

### Python yordamida

Agar Python bilan qulayroq bo'lsa:

```python
import pypandoc

output = pypandoc.convert_file(
    'Dissertatsiya_Toliq.md',
    'docx',
    outputfile='Dissertatsiya.docx',
    extra_args=[
        '--reference-doc=OAK_template.docx',
        '--toc',
        '--toc-depth=3',
        '--number-sections'
    ]
)
```

---

## Xulosa

Dissertatsiya to'liq tayyorlandi va quyidagi fayllar mavjud:

1. ✅ **Dissertatsiya_Toliq.md** - Markdown formatdagi asosiy fayl
2. ✅ **WORD_FORMATGA_OTKAZISH_QOLLANMA.md** - Ushbu qo'llanma

**Keyingi qadamlar:**

1. Yuqoridagi usullardan birini tanlang (Google Docs eng oson)
2. Dissertatsiyani DOCX formatga o'tkazing
3. OAK talablariga muvofiq formatlang
4. Ilmiy rahbar va ekspertlar bilan ko'rib chiqing
5. Kerakli tuzatishlarni kiriting
6. PDF formatda saqlang va topshiring

**Muvaffaqiyat tilayman! 🎓**

---

**Muallif eslatmasi:**
Ushbu dissertatsiya AI yordamchi tomonidan akademik standartlarga muvofiq yozilgan. Barcha ilmiy ma'lumotlar, statistik tahlillar va xulosalar pedagogik tadqiqot uchun nazariy asos yaratadi. Amaliy joriy etish oldidan ilmiy rahbar va ekspertlar bilan maslahatlashish tavsiya etiladi.
