# ƏVƏZETMƏ HESABATI — KONSOLİDASİYA TOOL-U (Texniki Spesifikasiya)

> **Bu sənədin təyinatı:** Claude Code bu spesifikasiyaya əsasən eyni GitHub reposunda tool yaratmalıdır.
> Tool 5 mənbə Excel faylını qəbul edir və aşağıda təsvir edilən strukturda **bir yekun .xlsx hesabat** qaytarır.
> Hesabat auditə təqdim edilir və məhkəmə predmeti ola bilər — **dəqiqlik kritikdir**. Bütün yoxlama
> mərhələləri (bölmə 12) icra edilmədən nəticə qaytarılmamalıdır.

---

## 1. Ümumi axın

```
5 giriş faylı
   │
   ├── Mərhələ 1: Normalizasiya (vahid sxemə çevirmə + filtr + xərc qaydaları)
   ├── Mərhələ 2: Duplikat aşkarlanması + üst-üstə düşmə / iç-içə yoxlamaları + gün/saat hesablanması
   ├── Mərhələ 3: Yekun workbook-un qurulması (bütün vərəqlər, düsturlar, dizayn)
   ├── Mərhələ 4: Backup vərəqlərinin əlavə edilməsi
   └── Mərhələ 5: Düsturların yenidən hesablanması (LibreOffice) + müstəqil cross-check
   │
   └──► Əvəzetmə_Hesabatı.xlsx
```

**Texnologiya:** Python 3, `pandas`, `openpyxl`, `xlrd>=2.0.1` (.xls üçün), LibreOffice headless
(düstur recalc üçün). `pip install ... --break-system-packages` lazım ola bilər.

---

## 2. Konfiqurasiya sabitləri

| Sabit | Dəyər | İzah |
|---|---|---|
| `FILTER_START` | `2026-01-01` | Yalnız başlama tarixi bu tarixdən (daxil) sonrakı sətirlər təmiz dataya düşür. Əvvəlkilər Xəta Loquna. |
| `CAP_DATE` | `2026-06-30` | Bitmə tarixi boş və ya `2999-12-31` olan (davam edən) sətirlərdə hesablama sərhədi. |
| `WIN_START` | `2026-01-03` | "Əməkdaş üzrə cəm" pəncərəsinin başlanğıcı. |
| `WIN_END` | `2026-07-02` | Pəncərənin sonu. WIN_START–WIN_END daxil olmaqla **dəqiq 181 gün** olmalıdır. |
| `WORK_HOURS_PER_DAY` | `8` | Günlük sətirlərdə saat = gün × 8 fərziyyəsi. |
| Həftəsonu | Şənbə + Bazar | Azərbaycan təqvimi (`weekday() >= 5`). |

> Qeyd: bu tarixlər hazırkı data çıxarışına bağlıdır. Tool-da parametrləşdirilə bilər, amma
> defolt dəyərlər yuxarıdakı kimi olmalıdır. `WIN_END` datanın çıxarılma tarixidir,
> `WIN_START = WIN_END − 180 gün` (181 günlük pəncərə).

---

## 3. Giriş faylları, sxemləri və mənbə adları

Hesabatda hər mənbə aşağıdakı **yeni adla** göstərilir:

| Fayl | Sheet | Hesabatdakı ad |
|---|---|---|
| `İnzibatciliq prosesi 2.xlsx` | `Sheet1` | **İnzibatçı prosesi** |
| `Vacation, Time off replacement.xls` | `Vacation replacement` | **İşburaxma hallarının əvəzedilməsi** |
| `Vacation, Time off replacement.xls` | `time off replacement` | **Saatlıq icazə** |
| `Filial baza .xlsx` | `Sheet1` | **Müvəqqəti keçid** |
| `Rəsmi əvəzetmə(position info).xlsx` | `Select d_position_info` | **Rəsmi əvəzetmə (həvalə)** |
| `Vəzifə-Maaş.xlsx` | 6 aylıq sheet | *(yalnız lookup üçün; hesabata mənbə kimi girmir)* |

### 3.1. İnzibatçı prosesi (`İnzibatciliq prosesi 2.xlsx`)
Sütunlar: `Personal Kod`, `Evezeden Emekdas`, `Evezedilen Emekdas` (praktikada boş),
`Baslama Tarixi`, `Bitme Tarixi`, `Period` (`{1}Müvəqqəti` / `{0}Daimi`), `Teyin Olundugu Yer`, `Növ`.

Xüsusiyyətlər:
- `Bitme Tarixi` bir çox sətirdə boşdur (davam edən / daimi təyinatlar).
- Faylda vəzifə/struktur məlumatı **yoxdur** → Vəzifə-Maaş lookup-u (bölmə 6).
- `Növ` dəyərləri: `1-ci xəzinə inzibatçısı - kassa əməliyyatlarına məsul şəxs`,
  `2-ci xəzinə inzibatçısı`, `3-cü xəzinə inzibatçısı`,
  `3-cü xəzinə inzibatçısı və Xidmət bölməsinin rəhbəri səlahiyyətlərini icra edən şəxs`.

### 3.2. İşburaxma hallarının əvəzedilməsi (`Vacation replacement` sheet-i)
Sütunlar (pandas oxuyanda təkrarlananlar `.1` şəkilçisi alır):
`EVEZ_EDEN_PERSONAL_KOD`, `EVEZ_EDEN_FULL_NAME`, `STRUKTUR`, `SHOBE`, `BOLME`, `VEZIFE`,
`HEVALE_STRUKTUR`, `HEVALE_VEZIFE`, `MEZUNIYYETIN_BASHLAMA_TARIXI`, `MEZUNIYYETIN_BITME_TARIXI`,
`ORDER_NUMBER`, `ORDER_DATE`, `EVEZ_ETDIYI_SEXS_PERSONAL_KOD`, `EVEZ_ETDIYI_FULL_NAME`,
`STRUKTUR.1`, `SHOBE.1`, `BOLME.1`, `VEZIFE.1`, `HEVALE_STRUKTUR.1`, `HEVALE_VEZIFE.1`, `BO/Filial`.

Xüsusiyyətlər:
- **Tarixlər `int` formatında `YYYYMMDD`** (məs. `20260213`) → `datetime.strptime(str(int(x)), '%Y%m%d')`.
- `EVEZ_EDEN_*` = əvəz edən; `EVEZ_ETDIYI_*` = məzuniyyətə çıxan (əvəz edilən) şəxs.
- Bitmə tarixi 2027-dən sonra olan sətirlər (məs. `20290305`) şübhəli daxiletmə xətası sayılır:
  bitmə boş kimi qəbul et, `CAP_DATE`-ə qədər hesabla, Qeyd + Xəta Loqu yazısı.
- `BO/Filial`: `BO` → `Baş ofis`, `Filial` → `Filial`.

### 3.3. Saatlıq icazə (`time off replacement` sheet-i)
Sütunlar: `Source_type`, `REPLACING_EMPLOYEE`, `Replacing_full_name`, `REPLACING_STRUCTURE_NAME`,
`REPLACING_SECTION_NAME`, `REPLACING_SUB_SECTION_NAME`, `REPLACING_POSITION`, `HR_CODE`,
`STRUCTURE_NAME`, `SECTION_NAME`, `SUB_SECTION_NAME`, `POSITION`, `Full_Name`, `TIME_OFF_DATE`,
`TIME_FROM_TO`, `Acting`, `Baş ofis/filial`.

Xüsusiyyətlər:
- Yalnız **başlama tarixi + saat aralığı** var. `TIME_FROM_TO` formatı: `"HH:MM HH:MM"`
  (regex: `^\s*(\d{1,2}):(\d{2})\s+(\d{1,2}):(\d{2})\s*$`).
- Başlama = `TIME_OFF_DATE + birinci saat`; Bitmə = `TIME_OFF_DATE + ikinci saat` (**eyni tarix, fərqli saat**).
- Bitmə saatı ≤ başlama saatı olan sətirlər → təmiz dataya salınmır, Xəta Loquna (`Saat aralığı məntiqsizdir`).
- `Acting` dolu olduqda Qeydlərə `Acting vəzifə: <dəyər>` əlavə et.
- `REPLACING_*` = əvəz edən; `HR_CODE/Full_Name/STRUCTURE_NAME/POSITION` = əvəz edilən.

### 3.4. Müvəqqəti keçid (`Filial baza .xlsx`)
Sütunlar: `Ezamın Növü`, `Personal kod`, `Soyad Ad Ata adı`, `Departament`, `Şöbə`, `Bölmə`,
`Vəzifə`, `Həva olunmuş vəzifələ`, `Ezam olunduğu departament/filial`, `Ezamın Başlama tarixi`,
`Ezamın Bitmə tarixi`, `Əvəzedici şəxs`, (boş sütun), `Overtime`, `Əmrin nömrəsi`, `Əmrin tarixi`.

Xüsusiyyətlər:
- `Ezamın Növü` dəyərləri: `Müvəqqəti keçid`, `Müvəqqəti keçidin dayandırılması`,
  `Müvəqqəti keçidin ləğvi`, `Təcrübə`, `Müdir əvəzi`, `kassa inzibatçısı əvəzi`.
- **`ləğvi` və `dayandırılması` sətirləri əvəzetmə deyil** → təmiz dataya salınmır, Xəta Loquna yazılır
  (`Əmr növü: Müvəqqəti keçidin ləğvi` / `... dayandırılması`). Bundan əlavə, həmin şəxsin eyni/kəsişən
  tarixli saxlanılan keçid sətrinə Qeyd əlavə edilir:
  `DİQQƏT: bu keçid üzrə '<növ>' əmri mövcuddur (mənbə sətri N, əmr X) – faktiki müddət fərqli ola bilər`.
- Sütun başlıqlarında artıq boşluqlar ola bilər → `str.strip()` tətbiq et.
- `Departament` bəzən `#` — o halda struktur kimi `Şöbə` götür.

### 3.5. Rəsmi əvəzetmə (həvalə) (`Rəsmi əvəzetmə(position info).xlsx`)
Sütunlar: `SOURCE_TYPE`, `LOAD_DATE`, `HR_CODE`, `Baş ofis/Filial`, `STRUCTURE_NAME`, `SECTION_NAME`,
`SUB_SECTION_NAME`, `POSITION`, `FULL_NAME`, `BEGIN_DATE`, `END_DATE`, `ACTING_STRUCTURE`, `ACTING`,
`ACTING_BEGIN_DATE`, `ACTING_END_DATE`.

Xüsusiyyətlər:
- **Qırmızı fon rəngi (`FFFF0000`) olan sətirlər xətalıdır** → təmiz dataya salınmır, Xəta Loquna
  (`Qırmızı işarələnmiş sətir`). Aşkarlanma: openpyxl ilə hər sətrin xanalarında
  `cell.fill.fgColor.rgb` içində `FF0000` axtar. Backup vərəqində bu sətirlərin qırmızı fonu saxlanılır.
- Əvəzetmə dövrü = `ACTING_BEGIN_DATE` – `ACTING_END_DATE` (BEGIN/END_DATE deyil!).
- `ACTING_END_DATE = 2999-12-31` → açıq müddət: bitmə boş kimi, `CAP_DATE`-ə qədər hesabla,
  Qeyd: `Bitmə tarixi 31.12.2999 (açıq müddət/davam edir) – gün sayı 30.06.2026-a qədər hesablanıb`.
- `ACTING_BEGIN_DATE < FILTER_START` (məs. 2001, 2025) → filtrdən kənar, Xəta Loquna.
- Əvəz edilən vəzifə sütununa `ACTING` dəyəri + `" (həvalə vəzifəsi)"` şəkilçisi yazılır.

### 3.6. Vəzifə-Maaş (`Vəzifə-Maaş.xlsx`) — yalnız lookup
Sheet-lər: `Yanvar 2026` … `Iyun 2026`. Sütunlar (başlıqlarda `\n` var, normallaşdır):
`Personal kod`, `Soyadı, adı və ata adı`, `Departament`, `Şöbə`, `Bölmə`, `Vəzifə`.

---

## 4. Vahid sxem (normalizasiya)

Hər mənbə sətri aşağıdakı sahələrə çevrilir:

| Sahə | İzah |
|---|---|
| `Mənbə` | Yeni mənbə adı (bölmə 3 cədvəli) |
| `Mənbə sətri` | Orijinal fayldakı **Excel sətir nömrəsi** (başlıq = 1 → data `index + 2`). Audit izlənəbilirliyi üçün mütləqdir. |
| `Əvəz edən kod / ad / struktur / vəzifə` | |
| `Əvəz edilən kod / ad / struktur / vəzifə` | Mənbədə yoxdursa boş |
| `Struktur / Təyin yeri` | Təyinat yeri (İnzibatçı: `Teyin Olundugu Yer`; Müvəqqəti keçid: `Ezam olunduğu departament/filial`; İşburaxma: `STRUKTUR.1`; Saatlıq: `STRUCTURE_NAME`; Rəsmi: `ACTING_STRUCTURE`) |
| `Baş ofis/Filial` | Bölmə 5.2 |
| `Əvəzetmə növü` | İnzibatçı: `<Növ> (Müvəqqəti/Daimi)`; İşburaxma: `Məzuniyyət/xəstəlik əvəzetməsi`; Saatlıq: `İşdən icazə (saatlıq)`; Müvəqqəti keçid: `Ezamın Növü`; Rəsmi: `Rəsmi əvəzetmə / acting: <ACTING>` |
| `Başlama`, `Bitmə` | datetime; açıq bitmə = `NaT` |
| `Xərc yaradır?` | `Bəli` / `Xeyr` (bölmə 5.1) |
| `Qeydlər` | list — bütün flaqlar bura yığılır, çıxışda `" || "` ilə birləşdirilir |
| `day_based` | `True` (günlük) / `False` (saatlıq) |

Mənbə üzrə xüsusi hallar:
- **İnzibatçı prosesi:** `Əvəz edən struktur/vəzifə` = Vəzifə-Maaş lookup nəticəsi +
  `" (Vəzifə-Maaş, <ay adı>)"` şəkilçisi.
- **Müvəqqəti keçid:** struktur = `Departament` (`#` olduqda `Şöbə`).
- Boş `Başlama` → sətir salınmır, Xəta Loquna. `Başlama < FILTER_START` → salınmır,
  Xəta Loquna (`Filtrdən kənar (2026-dan əvvəl)`).
- Boş `Bitmə` (günlük mənbələrdə): sətir **saxlanılır**, Qeyd: Daimi üçün
  `Daimi təyinat – bitmə tarixi yoxdur; gün sayı 30.06.2026-a qədər hesablanıb`,
  müvəqqəti üçün `Bitmə tarixi boşdur (müvəqqəti) – ...` + Xəta Loqu yazısı (`Bitmə tarixi boşdur`).

---

## 5. Biznes qaydaları

### 5.1. Xərc qaydaları (istifadəçi tərəfindən təsdiqlənib — DƏYİŞDİRMƏ!)

| Mənbə | Qayda |
|---|---|
| **İnzibatçı prosesi** | `Növ ∈ {"2-ci xəzinə inzibatçısı", "3-cü xəzinə inzibatçısı və Xidmət bölməsinin rəhbəri səlahiyyətlərini icra edən şəxs"}` → **Bəli**. `1-ci xəzinə inzibatçısı - kassa əməliyyatlarına məsul şəxs` və **TƏK** `3-cü xəzinə inzibatçısı` → **Xeyr**. |
| **İşburaxma** və **Saatlıq icazə** | `Baş ofis/Filial == Filial` **VƏ** həm əvəz edənin, həm əvəz edilənin vəzifəsi (kiçik hərflə, strip) **tam olaraq** `müdir` və ya `bölmə rəhbəri` → **Xeyr**. Qalan bütün hallar (Baş ofis daxil) → **Bəli**. Diqqət: substring yox, **tam bərabərlik** — `Şöbə müdiri` uyğun gəlmir. |
| **Müvəqqəti keçid** | `Ezam olunduğu departament/filial` mətnində (case-insensitive) `3-cü xəzinə inzibatçısı və Xidmət bölməsinin rəhbəri` **VƏ YA** `filial müdiri vəzifəsini icra edən` keçirsə → **Bəli**; qalanları (tək `(3-cü xəzinə inzibatçısı)` və `(1-ci xəzinə ...)` daxil) → **Xeyr**. *Yoxlanılmış fakt: hazırkı faylda bu qaydaya 174 sətir (32 unikal şəxs) uyğun gəlir.* |
| **Rəsmi əvəzetmə (həvalə)** | Qırmızı sətirlər istisna edildikdən sonra **bütün sətirlər → Bəli**. |
| **Duplikat konflikti** | Fayllar arası duplikat qrupunda mənbələr fərqli dəyər verirsə → **konservativ prinsip: qrupda hər hansı mənbə "Bəli" deyirsə, yekun "Bəli"**. Qeydlərə izah + Xəta Loquna `Xərc qiymətləndirməsi konflikti`. |

### 5.2. Baş ofis / Filial təsnifatı
- Mənbədə hazır sütun varsa (İşburaxma `BO/Filial`, Saatlıq `Baş ofis/filial`, Rəsmi `Baş ofis/Filial`) → ondan istifadə et.
- Yoxdursa (İnzibatçı, Müvəqqəti keçid) mətn əsaslı: struktur adında (kiçik hərflə)
  `departament | sahəsi | mərkəzi | mühasibatlıq | ofis | rəhbərlik | şura` keçirsə → `Baş ofis`,
  əks halda → `Filial` (filial, şöbə, notariat, mall və s. hamısı filial şəbəkəsidir;
  `Gəncə`, `Sahil` kimi tək adlar da filialdır).

### 5.3. Vəzifə-Maaş lookup (yalnız İnzibatçı prosesi üçün)
1. Başlama tarixinin ayına uyğun sheet (`Yanvar 2026` … `Iyun 2026`). İyuldan sonrakı başlama → `Iyun 2026`.
2. `Personal kod` üzrə axtar; həmin ayda tapılmasa digər ayları **İyundan geriyə** yoxla.
3. Nəticə: `Vəzifə` və `Departament` + `" (Vəzifə-Maaş, <sheet adı>)"` şəkilçisi.

---

## 6. Duplikat aşkarlanması (Mərhələ 2)

Açar normallaşdırması: `kod = str(int(float(x)))` (`110020.0` → `"110020"`).

### 6.1. Fayl daxili tam duplikat
Açar: `(Mənbə, kod, başlama_datetime, bitmə_datetime)`. İkinci və sonrakılar çıxarılır →
Duplikatlar vərəqinə (`Fayl daxili tam duplikat`).

### 6.2. Fayllar arası duplikat
Açar: `(kod, başlama.date(), bitmə.date())` — **tarix səviyyəsində** (saatlar fərqli ola bilər,
məs. İnzibatçı 10:00–19:00, Müvəqqəti keçid 09:00–18:00 — bunlar eyni sorğudur).
Bir açarda 2+ **fərqli mənbə** varsa → Request sisteminin fərqli tool-larından yaradılmış eyni sorğu.

- **Saxlama prioriteti:** `İnzibatçı prosesi (0) → İşburaxma (1) → Saatlıq icazə (2) → Müvəqqəti keçid (3) → Rəsmi əvəzetmə (4)`; bərabərlikdə kiçik sətir nömrəsi.
- Saxlanılan sətrin Qeydlərinə: `Fayllar arası duplikat: eyni əvəzetmə <mənbə (sətir N)>, ... mənbələrində də mövcuddur (Request sisteminin fərqli tool-larından yaradılıb); yalnız bu sətir saxlanılıb`.
- Xərc konflikti varsa → 5.1-dəki konservativ qayda + Qeyd + Xəta Loqu.
- Çıxarılanlar → Duplikatlar vərəqinə (`Fayllar arası duplikat`), saxlanılan mənbə göstərilməklə.

---

## 7. Xəta yoxlamaları (Mərhələ 2)

### 7.1. Üst-üstə düşən əvəzetmə (eyni mənbə daxilində, günlük sətirlər)
Eyni `(Mənbə, kod)` üzrə iki intervalın kəsişməsi (`s1 ≤ e2 && s2 ≤ e1`, açıq bitmə → `max(CAP_DATE, start)`):
- Hər iki sətrin Qeydlərinə `XƏTA – üst-üstə düşmə: ...` (qarşı sətrin nömrəsi və dövrü ilə).
- Xəta Loquna: `Üst-üstə düşən əvəzetmə` — `Bir şəxs eyni dövrdə iki əvəzetmədə ola bilməz.`
- Sətirlər Təmiz Datada **qalır** (flaqla), amma Yekun vərəqə **düşmür**.

### 7.2. İç-içə əvəzetmə (yalnız İşburaxma mənbəsi)
Sətir A-da əvəz edən şəxs X (`EVEZ_EDEN`), dövrü `[s1,e1]`. X başqa sətirdə **əvəz edilən**
(`EVEZ_ETDIYI`) kimi çıxırsa və o sətrin dövrü `[s2,e2]` `[s1,e1]` ilə kəsişirsə:
- A-nın Qeydlərinə: `İÇ-İÇƏ HAL: əvəz edən şəxs özü <s2–e2> dövründə <Y (kod)> tərəfindən əvəz edilib (sətir N); faktiki əvəzetmə müddəti sənəddəki tam müddətdən azdır`.
- Qarşı sətrə də simmetrik qeyd. Xəta Loquna `İç-içə əvəzetmə`. Zəncirvari (2+ adam) hallar avtomatik əhatə olunur.
- Bu sətirlər Yekun vərəqdə **qalır** (qeydlə) — prosedur etibarlıdır, yalnız məlumat xarakterlidir.

### 7.3. Tarix məntiqsizliyi
Günlük sətirdə real (boş olmayan) `Bitmə < Başlama` → sətir təmiz dataya salınmır,
Xəta Loquna `Tarix məntiqsizliyi (çıxarıldı)`.

### 7.4. Açıq bitmə + gec başlama
Bitmə boş və `Başlama > CAP_DATE` (məs. 06.07.2026) → hesablama bitməsi = başlama günü (1 gün),
Qeyd: `Başlama tarixi hesablama sərhədindən (30.06.2026) sonradır və bitmə açıqdır – gün sayı minimum 1 gün kimi hesablanıb`.
**Mənfi gün sayı heç bir halda yaranmamalıdır.**

---

## 8. Gün / saat hesablanması

Hesablamalar çıxış faylında **Excel düsturları** ilə aparılır (hardcode yox!). Hər data sətrində:

- `Hesablama bitməsi` (R sütunu) = real bitmə; boşdursa `max(CAP_DATE, başlama.normalize())`.
- `Gün sayı` (S): `=IF(O{r}="Saatlıq",ROUND((R{r}-P{r})*24/8,2),INT(R{r})-INT(P{r})+1)`
- `Saat sayı` (T): `=IF(O{r}="Saatlıq",ROUND((R{r}-P{r})*24,2),(INT(R{r})-INT(P{r})+1)*8)`

Yəni: günlük → təqvim günü, **hər iki tarix daxil** (14.01–27.01 = 14 gün); saat = gün × 8.
Saatlıq → saat = bitmə − başlama (saatla, 2 onluq); gün = saat / 8.

---

## 9. Çıxış faylının strukturu

Vərəq sırası:
1. **Ümumi Baxış**
2. **Xərc yaradan (yekun)**
3. **Təmiz Data**
4. **Əməkdaş üzrə cəm**
5. **Xəta Loqu**
6. **Duplikatlar**
7. **Statistika**
8. Backup vərəqləri (bölmə 10)

### 9.1. Ortaq data sütun sxemi (Təmiz Data və Yekun)

| Hərf | Sütun |
|---|---|
| A | № |
| B | Mənbə |
| C | Mənbə sətri |
| D | Əvəz edən kod |
| E | Əvəz edən (S.A.A.) |
| F | Əvəz edən struktur |
| G | Əvəz edən vəzifə |
| H | Əvəz edilən kod |
| I | Əvəz edilən (S.A.A.) |
| J | Əvəz edilən struktur |
| K | Əvəz edilən vəzifə |
| L | Təyin yeri |
| M | Baş ofis/Filial |
| N | Əvəzetmə növü |
| O | Tip (`Günlük` / `Saatlıq`) |
| P | Başlama |
| Q | Bitmə (açıq → boş) |
| R | Hesablama bitməsi |
| S | Gün sayı (düstur) |
| T | Saat sayı (düstur) |
| U | Xərc yaradır? *(yalnız Təmiz Datada; Yekunda bu sütun yoxdur — hamısı Bəli)* |
| V | Qeydlər *(Yekunda U)* |

Format: tarix sütunları `DD.MM.YYYY` (saatlıq sətirlərdə P,Q üçün `DD.MM.YYYY HH:MM`);
Excel Table (`TableStyleMedium16`, zolaqlı); başlıq tünd göy `#203864` fonda ağ qalın Arial 10;
data Arial 10, nazik `#D6DCE4` haşiyə; freeze panes `F2` (yekunda `F5`);
Qeydlər xanası: `XƏTA` keçərsə qırmızımtıl `#FADBD8`, `DİQQƏT/İÇ-İÇƏ/duplikat/Birləşdirilib` keçərsə sarımtıl `#FFF2CC` fon.
Table adları: `TemizData`, `XercYekun`, `EmekdasCem`, `XetaLoqu`, `DuplikatlarT`.

### 9.2. Xərc yaradan (yekun) — ƏSAS TƏLƏB OLUNAN VƏRƏQ
Başlıqda (A1–A2) izah mətnləri, cədvəl 4-cü sətirdən. Tərkib qaydaları:

1. Yalnız `Xərc yaradır? == Bəli` sətirlər (duplikatlar artıq çıxarılıb).
2. **İnzibatçı prosesi — davamlı sətirlərin birləşdirilməsi:**
   - Qruplaşdırma açarı: `(kod, Təyin yeri, Əvəzetmə növü)`.
   - Qrup daxilində başlamaya görə sırala; sətirlər epizodlara birləşdirilir:
     növbəti sətrin başlaması cari epizodun effektiv bitməsindən (açıq → `max(CAP_DATE, start)`)
     **+1 gün** və ya daha əvvəldirsə (kəsişir/bitişikdir) → eyni epizoda qoşulur.
   - Epizodun `Başlama` = qrupdakı **birinci başlama**; `Bitmə` = üzvlər arasında **son bitmə tarixi**;
     üzvlərdən hər hansı biri açıq bitməlidirsə → epizod açıqdır: `Bitmə` boş, hesablama `CAP_DATE = 30.06.2026`,
     Qeyd: `Bitmə tarixi açıqdır – 30.06.2026 kimi qəbul edilib`.
   - 2+ sətir birləşibsə `Mənbə sətri` = vergüllə bütün sətir nömrələri; Qeyd:
     `Birləşdirilib: davamlı gedən N mənbə sətri (sətirlər: ...); birinci başlama və son bitmə tarixi götürülüb`.
   - Birləşmədən **sonra** eyni şəxsin epizodları hələ də kəsişirsə (fərqli yer/növ üzrə) →
     hər iki epizod yekundan çıxarılır, Xəta Loquna `Yekun vərəqdən istisna: kəsişən epizodlar`.
3. **Digər mənbələr üzrə istisnalar:** Qeydlərində `XƏTA` (üst-üstə düşmə) və ya
   `ləğvi`/`dayandırılması` (Müvəqqəti keçid DİQQƏT qeydi) olan sətirlər yekuna düşmür
   (Təmiz Datada və Xəta Loqunda qalırlar).
4. Nəticə vərəqdə `(kod, Başlama, Bitmə)` üzrə **sıfır duplikat**, **sıfır XƏTA qeydi**,
   **sıfır boş/mənfi gün** olmalıdır — bunlar avtomatik test edilir (bölmə 12).

### 9.3. Əməkdaş üzrə cəm
Başlıq bloku: dövr xanaları (`B2`=WIN_START, `D2`=WIN_END, sarı fon, mavi şrift — dəyişdirilə bilən input)
+ izah mətni. Cədvəl 5-ci sətirdən, hər unikal `Əvəz edən kod` üçün bir sətir (ada görə sıralı):

| Sütun | Mənbə |
|---|---|
| A Əvəz edən kod, B Ad | |
| C Əvəzetmə sayı | `=COUNTIFS('Təmiz Data'!$D$2:$D$N, $A{r}, 'Təmiz Data'!$P$2:$P$N, ">="&$B$2, 'Təmiz Data'!$P$2:$P$N, "<"&$D$2+1)` |
| D Cəmi saat | eyni kriteriyalarla `SUMIFS` → T sütunu |
| E Sətirlərin cəm günü | eyni kriteriyalarla `SUMIFS` → S sütunu (kəsişmələr daxil ola bilər) |
| F Unikal təqvim günü | **Python-la hesablanır:** şəxsin bütün təmiz sətirlərinin tarix aralıqlarının **birləşməsi (union)** — hər aralıq `[başlama.date, hesablama_bitməsi.date]` (saatlıq → tək gün), `[WIN_START, WIN_END]` pəncərəsinə kəsilir, günlər set-də toplanır. **Maksimum 181 ola bilər.** |
| G — iş günü | F-dəki günlərdən `weekday() < 5` olanlar |
| H — şənbə/bazar | `weekday() >= 5` olanlar. **İnvariant: G + H = F.** |
| I Kəsişmə var? | `=IF(E{r}>F{r},"Bəli","")` |

### 9.4. Xəta Loqu
Sütunlar: `№`, `Mənbə`, `Mənbə sətri (Excel)`, `Xəta növü`, `Təfərrüat`, `Sətir məlumatı`
(orijinal sətrin `sütun=dəyər | ...` formatında dump-u). Mənbə + sətirə görə sıralı, Excel Table.
Xəta növləri kataloqu: `Filtrdən kənar (2026-dan əvvəl)`, `Bitmə tarixi boşdur`, `Şübhəli bitmə tarixi`,
`Saat aralığı oxunmadı`, `Saat aralığı məntiqsizdir`, `Tarix məntiqsizliyi (çıxarıldı)`,
`Üst-üstə düşən əvəzetmə`, `İç-içə əvəzetmə`, `Xərc qiymətləndirməsi konflikti`,
`Qırmızı işarələnmiş sətir`, `Əmr növü: Müvəqqəti keçidin ləğvi`, `Əmr növü: Müvəqqəti keçidin dayandırılması`,
`Yekun vərəqdən istisna: kəsişən epizodlar`, `Başlama tarixi boşdur`, `Başlama tarixi oxunmadı`, `Acting başlama tarixi boşdur`.

### 9.5. Duplikatlar
Yuxarıda mini-statistika (COUNTIF düsturları ilə: fayl daxili / fayllar arası / cəmi), 7-ci sətirdən cədvəl:
`Duplikat növü`, `Əvəz edən kod`, `Əvəz edən`, `Başlama`, `Bitmə`, `Saxlanılan mənbə (təmiz datada)`,
`Çıxarılan mənbə`, `Qeyd`.

### 9.6. Statistika
Yuxarıda 5 göstərici kartı (etiket `#2E75B6` fonda ağ, dəyər `#DEEAF6` fonda böyük qalın rəqəm):
`Yekun xərc sətri`, `Yekun cəmi gün`, `Yekun cəmi saat (saatlıq)`, `Təmiz sətir (cəmi)`, `Unikal əməkdaş`.
Sonra bölmələr (hamısı **düsturla**, digər vərəqlərə istinadən): Yekun vərəq cəmləri + istisna sayları;
Təmiz Data xərc bölgüsü; mənbə üzrə (sətir sayı | xərc yaradan); Baş ofis/Filial; Duplikatlar;
Xəta Loqu kateqoriyaları; sonda 9 bəndlik **Metodologiya** mətni (bu spesifikasiyadakı qaydaların xülasəsi).

### 9.7. Ümumi Baxış
Titul vərəqi (gridlines söndürülüb): başlıq, hazırlanma tarixi, filtr; mənbə fayl → hesabat adı xəritəsi;
vərəqlərin təyinatı; xərc qaydalarının xülasəsi (5.1 cədvəli).

---

## 10. Backup vərəqləri

Mənbə datanın **dəyişdirilməmiş** surətləri, hər vərəqin A1-də kursiv boz izah sətri, sonra başlıq (boz fon) + data:

| Vərəq adı | Mənbə |
|---|---|
| `BK_İnzibatçı prosesi` | İnzibatciliq prosesi 2.xlsx |
| `BK_İşburaxma əvəzedilməsi` | Vacation replacement sheet |
| `BK_Saatlıq icazə` | time off replacement sheet |
| `BK_SQL Statement` | həmin fayldakı texniki vərəq |
| `BK_Sheet2 (işburaxma faylı)` | həmin fayldakı Sheet2 (boş deyilsə) |
| `BK_Müvəqqəti keçid` | Filial baza .xlsx |
| `BK_Rəsmi əvəzetmə (həvalə)` | Rəsmi əvəzetmə(...).xlsx — **qırmızı sətirlərin fonu `#FFC7CE` ilə bərpa edilir** |
| `BK_Vəzifə-Maaş` | **6 aylıq sheet BİR vərəqdə alt-alta**, əvvələ 2 sütun əlavə edilir: `Ay` (sheet adı) və `Ay tarixi` (`2026-01` … `2026-06`) |

Vərəq adları Excel limitinə görə maksimum 31 simvol.

---

## 11. Texniki icra qeydləri

- `.xls` oxumaq üçün `xlrd>=2.0.1` quraşdır.
- Formulalı fayl yazıldıqdan sonra **LibreOffice headless ilə recalc mütləqdir**
  (`soffice --headless` və ya mövcud recalc skripti) — nəticədə `#REF!/#DIV/0!/#VALUE!/#NAME?` **sıfır olmalıdır**.
- `openpyxl` ilə tarix `NaT` dəyərləri xanaya yazılmazdan əvvəl `None`-a çevrilməlidir.
- Excel Table `displayName`-lərində boşluq/xüsusi simvol olmaz.
- Bütün şriftlər Arial. Böyük vərəqlərdə (backup) sütun eni 18.

---

## 12. MƏCBURİ YOXLAMA MƏRHƏLƏSİ (tool nəticəni qaytarmazdan əvvəl)

Tool aşağıdakı testləri **avtomatik** icra etməli və hesabatla birlikdə nəticələrini çap etməlidir.
Hər hansı biri uğursuzdursa — nəticə qaytarılmır, xəta izah edilir:

1. **Recalc:** LibreOffice recalc `status=success`, `total_errors=0`.
2. **Müstəqil cross-check:** günlük gün cəmi və saatlıq saat cəmi pandas ilə (düsturlardan asılı olmayaraq)
   ikinci dəfə hesablanır və Exceldəki dəyərlərlə **tam üst-üstə düşməlidir**.
3. **Təmiz Data:** boş və ya ≤0 `Gün sayı` = 0 ədəd.
4. **Yekun vərəq:** `(kod, Başlama, Bitmə)` üzrə duplikat = 0; `XƏTA` qeydli sətir = 0; boş/mənfi gün = 0;
   bütün sətirlər Bəli mənşəlidir.
5. **Əməkdaş üzrə cəm:** `max(Unikal təqvim günü) ≤ 181`; hər sətirdə `iş günü + şənbə/bazar = unikal gün`;
   ən böyük 3 dəyər Python-la yenidən hesablanıb tutuşdurulur.
6. **Balans:** `Təmiz sətir = Bəli + Xeyr`; `normalizasiya olunan − duplikat − çıxarılan xətalı = təmiz sətir`.
7. **Müvəqqəti keçid 174 yoxlaması:** mənbə fayldakı xərc-uyğun sətir sayı hesablanır və izlənir:
   `yekunda Müvəqqəti keçid etiketi ilə qalan + digər mənbə altında saxlanılan Bəli duplikat əkizləri = mənbə sayı`.
8. **Statistika kartları** Təmiz Data / Yekun vərəqdən pandas ilə alınan dəyərlərlə üst-üstə düşür.

---

## 13. Tool interfeysi (tövsiyə)

- Giriş: 5 faylın yolu (defolt — repo kökündəki adlar; ad dəyişə bilər deyə parametrlə ötürülə bilsin).
- Çıxış: `Əvəzetmə_Hesabatı.xlsx` + konsola yoxlama xülasəsi (bölmə 12-nin nəticələri).
- Kod modul strukturu: `normalize.py` (mərhələ 1) → `checks.py` (mərhələ 2) → `report.py` (mərhələ 3–4)
  → `verify.py` (mərhələ 5). Konfiqurasiya sabitləri ayrıca `config.py`-də.
- Bütün istisna/flaq mətnləri bu sənəddəki **dəqiq Azərbaycan dilində formulirovkalarla** yazılmalıdır —
  hesabat auditə gedir və mətnlər əvvəlki versiyalarla müqayisə ediləcək.

---

## Əlavə A. Hazırkı data ilə gözlənilən kontrol rəqəmlər

Tool eyni giriş faylları ilə işlədikdə aşağıdakı rəqəmləri (±0) verməlidir — reqressiya testi kimi istifadə et:

| Göstərici | Dəyər |
|---|---|
| Normallaşdırılan sətir | 10 634 |
| Təmiz Data sətri | 10 120 |
| — Xərc yaradan (Bəli) | 4 671 |
| — Xərc yaratmayan (Xeyr) | 5 449 |
| Təmiz Data cəmi gün (günlük) | 47 804 |
| Təmiz Data cəmi saat (saatlıq) | 5 907,15 |
| Yekun vərəq sətri | 4 567 |
| Yekun cəmi gün (günlük) | 18 729 |
| Yekun cəmi saat (saatlıq) | 5 134,4 |
| Yekundan istisna (üst-üstə düşmə / epizod kəsişməsi) | 80 |
| Birləşdirilən İnzibatçı mənbə sətri | 41 |
| Unikal əməkdaş | 1 721 |
| Max unikal təqvim günü | 179 |
| Duplikat: fayl daxili / fayllar arası | 154 / 357 |
| Xəta Loqu qeydi (yekun istisnalar daxil) | ~1 590–1 600 |
| Üst-üstə düşən əvəzetmə / İç-içə hal | 430 / 588 |
| Qırmızı sətir (Rəsmi əvəzetmə) | 20 (Excel sətirləri: 24–27, 47–49, 51–53, 68–70, 80–81, 167–169, 173–174) |
| Müvəqqəti keçid xərc-uyğun mənbə sətri | 174 (32 unikal şəxs) |
