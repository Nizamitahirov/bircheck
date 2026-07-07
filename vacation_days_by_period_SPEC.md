# Texniki Şərt: "Vacation Days by Period" Hesabatının Generasiyası

**Versiya:** 1.0
**Platforma:** bircheck (Claude Code task)
**Nəticə:** Ayrıca Excel faylı, tək səhifə — `Vacation days by period`

---

## 1. Məqsəd

Şirkət əməkdaşlarının məzuniyyət balansını **iş illəri (work-year period) üzrə** hesablamaq: hər əməkdaş × hər iş ili üçün bir sətir — illik hüquq (əsas + əlavə), istifadə olunmuş və istifadə olunmamış günlər.

Hesablanma məntiqi: **hər başlamış iş ilinə tam illik hüquq** yazılır (proporsional accrual DEYİL), istifadə isə **FIFO** prinsipi ilə ən köhnə dövrdən silinir.

---

## 2. İnputlar

Bir Excel workbook (və ya ayrı-ayrı fayllar) aşağıdakı mənbələrlə:

### 2.1. Master siyahı — `2015-2026` səhifəsi (və ya ən geniş dövrü əhatə edən icmal səhifə)

Cari əməkdaşların reyestri. Header 2 sətirdir, data 3-cü sətirdən başlayır.

| Sütun (0-index) | Məzmun |
|---|---|
| 0 | ID / Badge (məs. `COM 308`) |
| 1 | Ad Soyad Ata adı |
| 3 | İşə qəbul tarixi (`dd.mm.yyyy`, string) |
| 4 | İşdən çıxma tarixi (varsa) |
| 6 | Department → **Structure name** |
| 7 | Division → **Section name** |
| 10 | Position |
| 14–17 | Normalar: main, Child, Exper., Bad Cond. |

Nəticə hesabatına **yalnız bu səhifədəki əməkdaşlar** daxil edilir.

### 2.2. İllik səhifələr — `2018`, `2019`, … `2026`

Hər təqvim ili üzrə eyni strukturda səhifələr. Buradan **hər il üzrə normalar** götürülür (normalar illər üzrə dəyişir: staj artımı, uşaq günləri, 21→30 keçidi və s.):

- Sütun 14: `main` (əsas məzuniyyət norması)
- Sütun 15: `Child` (uşaq)
- Sütun 16: `Exper.` (staj)
- Sütun 17: `Bad Cond.` (zərərli şərait)
- **Əlavə günlər = Child + Exper. + Bad Cond.** (boş = 0)

Sütun 21 (`Used`) və 22 (`Balance Ending`) yalnız **validasiya** üçün istifadə olunur.

### 2.3. Əmrlər — `orders` səhifəsi

Bütün məzuniyyət əmrləri (1C ixracı, rus dilində sütunlar). İstifadə olunan sütunlar:

| Sütun | Məzmun |
|---|---|
| `Работник` | Əməkdaşın adı |
| `Badge` | Əməkdaş kodu |
| `Поле1` | Əmr növü (`Отпуск очередной`, `Больничный`, `Отпуск БСЗ`, `Компенсация` və s.) |
| `Шаблон` | Şablon (`Main Vacation`, `Məzuniyyət-ödənişli_2024` və s.) |
| `Готов к утверждению` | Təsdiq statusu (`Да`/`Нет`) |
| `from`, `to` | Məzuniyyət başlanğıc/son tarixi |
| `(дн)` | Gün sayı |
| `Дата приказа` | Əmr tarixi (fallback tarix kimi) |

---

## 3. Alqoritm — addım-addım

### Addım 1. Faylın oxunması və açar normalizasiyası

1. Bütün səhifələri `pandas` ilə oxu (`skiprows=2` illik və master səhifələr üçün, `header=0` orders üçün).
2. Uyğunlaşdırma açarları:
   - `nkey = UPPER(bütün boşluqlar silinmiş ad)` — **əsas açar addır**, çünki badge kodları müxtəlif adamlar arasında təkrar istifadə olunur (badge reuse), boşluq fərqləri var (`COM3` vs `COM 3`).
   - Master siyahıda `nkey` unikallığını yoxla; unikal deyilsə, `nkey + badge` kombinasiyasına keç və logla.
3. Tarixləri `dd.mm.yyyy` formatından parse et (`errors='coerce'`, boşları logla).

### Addım 2. Hər əməkdaş üçün illik normalar lüğəti

1. Hər illik səhifədən: `norms[(nkey, il)] = (main, child+exper+bad)`.
2. **Rehire filtri:** yalnız `Start date == master Start date` olan sətirləri götür (yenidən işə qəbul olunanların əvvəlki iş dövrünün normaları istisna edilir). Heç bir sətir tapılmasa, `nkey` üzrə istənilən sətri fallback kimi götür.
3. Son fallback: master səhifədəki normalar.

### Addım 3. Əmrlərin filtrasiyası (istifadə olunmuş günlər)

1. **Daxil et:** `Поле1 ∈ {Отпуск очередной, Компенсация}` VƏ `Готов к утверждению == Да`.
2. **İstisna et** (illik hüquqdan sayılmır): xəstəlik (`Больничный`), ödənişsiz (`Отпуск БСЗ` / Unpaid), təhsil (`Отпуск учебный` / Studing vacation), sosial/analıq, atalıq, hərbi çağırış, 125-ci maddə.
3. **Dedup:** `(nkey, from, to, дн)` üzrə dublikatları sil (eyni məzuniyyətə iki əmr nömrəsi halları mövcuddur).
4. Effektiv tarix: `effD = from` (boşdursa `Дата приказа`).
5. **Cari iş dövrü filtri:** yalnız `effD >= master işə qəbul tarixi` olan əmrlər sayılır — rehire olunanların əvvəlki iş dövrü istifadələri cari balansa təsir etmir.
6. Hər əməkdaş üçün: `total_used = Σ (дн)`.

### Addım 4. İş illərinin (dövrlərin) generasiyası

Hər master əməkdaş üçün, `hire = işə qəbul tarixi`:

```
i = 0
while (hire + i il) <= BUGÜN:            # relativedelta(years=i)
    period_start = hire + i il
    period_end   = hire + (i+1) il - 1 gün
    period_norması → Addım 2 lüğətindən, il = period_start.year
                     (o il yoxdursa → ən yaxın mövcud il)
    i += 1
```

Qeydlər: 29 fevral halları `relativedelta` ilə avtomatik idarə olunur; işdən çıxma tarixi varsa, ondan sonra başlayan dövrlər generasiya edilmir.

### Addım 5. FIFO bölüşdürməsi

```
rem = total_used
hər dövr (köhnədən yeniyə):
    used_main = min(rem, main_norma);  rem -= used_main
    used_add  = min(rem, add_norma);   rem -= used_add
əgər rem > 0:                          # avans məzuniyyət halı
    son dövrün used_main += rem        # unused mənfi ola bilər — normaldır, logla
```

İnvariant: `Σ(used_main + used_add) == total_used` (bütün əməkdaşlar üzrə cəm əmrlərlə tam üst-üstə düşməlidir).

### Addım 6. Nəticə faylının yazılması

**Ayrıca yeni fayl**, tək səhifə `Vacation days by period`, sütunlar A–Q:

| Sütun | Ad | Mənbə |
|---|---|---|
| A | Emp_code | master |
| B | Full name | master |
| C | Structure name | master (Department) |
| D | Section name | master (Division) |
| E | Position | master |
| F | Hire date | master |
| G | Period start date | Addım 4 |
| H | Period end date | Addım 4 |
| I | Main vacation days | illik səhifə norması |
| J | Total additional vacation days | Child+Exper+Bad |
| K | Total annual entitlement | **düstur** `=I+J` |
| L | Used days for main vacation | FIFO |
| M | Unused days for main vacation | **düstur** `=I-L` |
| N | Used days for additional vacation | FIFO |
| O | Unused days for additional vacation | **düstur** `=J-N` |
| P | Total used vacation days | **düstur** `=L+N` |
| Q | Total unused vacation days | **düstur** `=K-P` |

Format tələbləri: tarixlər `dd.mm.yyyy`; günlər tam ədəd; K, M, O, P, Q **Excel düsturu** kimi (hardcode yox); şrift Times New Roman 10; header tünd-mavi fon, bold; freeze panes `A2`; autofilter `A1:Q{son}`; nazik sərhədlər.

Yazıdan sonra düsturları recalc et (LibreOffice headless / `recalc` skripti) və **0 düstur xətası** (`#REF!`, `#DIV/0!` və s.) təsdiqlə.

### Addım 7. Validasiya (məcburi yoxlamalar)

1. **İstifadə cəmi:** `Σ used (nəticə) == Σ (дн) (filtrlənmiş əmrlər)` — dəqiq bərabər olmalıdır.
2. **İllik səhifə uzlaşması:** əməkdaş səviyyəsində `total_used` illik səhifələrin `Used` cəmi ilə müqayisə edilir; ≥95% uyğunluq gözlənilir. Fərqlər logda göstərilir (əsas səbəb: rehire kompensasiya qeydləri — bunlar qəsdən istisna olunur).
3. **Balans uzlaşması (məlumat xarakterli):** master `Balance Ending` ilə fərq gözləniləndir, çünki mənbə fayl proporsional accrual, bu hesabat isə tam illik hüquq metodunu istifadə edir. Fərq yalnız cari natamam dövrün payı qədər olmalıdır.
4. Əhatə: nəticədə əməkdaş sayı == master siyahıdakı say; hər əməkdaşın ən azı 1 dövrü var.
5. Anomaliya loqu: mənfi unused (avans), norması tapılmayan əməkdaş, parse olunmayan tarix, dublikat əmrlər.

---

## 4. Sərhəd halları (edge cases)

- **Badge reuse:** eyni kod müxtəlif adamlarda — uyğunlaşdırma ad üzrə.
- **Rehire:** master `Start` = son işə qəbul; əvvəlki iş dövrünün həm normaları, həm istifadəsi istisna.
- **Norma dəyişməsi** (məs. 21→30): hər dövr öz başlanğıc ilinin normasını alır.
- **Avans məzuniyyət:** istifadə > hüquq → son dövrdə mənfi unused, gizlədilmir.
- **Kompensasiya əmrləri:** `from` boş → əmr tarixi istifadə olunur.
- **Cari natamam dövr:** tam illik hüquqla yazılır (parametrlə proporsional rejimə keçirilə bilər — `PRORATE_LAST_PERIOD=true` opsiyası nəzərdə tutulsun).

## 5. Parametrlər

| Parametr | Default | Təsvir |
|---|---|---|
| `AS_OF_DATE` | bugün | Dövr generasiyasının son həddi |
| `PRORATE_LAST_PERIOD` | `false` | Son dövrün hüququnu `AS_OF_DATE`-ə qədər proporsional hesabla |
| `INCLUDE_COMPENSATION` | `true` | `Компенсация` əmrlərini istifadəyə say |
| `MASTER_SHEET` | `2015-2026` | Əməkdaş reyestri səhifəsi |
| `YEAR_SHEETS` | avtodetekt (`^\d{4}$`) | İllik norma səhifələri |
