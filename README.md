# BirCheck

Excel VBA makroslarının web variantı — iş günü, davamiyyət və kəsişən əmrlərin yoxlanılması. Əvvəlcə saatlarla çəkən hesablamalar indi bir neçə saniyəyə tamamlanır.

Platforma parol ilə qorunur (`Payroll`) və üç tabdan ibarətdir:

- **Working days check** — MHMD ilə HR iş günlərinin müqayisəsi
- **Absenteeism** — son 63 gün üzrə davamiyyət hesablaması
- **Cross-checking** — eyni işçidə tarixləri kəsişən əmrlərin aşkarlanması

## Necə işləyir

Tamamilə **brauzerdə** işləyir (client-side). Excel faylı serverə yüklənmir, bütün hesablama istifadəçinin kompüterində aparılır. Vercel sadəcə static fayllar serve edir.

### Working days check — tələb olunan Excel strukturu

Bir fayl, 4 sheet:

| Sheet | Məzmun |
|---|---|
| `From_Excel_Baza` | İşçi bazası (A:I) — A: personalkod, ..., H: işə qəbul tarixi |
| `From_MHMD` | Muhasibatlıqdan gələn (A: kod, C: tip "001", D-E: tarix aralığı, F: gün) |
| `From_HRB_Otpusk` | Məzuniyyət (A: kod, E-F-G: detallar, H: günlər, I2:I29: bayramlar, J2: başlama, J3: bitmə, K2: ümumi iş günü, L2:L29: 6 günlük qeyri-iş günləri, M2: 6 günlük ayda iş günü sayı `=(DAYS(J3;J2)+1)-COUNT($L$2:$L$29)`) |
| `Problems` | Nəticə yazılacaq (boş ola bilər) |

### Hesablama məntiqi (VBA-dan tərcümə)

1. `From_MHMD` təmizlənir — J2..J3 aralığından kənar sətirlər atılır
2. `From_Excel_Baza` A:I → `Problems` A:I köçürülür
3. Hər işçi üçün:
   - **J** = `MHMD` -dən sum(F) (yalnız C="001")
   - **K** = işə qəbul tarixi J2-dən əvvəldirsə `ay_iş_günü - otpusk`, əks halda `NETWORKDAYS.INTL(H, J3, 11, bayramlar) - otpusk`
     - **6 günlük iş** rejimində (Baza I sütunu "6 günlük iş" ilə başlayır): bayramlar = `L2:L29`, ay iş günü = `M2`
     - digər işçilər: bayramlar = `I2:I29`, ay iş günü = `K2`
   - **L** = K - J
4. Fərqi 0 olan sətirlər interfeysdə gizlənir (fayla hamısı daxil olur)
5. K < ay iş günü olanlara `From_HRB_Otpusk`-dan məzuniyyət detalları əlavə olunur (M)

## Cross-checking — kəsişən əmrlər

Bir sheet (`Dates` və ya ilk sheet), sütunlar: **A**=Personal kod, **B**=Növ, **C**=Tarixdən, **D**=Tarixə.

Eyni işçi (Personal kod) üzrə tarix aralıqları kəsişən (üst-üstə düşən və ya sərhəd günü paylaşan) əmr cütləri tapılır. Nəticə unikal işçi (Employee Badge) üzrə qruplaşdırılır və hər cütün statusları (Növ) ilə tarixləri göstərilir. `NETWORKDAYS` deyil, sadə interval kəsişməsi: `start₁ ≤ end₂ AND start₂ ≤ end₁`. Yalnız kəsişən caseler göstərilir; nəticə Excel-ə export oluna bilər.

## Absenteeism — davamiyyət

Bir sheet (`Data`), sütunlar: A=Personal kod, C=Növ kodu, E=Başlama, F=Bitmə. Son 63 gün aralığında hər işçi üzrə davamiyyət müddəti hesablanır (42/29 növləri atılır, 30/32 üçün xüsusi çəki, zəncir filtri). Gün üzrə filtr (`<`, `≤`, `=`, `≥`, `>`, aralıq), sıralama və seçilmiş/hamısı export.

## Lokal işlətmək

```bash
npm install
npm run dev
```

`http://localhost:3000`

## Vercel-ə deploy

Repo-nu Vercel-ə qoş — heç bir konfiqurasiya lazım deyil. Next.js avtomatik aşkarlanır.

## Stack

- Next.js 14 (App Router) + TypeScript
- Tailwind CSS
- [SheetJS](https://sheetjs.com/) (xlsx)
