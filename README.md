# BirCheck

Excel VBA makroslarının web variantı — iş günü və mezuniyyət fərqlərinin yoxlanılması. Əvvəlcə 2 saat çəkən hesablama indi bir neçə saniyəyə tamamlanır.

## Necə işləyir

Tamamilə **brauzerdə** işləyir (client-side). Excel faylı serverə yüklənmir, bütün hesablama istifadəçinin kompüterində aparılır. Vercel sadəcə static fayllar serve edir.

### Tələb olunan Excel strukturu

Bir fayl, 4 sheet:

| Sheet | Məzmun |
|---|---|
| `From_Excel_Baza` | İşçi bazası (A:I) — A: personalkod, ..., H: işə qəbul tarixi |
| `From_MHMD` | Muhasibatlıqdan gələn (A: kod, C: tip "001", D-E: tarix aralığı, F: gün) |
| `From_HRB_Otpusk` | Mezuniyyət (A: kod, E-F-G: detallar, H: günlər, I2:I29: bayramlar, J2: başlama, J3: bitmə, K2: ümumi iş günü) |
| `Problems` | Nəticə yazılacaq (boş ola bilər) |

### Hesablama məntiqi (VBA-dan tərcümə)

1. `From_MHMD` təmizlənir — J2..J3 aralığından kənar sətirlər atılır
2. `From_Excel_Baza` A:I → `Problems` A:I köçürülür
3. Hər işçi üçün:
   - **J** = `MHMD` -dən sum(F) (yalnız C="001")
   - **K** = işə qəbul tarixi J2-dən əvvəldirsə `K2 - otpusk`, əks halda `NETWORKDAYS.INTL(H, J3, 11, I2:I29) - otpusk`
   - **L** = K - J
4. L=0 olan sətirlər silinir
5. K < K2 olanlara `From_HRB_Otpusk`-dan mezuniyyət detalları əlavə olunur (M)

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
