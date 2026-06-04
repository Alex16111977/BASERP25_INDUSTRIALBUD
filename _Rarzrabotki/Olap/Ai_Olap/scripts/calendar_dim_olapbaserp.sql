-- =========================================================================
-- OlapBASERP.dbo.Calendar — повне перестворення (Stage 2026-05-06)
-- Період: 2024-01-01 → 2027-12-31 (1461 рядок)
--
-- 40 колонок: дата + рік/квартал/місяць/тиждень/день + назви ru/uk/en + short
-- + month_decade/month_week + last_day flags + days_in_* + weekend
-- + 8 NEW: Квартал, Period_Month/Quarter/Year, Year_*_Sort, Is_Working_Day, Days_To_Month_End
--
-- Прибрано (15 cols vs оригіналу 47): zodiac/chinese/season/semester/century/by/week_date
--
-- Параметри періоду на початку — для зміни в майбутньому без правок логіки.
-- Запуск: idempotent (DROP IF EXISTS + CREATE + INSERT)
-- =========================================================================
USE OlapBASERP;
GO
SET LANGUAGE Russian;
GO

IF OBJECT_ID('dbo.Calendar', 'U') IS NOT NULL DROP TABLE dbo.Calendar;
GO

CREATE TABLE dbo.Calendar (
    date_              DATETIME      NOT NULL PRIMARY KEY,
    year_              INT           NOT NULL,
    quarter_           INT           NOT NULL,
    Квартал            NVARCHAR(3)   NOT NULL,
    month_             INT           NOT NULL,
    day_               INT           NOT NULL,
    year_day           INT           NOT NULL,
    week_              INT           NOT NULL,
    week_day           INT           NOT NULL,
    month_name              NVARCHAR(20),
    month_name_en           NVARCHAR(20),
    month_name_ua           NVARCHAR(20),
    month_name_short        NVARCHAR(5),
    month_name_short_en     NVARCHAR(5),
    month_name_short_ua     NVARCHAR(5),
    year_month_short_ua     NVARCHAR(15),
    week_day_str            NVARCHAR(20),
    week_day_str_en         NVARCHAR(20),
    week_day_str_ua         NVARCHAR(20),
    week_day_short          NVARCHAR(3),
    week_day_short_en       NVARCHAR(3),
    year_week         VARCHAR(7)   NOT NULL,
    year_month        VARCHAR(7)   NOT NULL,
    Year_Quarter      VARCHAR(11)  NOT NULL,
    year_quarter_en   VARCHAR(15)  NOT NULL,
    date_str          VARCHAR(10)  NOT NULL,
    date_full_str     NVARCHAR(40),
    date_full_str_en  VARCHAR(40),
    month_week        INT,
    month_decade      INT,
    month_last_day    NVARCHAR(2),
    year_last_day     NVARCHAR(2),
    days_in_month     INT,
    days_in_year      INT,
    weekend           BIT,
    Period_Month      DATE NOT NULL,
    Period_Quarter    DATE NOT NULL,
    Period_Year       DATE NOT NULL,
    Year_Month_Sort   INT  NOT NULL,
    Year_Quarter_Sort INT  NOT NULL,
    Year_Week_Sort    INT  NOT NULL,
    Is_Working_Day    BIT  NOT NULL,
    Days_To_Month_End INT  NOT NULL
);
GO

-- =========================================================================
-- INSERT через recursive CTE (1461 ітерацій → MAXRECURSION 0)
-- =========================================================================
DECLARE @d_start DATE = '20240101';
DECLARE @d_end   DATE = '20271231';

;WITH dates AS (
    SELECT @d_start AS d
    UNION ALL
    SELECT DATEADD(day, 1, d) FROM dates WHERE d < @d_end
)
INSERT INTO dbo.Calendar (
    date_, year_, quarter_, Квартал, month_, day_, year_day, week_, week_day,
    month_name, month_name_en, month_name_ua, month_name_short, month_name_short_en,
    month_name_short_ua, year_month_short_ua,
    week_day_str, week_day_str_en, week_day_str_ua, week_day_short, week_day_short_en,
    year_week, year_month, Year_Quarter, year_quarter_en, date_str,
    date_full_str, date_full_str_en,
    month_week, month_decade, month_last_day, year_last_day,
    days_in_month, days_in_year, weekend,
    Period_Month, Period_Quarter, Period_Year,
    Year_Month_Sort, Year_Quarter_Sort, Year_Week_Sort,
    Is_Working_Day, Days_To_Month_End
)
SELECT
    CAST(d AS DATETIME)                                          AS date_,
    YEAR(d)                                                      AS year_,
    DATEPART(qq, d)                                              AS quarter_,
    'Кв' + CAST(DATEPART(qq, d) AS VARCHAR(1))                   AS Квартал,
    MONTH(d)                                                     AS month_,
    DAY(d)                                                       AS day_,
    DATEPART(dy, d)                                              AS year_day,
    DATEPART(ww, d)                                              AS week_,
    DATEPART(dw, d)                                              AS week_day,
    DATENAME(mm, d)                                              AS month_name,
    CASE MONTH(d)
        WHEN 1 THEN 'January' WHEN 2 THEN 'February' WHEN 3 THEN 'March'
        WHEN 4 THEN 'April'   WHEN 5 THEN 'May'      WHEN 6 THEN 'June'
        WHEN 7 THEN 'July'    WHEN 8 THEN 'August'   WHEN 9 THEN 'September'
        WHEN 10 THEN 'October' WHEN 11 THEN 'November' WHEN 12 THEN 'December'
    END                                                          AS month_name_en,
    CASE MONTH(d)
        WHEN 1 THEN N'Січень'   WHEN 2 THEN N'Лютий'    WHEN 3 THEN N'Березень'
        WHEN 4 THEN N'Квітень'  WHEN 5 THEN N'Травень'  WHEN 6 THEN N'Червень'
        WHEN 7 THEN N'Липень'   WHEN 8 THEN N'Серпень'  WHEN 9 THEN N'Вересень'
        WHEN 10 THEN N'Жовтень' WHEN 11 THEN N'Листопад' WHEN 12 THEN N'Грудень'
    END                                                          AS month_name_ua,
    CASE MONTH(d)
        WHEN 1 THEN N'Янв' WHEN 2 THEN N'Фев' WHEN 3 THEN N'Мар'
        WHEN 4 THEN N'Апр' WHEN 5 THEN N'Май' WHEN 6 THEN N'Июн'
        WHEN 7 THEN N'Июл' WHEN 8 THEN N'Авг' WHEN 9 THEN N'Сен'
        WHEN 10 THEN N'Окт' WHEN 11 THEN N'Ноя' WHEN 12 THEN N'Дек'
    END                                                          AS month_name_short,
    CASE MONTH(d)
        WHEN 1 THEN 'Jan' WHEN 2 THEN 'Feb' WHEN 3 THEN 'Mar'
        WHEN 4 THEN 'Apr' WHEN 5 THEN 'May' WHEN 6 THEN 'Jun'
        WHEN 7 THEN 'Jul' WHEN 8 THEN 'Aug' WHEN 9 THEN 'Sep'
        WHEN 10 THEN 'Oct' WHEN 11 THEN 'Nov' WHEN 12 THEN 'Dec'
    END                                                          AS month_name_short_en,
    CASE MONTH(d)
        WHEN 1 THEN N'Січ' WHEN 2 THEN N'Лют' WHEN 3 THEN N'Бер'
        WHEN 4 THEN N'Кві' WHEN 5 THEN N'Тра' WHEN 6 THEN N'Чер'
        WHEN 7 THEN N'Лип' WHEN 8 THEN N'Сер' WHEN 9 THEN N'Вер'
        WHEN 10 THEN N'Жов' WHEN 11 THEN N'Лис' WHEN 12 THEN N'Гру'
    END                                                          AS month_name_short_ua,
    CAST(YEAR(d) AS NVARCHAR(4)) + N' ' + CASE MONTH(d)
        WHEN 1 THEN N'Січ' WHEN 2 THEN N'Лют' WHEN 3 THEN N'Бер'
        WHEN 4 THEN N'Кві' WHEN 5 THEN N'Тра' WHEN 6 THEN N'Чер'
        WHEN 7 THEN N'Лип' WHEN 8 THEN N'Сер' WHEN 9 THEN N'Вер'
        WHEN 10 THEN N'Жов' WHEN 11 THEN N'Лис' WHEN 12 THEN N'Гру'
    END                                                          AS year_month_short_ua,
    DATENAME(dw, d)                                              AS week_day_str,
    CASE DATEPART(dw, d)
        WHEN 1 THEN 'Monday' WHEN 2 THEN 'Tuesday' WHEN 3 THEN 'Wednesday'
        WHEN 4 THEN 'Thursday' WHEN 5 THEN 'Friday'
        WHEN 6 THEN 'Saturday' WHEN 7 THEN 'Sunday'
    END                                                          AS week_day_str_en,
    CASE DATEPART(dw, d)
        WHEN 1 THEN N'понеділок' WHEN 2 THEN N'вівторок' WHEN 3 THEN N'середа'
        WHEN 4 THEN N'четвер'    WHEN 5 THEN N'п''ятниця'
        WHEN 6 THEN N'субота'    WHEN 7 THEN N'неділя'
    END                                                          AS week_day_str_ua,
    CASE DATEPART(dw, d)
        WHEN 1 THEN N'Пн' WHEN 2 THEN N'Вт' WHEN 3 THEN N'Ср'
        WHEN 4 THEN N'Чт' WHEN 5 THEN N'Пт' WHEN 6 THEN N'Сб' WHEN 7 THEN N'Вс'
    END                                                          AS week_day_short,
    CASE DATEPART(dw, d)
        WHEN 1 THEN 'Mo' WHEN 2 THEN 'Tu' WHEN 3 THEN 'We'
        WHEN 4 THEN 'Th' WHEN 5 THEN 'Fr' WHEN 6 THEN 'Sa' WHEN 7 THEN 'Su'
    END                                                          AS week_day_short_en,
    CAST(YEAR(d) AS VARCHAR(4)) +
        CASE WHEN DATEPART(ww, d) < 10 THEN '.0' ELSE '.' END +
        CAST(DATEPART(ww, d) AS VARCHAR(2))                      AS year_week,
    CAST(YEAR(d) AS VARCHAR(4)) +
        CASE WHEN MONTH(d) < 10 THEN '.0' ELSE '.' END +
        CAST(MONTH(d) AS VARCHAR(2))                             AS year_month,
    CAST(YEAR(d) AS VARCHAR(4)) + '.' +
        CAST(DATEPART(qq, d) AS VARCHAR(1)) + ' кв.'             AS Year_Quarter,
    CAST(YEAR(d) AS VARCHAR(4)) + '.' +
        CAST(DATEPART(qq, d) AS VARCHAR(1)) + ' quarter'         AS year_quarter_en,
    CONVERT(VARCHAR(10), d, 104)                                 AS date_str,
    -- date_full_str: '06 травня 2026р.'
    (CASE WHEN DAY(d) < 10 THEN '0' ELSE '' END) +
    CAST(DAY(d) AS VARCHAR(2)) + N' ' +
    CASE MONTH(d)
        WHEN 1 THEN N'січня' WHEN 2 THEN N'лютого' WHEN 3 THEN N'березня'
        WHEN 4 THEN N'квітня' WHEN 5 THEN N'травня' WHEN 6 THEN N'червня'
        WHEN 7 THEN N'липня' WHEN 8 THEN N'серпня' WHEN 9 THEN N'вересня'
        WHEN 10 THEN N'жовтня' WHEN 11 THEN N'листопада' WHEN 12 THEN N'грудня'
    END +
    N' ' + CAST(YEAR(d) AS VARCHAR(4)) + N'р.'                   AS date_full_str,
    -- date_full_str_en: 'May 06 2026'
    CASE MONTH(d)
        WHEN 1 THEN 'January' WHEN 2 THEN 'February' WHEN 3 THEN 'March'
        WHEN 4 THEN 'April'   WHEN 5 THEN 'May'      WHEN 6 THEN 'June'
        WHEN 7 THEN 'July'    WHEN 8 THEN 'August'   WHEN 9 THEN 'September'
        WHEN 10 THEN 'October' WHEN 11 THEN 'November' WHEN 12 THEN 'December'
    END +
    (CASE WHEN DAY(d) < 10 THEN ' 0' ELSE ' ' END) +
    CAST(DAY(d) AS VARCHAR(2)) + ' ' +
    CAST(YEAR(d) AS VARCHAR(4))                                  AS date_full_str_en,
    -- month_week (5 категорій)
    CASE
        WHEN DAY(d) < 8                  THEN 1
        WHEN DAY(d) BETWEEN 8 AND 14     THEN 2
        WHEN DAY(d) BETWEEN 15 AND 21    THEN 3
        WHEN DAY(d) BETWEEN 22 AND 28    THEN 4
        WHEN DAY(d) > 28                 THEN 5
    END                                                          AS month_week,
    -- month_decade (декада 1-3)
    CASE
        WHEN DAY(d) < 11                  THEN 1
        WHEN DAY(d) BETWEEN 11 AND 20     THEN 2
        WHEN DAY(d) > 20                  THEN 3
    END                                                          AS month_decade,
    -- month_last_day flag
    CASE WHEN DAY(d) = DAY(EOMONTH(d)) THEN N'да' ELSE N'' END   AS month_last_day,
    -- year_last_day flag
    CASE WHEN MONTH(d) = 12 AND DAY(d) = 31 THEN N'да' ELSE N'' END AS year_last_day,
    DAY(EOMONTH(d))                                              AS days_in_month,
    DATEPART(dy, DATEFROMPARTS(YEAR(d), 12, 31))                 AS days_in_year,
    CASE WHEN DATEPART(dw, d) IN (6, 7) THEN 1 ELSE 0 END        AS weekend,
    -- ⭐ NEW
    DATEFROMPARTS(YEAR(d), MONTH(d), 1)                          AS Period_Month,
    DATEFROMPARTS(YEAR(d), (DATEPART(qq, d) - 1) * 3 + 1, 1)     AS Period_Quarter,
    DATEFROMPARTS(YEAR(d), 1, 1)                                 AS Period_Year,
    YEAR(d) * 100 + MONTH(d)                                     AS Year_Month_Sort,
    YEAR(d) * 10  + DATEPART(qq, d)                              AS Year_Quarter_Sort,
    YEAR(d) * 100 + DATEPART(ww, d)                              AS Year_Week_Sort,
    CASE WHEN DATEPART(dw, d) IN (6, 7) THEN 0 ELSE 1 END        AS Is_Working_Day,
    DATEDIFF(day, d, EOMONTH(d))                                 AS Days_To_Month_End
FROM dates
OPTION (MAXRECURSION 0);
GO

-- Sanity checks
SELECT COUNT(*) AS Cnt, MIN(date_) AS DateMin, MAX(date_) AS DateMax FROM dbo.Calendar;
SELECT Квартал, COUNT(*) AS Cnt FROM dbo.Calendar GROUP BY Квартал ORDER BY 1;
SELECT TOP 5 date_, year_month, Квартал, Year_Month_Sort, Period_Month, Is_Working_Day, Days_To_Month_End
FROM dbo.Calendar ORDER BY date_;
GO
