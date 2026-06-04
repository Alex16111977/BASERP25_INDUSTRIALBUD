-- =====================================================================================
-- 08_dim_warehouses.sql
--
-- Dim_Warehouses — измерение «Склад» для Fact_Balance.Warehouse_ID.
-- Источник: Справочник.Склады (заполняется Свод_СебестоимостьТоваров —
--   субконто МестоХранения → Склад в РегистрСведений.А_ОтчетБаланс_Свод).
-- До этого Fact_Balance.Warehouse_ID (4775 строк Себест) не имел Dim — пробел.
--
-- Справочник.Склады ИЕРАРХИЧЕСКИЙ (Иерархия групп и элементов, 2 уровня),
-- но БЕЗ Кода (Длина кода=0 → в _Reference502 нет _Code; _ParentIDRRef/_Folder
-- есть). Рекурсивный CTE как dim_items, но без Кода (паттерн dim_organizations).
-- Даёт Parent_ID + Level1..5 для иерархии «ИерархияСкладов» в PL.pbix.
-- ETL via pipelines/dim_catalogs.json step "dim_warehouses" (raw_sql).
-- Idempotent: DROP + CREATE + unknown member.
-- =====================================================================================

USE OlapBASERP;
GO

PRINT '=== 08_dim_warehouses.sql -- start ===';
GO

IF OBJECT_ID('dbo.Dim_Warehouses', 'U') IS NOT NULL
BEGIN
    DROP TABLE dbo.Dim_Warehouses;
    PRINT '  Dropped existing Dim_Warehouses';
END
GO

CREATE TABLE dbo.Dim_Warehouses (
    Warehouse_ID         char(32) PRIMARY KEY,
    Warehouse_Name       nvarchar(250) NOT NULL,
    Parent_ID            char(32)      NULL,
    Is_Group             bit NOT NULL DEFAULT 0,
    Marked_For_Deletion  bit NOT NULL DEFAULT 0,
    Hierarchy_Path       nvarchar(500) NULL,
    Hierarchy_Depth      int           NULL,
    Level1               nvarchar(150) NULL,
    Level2               nvarchar(150) NULL,
    Level3               nvarchar(150) NULL,
    Level4               nvarchar(150) NULL,
    Level5               nvarchar(150) NULL,
    Loaded_At            datetime2 NOT NULL DEFAULT SYSDATETIME()
);
PRINT '  Created: Dim_Warehouses';
GO

-- Unknown member (orphan FK fallback: плуги Свода имеют пустой Склад)
INSERT INTO dbo.Dim_Warehouses
    (Warehouse_ID, Warehouse_Name, Parent_ID, Is_Group, Marked_For_Deletion,
     Hierarchy_Path, Hierarchy_Depth, Level1)
VALUES (REPLICATE('0', 32), N'(не вказано)', NULL, 0, 0,
        N'(не вказано)', 1, N'(не вказано)');
PRINT '  Inserted unknown member';
GO

PRINT '=== Final Dim_Warehouses schema ===';
SELECT
    ORDINAL_POSITION AS pos,
    COLUMN_NAME      AS col,
    DATA_TYPE        AS type,
    CHARACTER_MAXIMUM_LENGTH AS len,
    IS_NULLABLE      AS nullable
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Dim_Warehouses'
ORDER BY ORDINAL_POSITION;

PRINT '=== 08_dim_warehouses.sql -- done ===';
GO
