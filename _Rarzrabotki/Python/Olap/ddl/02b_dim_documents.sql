-- =====================================================================================
-- 02b_dim_documents.sql
--
-- Create Dim_Documents — drill-down dimension for Fact_PnL.Document_ID FK.
-- Mirrors InformationRegister.А_ДокументРасшифровка (1 dim + 7 res).
--
-- Sourced via Python ETL from `_InfoRg<N>` (А_ДокументРасшифровка storage table).
-- After this DDL, ETL pipeline must add `dim_documents` to load list.
--
-- Idempotent: re-run safe (CREATE IF NOT EXISTS, UNIQUE INSERT for unknown member).
-- =====================================================================================

USE OlapBASERP;
GO

PRINT '=== 02b_dim_documents.sql — start ===';
GO

IF OBJECT_ID('dbo.Dim_Documents', 'U') IS NULL
BEGIN
    CREATE TABLE Dim_Documents (
        Document_ID                 char(32)        NOT NULL PRIMARY KEY,
        Document_Type_ERP           varchar(60)     NOT NULL,
        Document_Number_ERP         varchar(20)     NULL,
        Document_Date               datetime2       NULL,
        Document_Number_BuhBud      varchar(20)     NULL,
        Document_Type_BuhBud        varchar(60)     NULL,
        Document_Presentation       nvarchar(500)   NULL,
        Document_Comment            nvarchar(1000)  NULL,
        Marked_For_Deletion         bit             NOT NULL DEFAULT 0,
        Loaded_At                   datetime2       NOT NULL DEFAULT SYSDATETIME(),
        INDEX IX_DimDocs_TypeERP   (Document_Type_ERP),
        INDEX IX_DimDocs_TypeBuh   (Document_Type_BuhBud),
        INDEX IX_DimDocs_DateMonth (Document_Date)
    );
    PRINT '  Created: Dim_Documents (10 columns, 3 indexes)';
END
ELSE PRINT '  Skip create Dim_Documents (already exists)';
GO

-- Unknown member — для NULL FK на Fact_PnL.Document_ID; гарантує що JOIN
-- завжди знайде пару → запобігає broken refs у Power BI.
INSERT INTO Dim_Documents (
    Document_ID,
    Document_Type_ERP,
    Document_Number_ERP,
    Document_Presentation
)
SELECT
    '00000000000000000000000000000001',
    '(Empty)',
    '',
    '(Empty)'
WHERE NOT EXISTS (
    SELECT 1 FROM Dim_Documents
    WHERE Document_ID = '00000000000000000000000000000001'
);
GO

PRINT '=== Dim_Documents content ===';
SELECT Document_ID, Document_Type_ERP, Document_Presentation FROM Dim_Documents ORDER BY Document_ID;

PRINT '=== Dim_Documents schema ===';
SELECT
    ORDINAL_POSITION AS pos,
    COLUMN_NAME      AS col,
    DATA_TYPE        AS type,
    CHARACTER_MAXIMUM_LENGTH AS len,
    IS_NULLABLE      AS nullable
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Dim_Documents'
ORDER BY ORDINAL_POSITION;

PRINT '=== 02b_dim_documents.sql — done ===';
GO
