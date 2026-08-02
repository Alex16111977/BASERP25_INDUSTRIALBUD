-- =====================================================================================
-- OlapFactory -- data mart for production plan/fact (BaseERP -> Power BI)
-- Source: InformationRegister A_PlanFactProizvodstvo_Svod (_InfoRg56577)
-- ASCII ONLY: Cyrillic seed values are inserted by parametrized pyodbc (sqlcmd mangles them).
-- Apply: .venv/Scripts/python.exe scripts/apply_ddl.py
-- =====================================================================================

IF OBJECT_ID('dbo.ETL_Runs','U') IS NULL
CREATE TABLE ETL_Runs (
    Run_ID       bigint IDENTITY(1,1) PRIMARY KEY,
    Script       varchar(100) NOT NULL,
    Period_Month date NULL,
    Started_At   datetime2 NOT NULL DEFAULT SYSDATETIME(),
    Finished_At  datetime2 NULL,
    Rows_Loaded  int NULL,
    Status       varchar(20) NOT NULL,
    Error        nvarchar(4000) NULL
);
GO

-- ------------------------------------------------------------------ Calendar
-- NOT created here on purpose. dbo.Calendar is a 1:1 MIRROR of the live reference table
-- OlapBASERP.dbo.Calendar (43 columns, 2024-01-01 .. 2027-12-31) -- the very table that
-- PowerBi/Industrial/PL.pbix shows. It is created and filled by scripts/apply_ddl.py
-- (function sinhronizirovat_kalendar): the structure is read from the neighbour database
-- INFORMATION_SCHEMA and reproduced verbatim, the rows are copied by a cross-database
-- INSERT ... SELECT.
--   * why not here: the reference has a Cyrillic column name (Kvartal) and this file is
--     ASCII-only by convention;
--   * why not by re-running Olap/Ai_Olap/scripts/calendar_dim_olapbaserp.sql: that script is
--     3 columns behind the live table and carries USE OlapBASERP -- writing into the
--     neighbour contour is forbidden (reading is fine).

-- ------------------------------------------------------------------ Dimensions
IF OBJECT_ID('dbo.Dim_Organizations','U') IS NULL
CREATE TABLE Dim_Organizations (
    Organization_ID     char(32) NOT NULL PRIMARY KEY,
    Organization_Name   nvarchar(300) NOT NULL,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
GO

IF OBJECT_ID('dbo.Dim_Departments','U') IS NULL
CREATE TABLE Dim_Departments (
    Department_ID       char(32) NOT NULL PRIMARY KEY,
    Department_Code     varchar(50) NULL,
    Department_Name     nvarchar(300) NOT NULL,
    Parent_ID           char(32) NULL,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
GO

IF OBJECT_ID('dbo.Dim_Stages','U') IS NULL
CREATE TABLE Dim_Stages (
    Stage_ID            char(32) NOT NULL PRIMARY KEY,
    Stage_Code          varchar(50) NULL,
    Stage_Name          nvarchar(300) NOT NULL,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
GO

IF OBJECT_ID('dbo.Dim_Items','U') IS NULL
CREATE TABLE Dim_Items (
    Item_ID             char(32) NOT NULL PRIMARY KEY,
    Item_Code           varchar(50) NULL,
    Item_Name           nvarchar(300) NOT NULL,
    Is_Group            bit NOT NULL DEFAULT 0,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
GO

IF OBJECT_ID('dbo.Dim_CommonNames','U') IS NULL
CREATE TABLE Dim_CommonNames (
    CommonName_ID       char(32) NOT NULL PRIMARY KEY,
    CommonName_Code     varchar(50) NULL,
    CommonName          nvarchar(300) NOT NULL,
    Is_Group            bit NOT NULL DEFAULT 0,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
GO

IF OBJECT_ID('dbo.Dim_Units','U') IS NULL
CREATE TABLE Dim_Units (
    Unit_ID             char(32) NOT NULL PRIMARY KEY,
    Unit_Name           nvarchar(300) NOT NULL,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
GO

IF OBJECT_ID('dbo.Dim_Individuals','U') IS NULL
CREATE TABLE Dim_Individuals (
    Individual_ID       char(32) NOT NULL PRIMARY KEY,
    Individual_Name     nvarchar(300) NOT NULL,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
GO

-- Row type: seeded from 1C enum order (see seed_dim_row_types.py). Label matches FROZEN_ENUMS.
IF OBJECT_ID('dbo.Dim_RowTypes','U') IS NULL
CREATE TABLE Dim_RowTypes (
    RowType       varchar(20) COLLATE Cyrillic_General_CI_AS NOT NULL PRIMARY KEY,
    RowType_Sort  int NOT NULL,
    RowType_Group varchar(20) COLLATE Cyrillic_General_CI_AS NOT NULL
);
GO

-- Documents: drill-down from any fact number to the source 1C document
-- (timesheet / purchase / module-completion / cost-structure card).
IF OBJECT_ID('dbo.Dim_Documents','U') IS NULL
CREATE TABLE Dim_Documents (
    Document_ID           char(32) NOT NULL PRIMARY KEY,
    Document_Type         nvarchar(100) NOT NULL,
    Document_Number       varchar(50) NULL,
    Document_Date         datetime NULL,
    Document_Presentation nvarchar(300) NOT NULL,
    Loaded_At             datetime2 NOT NULL DEFAULT SYSDATETIME()
);
GO

-- ------------------------------------------------------------------ Fact
IF OBJECT_ID('dbo.Fact_PlanFactProizvodstvo','U') IS NULL
CREATE TABLE Fact_PlanFactProizvodstvo (
    Fact_ID         bigint IDENTITY(1,1) PRIMARY KEY,
    Period          datetime NOT NULL,
    Period_Month    date NOT NULL,
    Recorder_ID     char(32) NULL,
    Organization_ID char(32) NOT NULL,
    Department_ID   char(32) NULL,
    ExecutorDept_ID char(32) NULL,
    RowType         varchar(20) COLLATE Cyrillic_General_CI_AS NOT NULL,
    Block           nvarchar(100) NULL,
    Category        nvarchar(100) NULL,
    Stage_ID        char(32) NULL,
    Work_ID         char(32) NULL,
    CommonName_ID   char(32) NULL,
    Analytics_ID    char(32) NULL,
    Analytics_Text  nvarchar(400) NULL,
    Document_ID     char(32) NULL,
    Unit_ID         char(32) NULL,
    ItemName        nvarchar(300) NULL,
    PlanHours       decimal(18,3) NOT NULL DEFAULT 0,
    FactHours       decimal(18,3) NOT NULL DEFAULT 0,
    EarnedHours     decimal(18,3) NOT NULL DEFAULT 0,
    PlanQty         decimal(18,3) NOT NULL DEFAULT 0,
    FactQty         decimal(18,3) NOT NULL DEFAULT 0,
    PlanUAH         decimal(18,2) NOT NULL DEFAULT 0,
    FactUAH         decimal(18,2) NOT NULL DEFAULT 0,
    ETC_UAH         decimal(18,2) NOT NULL DEFAULT 0,
    Loaded_At       datetime2 NOT NULL DEFAULT SYSDATETIME(),
    INDEX IX_Fact_PFP_Period NONCLUSTERED (Period_Month, Department_ID),
    INDEX IX_Fact_PFP_RowType NONCLUSTERED (RowType)
);
GO

-- ------------------------------------------------------------------ Views
-- Department tree: ready-made drill Level1..Level5 for Power BI
-- (Proizvodstvo -> MD IRS 2026 -> 15 m -> No1). Recursive CTE, no extra pipeline.
IF OBJECT_ID('dbo.Dim_Departments_Tree','V') IS NOT NULL DROP VIEW dbo.Dim_Departments_Tree;
GO
CREATE VIEW dbo.Dim_Departments_Tree AS
WITH tree AS (
    SELECT d.Department_ID, d.Department_Code, d.Department_Name, d.Parent_ID,
           CAST(d.Department_Name AS nvarchar(1000)) AS Hierarchy_Path,
           1 AS Hierarchy_Depth,
           CAST(d.Department_Name AS nvarchar(300)) AS Level1,
           CAST(NULL AS nvarchar(300)) AS Level2,
           CAST(NULL AS nvarchar(300)) AS Level3,
           CAST(NULL AS nvarchar(300)) AS Level4,
           CAST(NULL AS nvarchar(300)) AS Level5
    FROM Dim_Departments d
    WHERE d.Parent_ID IS NULL OR d.Parent_ID = REPLICATE('0',32)
    UNION ALL
    SELECT c.Department_ID, c.Department_Code, c.Department_Name, c.Parent_ID,
           CAST(p.Hierarchy_Path + N' / ' + c.Department_Name AS nvarchar(1000)),
           p.Hierarchy_Depth + 1,
           p.Level1,
           CASE WHEN p.Hierarchy_Depth = 1 THEN c.Department_Name ELSE p.Level2 END,
           CASE WHEN p.Hierarchy_Depth = 2 THEN c.Department_Name ELSE p.Level3 END,
           CASE WHEN p.Hierarchy_Depth = 3 THEN c.Department_Name ELSE p.Level4 END,
           CASE WHEN p.Hierarchy_Depth = 4 THEN c.Department_Name ELSE p.Level5 END
    FROM Dim_Departments c
    JOIN tree p ON p.Department_ID = c.Parent_ID
    WHERE p.Hierarchy_Depth < 5
)
SELECT * FROM tree;
GO

-- Fact view for Power BI: resolves composite Analytics (item / individual / free text)
-- into one readable column, so the model needs no ambiguous relationships and no DAX lookups.
IF OBJECT_ID('dbo.vw_Fact_PlanFact','V') IS NOT NULL DROP VIEW dbo.vw_Fact_PlanFact;
GO
CREATE VIEW dbo.vw_Fact_PlanFact AS
SELECT f.Fact_ID, f.Period, f.Period_Month, f.Recorder_ID, f.Organization_ID,
       f.Department_ID, f.ExecutorDept_ID, f.RowType, f.Block, f.Category,
       f.Stage_ID, f.Work_ID, f.CommonName_ID, f.Analytics_ID, f.Analytics_Text,
       f.Document_ID, f.Unit_ID, f.ItemName,
       COALESCE(i.Item_Name, ind.Individual_Name, NULLIF(f.Analytics_Text, N''), N'(не задано)')
           AS Analytics_Name,
       f.PlanHours, f.FactHours, f.EarnedHours, f.PlanQty, f.FactQty,
       f.PlanUAH, f.FactUAH, f.ETC_UAH, f.Loaded_At
FROM Fact_PlanFactProizvodstvo f
LEFT JOIN Dim_Items i        ON i.Item_ID       = f.Analytics_ID
LEFT JOIN Dim_Individuals ind ON ind.Individual_ID = f.Analytics_ID;
GO
