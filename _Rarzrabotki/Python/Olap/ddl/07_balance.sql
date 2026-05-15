-- 07_balance.sql -- Fact_Balance + Dim_PAP_Articles (idempotent). ASCII-only (no Cyrillic).
IF OBJECT_ID('dbo.Fact_Balance','U') IS NULL
CREATE TABLE Fact_Balance (
    Balance_ID          bigint IDENTITY(1,1) PRIMARY KEY,
    Period_Month        date NOT NULL,
    Period              datetime2 NULL,
    Source              varchar(40) NOT NULL,
    Recorder_Balance_ID char(32) NULL,
    Organization_ID     char(32) NOT NULL,
    Department_ID       char(32) NULL,
    PAP_Article_ID      char(32) NULL,
    Item_ID             char(32) NULL,
    Counterparty_ID     char(32) NULL,
    Partner_ID          char(32) NULL,
    Warehouse_ID        char(32) NULL,
    OperObject_ID       char(32) NULL,
    Contract_ID         char(32) NULL,
    Individual_ID       char(32) NULL,
    Cash_ID             char(32) NULL,
    SettlementObj_ID    char(32) NULL,
    Intangible_ID       char(32) NULL,
    Analytics1          nvarchar(150) NULL,
    Analytics2          nvarchar(150) NULL,
    Analytics3          nvarchar(150) NULL,
    Sum_Open            decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Inflow          decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Outflow         decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Close           decimal(15,2) NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME(),
    INDEX IX_Bal_Period_Source (Period_Month, Source),
    INDEX IX_Bal_Article       (PAP_Article_ID, Period_Month),
    INDEX IX_Bal_Individual    (Individual_ID, Period_Month),
    INDEX IX_Bal_SettlementObj (SettlementObj_ID, Period_Month)
);

IF OBJECT_ID('dbo.Dim_PAP_Articles','U') IS NULL
CREATE TABLE Dim_PAP_Articles (
    PAP_Article_ID      char(32) PRIMARY KEY,
    PAP_Article_Code    varchar(50) NULL,
    PAP_Article_Name    nvarchar(150) NOT NULL,
    Parent_ID           char(32) NULL,
    Is_Group            bit NOT NULL DEFAULT 0,
    AktivPassiv         varchar(15) NULL,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
