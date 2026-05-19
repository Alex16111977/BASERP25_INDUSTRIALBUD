-- Dim_Contracts — Справочник.ДоговорыКонтрагентов (_Reference171).
-- Recreate: сохраняет старые колонки Stage-2 + добавляет денорм. имена,
-- Bool, FK на Dim_TipyDogovorov/Dim_FinAgents. Idempotent: DROP+CREATE
-- (loader full_reload = TRUNCATE+INSERT; колонки автодискавер из sys.columns).
IF OBJECT_ID('dbo.Dim_Contracts', 'U') IS NOT NULL
    DROP TABLE dbo.Dim_Contracts;
GO
CREATE TABLE dbo.Dim_Contracts (
    Contract_ID              char(32)      NOT NULL PRIMARY KEY,
    Contract_Code            varchar(50)   NULL,
    Contract_Name            nvarchar(300) NOT NULL,
    Parent_ID                char(32)      NULL,
    Is_Group                 bit           NOT NULL DEFAULT 0,
    Marked_For_Deletion      bit           NOT NULL DEFAULT 0,
    Is_FinAgent_Contract     bit           NOT NULL DEFAULT 0,
    TipDogovora              varchar(50)   COLLATE Cyrillic_General_CI_AS NULL,
    FinAgent_ID              char(32)      NULL,
    Department_Name          nvarchar(300) NULL,
    Dept_OkazUslug_Name      nvarchar(300) NULL,
    Partner_Name             nvarchar(300) NULL,
    Counterparty_Name        nvarchar(300) NULL,
    DDS_Article_Forced_Name  nvarchar(300) NULL,
    Org_Buh_Name             nvarchar(300) NULL,
    Loaded_At                datetime2     NOT NULL DEFAULT SYSDATETIME()
);
GO
