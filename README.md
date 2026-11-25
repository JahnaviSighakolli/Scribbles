
IF OBJECT_ID('tempdb..#SQLVersions') IS NOT NULL
    DROP TABLE #SQLVersions;

CREATE TABLE #SQLVersions
(
    LinkedServerName  SYSNAME,
    ServerName        NVARCHAR(128),
    InstanceName      NVARCHAR(128),
    SQLVersion        NVARCHAR(50),
    ProductVersion    NVARCHAR(50),
    ProductLevel      NVARCHAR(50),
    CU_Level          NVARCHAR(50),
    KB_Number         NVARCHAR(50),
    BuildType         NVARCHAR(50),
    Edition           NVARCHAR(200),
    Status            NVARCHAR(500)
);

DECLARE @LinkedServer SYSNAME;
DECLARE @SQL NVARCHAR(MAX);

DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
SELECT name
FROM sys.servers
WHERE is_linked = 1
  AND name <> 'MSX'
  AND data_source IS NOT NULL;

OPEN cur;

FETCH NEXT FROM cur INTO @LinkedServer;

WHILE @@FETCH_STATUS = 0
BEGIN
    BEGIN TRY

        SET @SQL = N'
        INSERT INTO #SQLVersions
        (
            LinkedServerName,
            ServerName,
            InstanceName,
            SQLVersion,
            ProductVersion,
            ProductLevel,
            CU_Level,
            KB_Number,
            BuildType,
            Edition,
            Status
        )
        SELECT
            N''' + REPLACE(@LinkedServer,'''','''''') + N''',
            ServerName,
            InstanceName,

            CASE
                WHEN ProductVersion LIKE ''13.%''
                    THEN ''SQL Server 2016''
                WHEN ProductVersion LIKE ''14.%''
                    THEN ''SQL Server 2017''
                WHEN ProductVersion LIKE ''15.%''
                    THEN ''SQL Server 2019''
                WHEN ProductVersion LIKE ''16.%''
                    THEN ''SQL Server 2022''
                WHEN ProductVersion LIKE ''17.%''
                    THEN ''SQL Server 2025''
                ELSE ''Other / Unknown''
            END,

            ProductVersion,
            ProductLevel,
            CU_Level,
            KB_Number,
            BuildType,
            Edition,
            ''Success''

        FROM OPENQUERY(' + QUOTENAME(@LinkedServer) + N',
        ''
            SELECT
                CAST(@@SERVERNAME AS NVARCHAR(128))
                    AS ServerName,

                CAST(
                    ISNULL(
                        CAST(SERVERPROPERTY(''''InstanceName'''')
                        AS NVARCHAR(128)),
                        ''''DEFAULT''''
                    ) AS NVARCHAR(128)
                ) AS InstanceName,

                CAST(SERVERPROPERTY(''''ProductVersion'''')
                    AS NVARCHAR(50)) AS ProductVersion,

                CAST(SERVERPROPERTY(''''ProductLevel'''')
                    AS NVARCHAR(50)) AS ProductLevel,

                CAST(SERVERPROPERTY(''''ProductUpdateLevel'''')
                    AS NVARCHAR(50)) AS CU_Level,

                CAST(SERVERPROPERTY(''''ProductUpdateReference'''')
                    AS NVARCHAR(50)) AS KB_Number,

                CAST(SERVERPROPERTY(''''ProductBuildType'''')
                    AS NVARCHAR(50)) AS BuildType,

                CAST(SERVERPROPERTY(''''Edition'''')
                    AS NVARCHAR(200)) AS Edition
        '');
        ';

        EXEC sp_executesql @SQL;

    END TRY

    BEGIN CATCH

        INSERT INTO #SQLVersions
        (
            LinkedServerName,
            Status
        )
        VALUES
        (
            @LinkedServer,
            ERROR_MESSAGE()
        );

    END CATCH;

    FETCH NEXT FROM cur INTO @LinkedServer;
END;

CLOSE cur;
DEALLOCATE cur;

SELECT
    LinkedServerName,
    ServerName,
    InstanceName,
    SQLVersion,
    ProductVersion,
    ProductLevel,
    CU_Level,
    KB_Number,
    BuildType,
    Edition,
    Status
FROM #SQLVersions
ORDER BY LinkedServerName;
