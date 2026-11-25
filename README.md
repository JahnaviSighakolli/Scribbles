SELECT *
FROM OPENQUERY([YOUR_LINKED_SERVER],
'
    SELECT
        SERVERPROPERTY(''ProductVersion'') AS ProductVersion,
        SERVERPROPERTY(''ProductUpdateLevel'') AS CU_Level,
        SERVERPROPERTY(''ProductUpdateReference'') AS KB_Number,
        SERVERPROPERTY(''ProductBuildType'') AS BuildType
');
