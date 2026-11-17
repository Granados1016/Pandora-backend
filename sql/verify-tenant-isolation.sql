/* ============================================================================
   #3 Fase 2 — Verificación de aislamiento multi-tenant (Pandora)

   Recorre DINÁMICAMENTE todas las tablas base con columna TenantId y verifica
   que ninguna fila esté:
     - huérfana  (TenantId NULL)  -> sería invisible a los filtros por tenant
                                      o podría filtrarse entre clientes.
     - colgada   (TenantId que no existe en dbo.Tenants) -> apunta a un cliente
                                      borrado/inexistente.
   Excepción legítima: en dbo.AppUsers, las cuentas con Role = 'SuperAdmin'
   (el proveedor) no tienen tenant por diseño y NO cuentan como huérfanas.

   Uso (mismo resultado en local, staging o producción):
     sqlcmd -S <server> -d <PandoraDB> -b -i sql/verify-tenant-isolation.sql
   Termina con error (RAISERROR, nivel 16) si encuentra algún problema, para
   que un pipeline de CI o un despliegue lo detecte. Solo lee; no modifica nada.
   ============================================================================ */
SET NOCOUNT ON;

IF OBJECT_ID('tempdb..#iso') IS NOT NULL DROP TABLE #iso;
CREATE TABLE #iso (
    TableName        SYSNAME,
    Total            BIGINT,
    NullTenantId     BIGINT,
    DanglingTenantId BIGINT
);

DECLARE @sql NVARCHAR(MAX) = N'';

SELECT @sql = @sql +
    CASE WHEN @sql = N'' THEN N'' ELSE N' UNION ALL ' END +
    N'SELECT ' + QUOTENAME(c.TABLE_SCHEMA + N'.' + c.TABLE_NAME, '''') + N' AS TableName, ' +
    N'COUNT_BIG(*) AS Total, ' +
    N'ISNULL(SUM(CASE WHEN src.TenantId IS NULL' +
        /* En AppUsers se excluyen las cuentas SuperAdmin (globales por diseño). */
        CASE WHEN c.TABLE_NAME = N'AppUsers'
                  AND EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS rc
                              WHERE rc.TABLE_SCHEMA = c.TABLE_SCHEMA
                                AND rc.TABLE_NAME   = c.TABLE_NAME
                                AND rc.COLUMN_NAME  = N'Role')
             THEN N' AND ISNULL(src.Role,'''') <> ''SuperAdmin'''
             ELSE N'' END +
    N' THEN 1 ELSE 0 END), 0) AS NullTenantId, ' +
    /* "Colgado" vía LEFT JOIN (t.Id NULL): no se puede usar subquery dentro de SUM. */
    N'ISNULL(SUM(CASE WHEN src.TenantId IS NOT NULL AND t.Id IS NULL ' +
        N'THEN 1 ELSE 0 END), 0) AS DanglingTenantId ' +
    N'FROM ' + QUOTENAME(c.TABLE_SCHEMA) + N'.' + QUOTENAME(c.TABLE_NAME) + N' src ' +
    N'LEFT JOIN dbo.Tenants t ON t.Id = src.TenantId'
FROM INFORMATION_SCHEMA.COLUMNS c
JOIN INFORMATION_SCHEMA.TABLES t
  ON t.TABLE_SCHEMA = c.TABLE_SCHEMA AND t.TABLE_NAME = c.TABLE_NAME
 AND t.TABLE_TYPE = 'BASE TABLE'
WHERE c.COLUMN_NAME = 'TenantId';

IF @sql = N''
BEGIN
    PRINT 'No se encontraron tablas con columna TenantId. ¿Base de datos vacía o esquema incorrecto?';
    RETURN;
END

INSERT INTO #iso (TableName, Total, NullTenantId, DanglingTenantId)
EXEC sys.sp_executesql @sql;

-- Reporte completo (ordenado: primero lo problemático).
SELECT TableName, Total, NullTenantId, DanglingTenantId,
       CASE WHEN NullTenantId = 0 AND DanglingTenantId = 0 THEN 'OK' ELSE 'REVISAR' END AS Estado
FROM #iso
ORDER BY (NullTenantId + DanglingTenantId) DESC, TableName;

DECLARE @problemTables INT = (SELECT COUNT(*) FROM #iso WHERE NullTenantId > 0 OR DanglingTenantId > 0);
DECLARE @problemRows   BIGINT = (SELECT ISNULL(SUM(NullTenantId + DanglingTenantId), 0) FROM #iso);
DECLARE @tablesChecked INT = (SELECT COUNT(*) FROM #iso);

PRINT '----------------------------------------------------------------';
PRINT 'Tablas con TenantId revisadas: ' + CAST(@tablesChecked AS VARCHAR(10));

IF @problemRows > 0
BEGIN
    PRINT 'AISLAMIENTO: ' + CAST(@problemRows AS VARCHAR(20)) + ' fila(s) problemática(s) en '
        + CAST(@problemTables AS VARCHAR(10)) + ' tabla(s). Ver columna Estado = REVISAR.';
    DROP TABLE #iso;
    RAISERROR('Verificacion de aislamiento multi-tenant FALLIDA: hay filas con TenantId NULL o colgado.', 16, 1);
END
ELSE
BEGIN
    PRINT 'AISLAMIENTO OK: ninguna fila huérfana ni colgada en las tablas con TenantId.';
    DROP TABLE #iso;
END
