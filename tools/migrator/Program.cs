using System;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;

var connStr = Environment.GetEnvironmentVariable("ConnectionStrings__PandoraDb")
           ?? throw new Exception("No ConnectionStrings__PandoraDb");

Console.WriteLine($"Conectando a BD...");
await using var conn = new SqlConnection(connStr);
await conn.OpenAsync();
Console.WriteLine("Conexión OK. Ejecutando migración...");

await using var cmd = conn.CreateCommand();
cmd.CommandText = """
    -- 1. Limpiar columna temporal de intentos previos
    IF EXISTS (SELECT 1 FROM sys.columns
               WHERE object_id = OBJECT_ID('dbo.InventoryItems') AND name = 'StatusText')
        ALTER TABLE dbo.InventoryItems DROP COLUMN StatusText;

    -- 2. Migrar Status INT → NVARCHAR solo si sigue siendo INT
    IF EXISTS (
        SELECT 1 FROM sys.columns
        WHERE object_id = OBJECT_ID('dbo.InventoryItems')
          AND name = 'Status'
          AND system_type_id = TYPE_ID('int')
    )
    BEGIN
        ALTER TABLE dbo.InventoryItems ADD StatusText NVARCHAR(50) NULL;
        UPDATE dbo.InventoryItems SET StatusText = CASE Status
            WHEN 1 THEN 'Activo' WHEN 2 THEN 'Mantenimiento'
            WHEN 3 THEN 'Dado de baja' WHEN 4 THEN 'En almacén' ELSE 'Activo' END;
        DECLARE @ddl NVARCHAR(MAX) = N'';
        SELECT @ddl += N'ALTER TABLE dbo.InventoryItems DROP CONSTRAINT [' + dc.name + N'];'
        FROM sys.default_constraints dc
        JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
        WHERE c.object_id = OBJECT_ID('dbo.InventoryItems') AND c.name = 'Status';
        SELECT @ddl += N'DROP INDEX [' + i.name + N'] ON dbo.InventoryItems;'
        FROM sys.indexes i
        JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
        JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
        WHERE c.object_id = OBJECT_ID('dbo.InventoryItems') AND c.name = 'Status'
          AND i.is_primary_key = 0 AND i.is_unique_constraint = 0;
        IF LEN(@ddl) > 0 EXEC sp_executesql @ddl;
        ALTER TABLE dbo.InventoryItems DROP COLUMN Status;
        ALTER TABLE dbo.InventoryItems ADD Status NVARCHAR(50) NULL DEFAULT 'Activo';
        UPDATE dbo.InventoryItems SET Status = ISNULL(StatusText, 'Activo');
        ALTER TABLE dbo.InventoryItems DROP COLUMN StatusText;
        PRINT 'Migración Status INT->NVARCHAR completada.';
    END
    ELSE
        PRINT 'Status ya es NVARCHAR, sin cambios.';

    -- 3. Agregar columnas faltantes en InventoryItems
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.InventoryItems') AND name='AssignedEmployeeId')
        ALTER TABLE dbo.InventoryItems ADD AssignedEmployeeId UNIQUEIDENTIFIER NULL;
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.InventoryItems') AND name='IsPhone')
        ALTER TABLE dbo.InventoryItems ADD IsPhone BIT NOT NULL DEFAULT 0;
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.InventoryItems') AND name='PurchaseDate')
        ALTER TABLE dbo.InventoryItems ADD PurchaseDate DATETIME2 NULL;
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.InventoryItems') AND name='PurchasePrice')
        ALTER TABLE dbo.InventoryItems ADD PurchasePrice DECIMAL(18,2) NULL;
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.InventoryItems') AND name='Accessories')
        ALTER TABLE dbo.InventoryItems ADD Accessories NVARCHAR(MAX) NULL;
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.InventoryItems') AND name='DecommissionDate')
        ALTER TABLE dbo.InventoryItems ADD DecommissionDate DATETIME2 NULL;
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.InventoryItems') AND name='DecommissionReason')
        ALTER TABLE dbo.InventoryItems ADD DecommissionReason NVARCHAR(MAX) NULL;

    -- 4. Crear tabla Employees si no existe
    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='Employees' AND schema_id=SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE dbo.Employees (
            Id UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
            FullName NVARCHAR(200) NOT NULL,
            Email NVARCHAR(200) NULL, Phone NVARCHAR(50) NULL,
            Position NVARCHAR(200) NULL, DepartmentId UNIQUEIDENTIFIER NULL,
            IsActive BIT NOT NULL DEFAULT 1,
            CreatedAt DATETIME2 NOT NULL DEFAULT GETUTCDATE(), UpdatedAt DATETIME2 NULL
        );
        PRINT 'Tabla Employees creada.';
    END
    ELSE PRINT 'Tabla Employees ya existe.';

    PRINT 'Migración completa.';
    """;
cmd.CommandTimeout = 60;
await cmd.ExecuteNonQueryAsync();
Console.WriteLine("✅ Migración ejecutada correctamente.");
