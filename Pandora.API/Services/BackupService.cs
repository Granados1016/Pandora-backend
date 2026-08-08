using System.Text;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Services;

/// <summary>
/// Genera el respaldo completo de la base de datos, reutilizado tanto por el
/// botón manual del panel Admin como por <see cref="AutomatedBackupWorker"/>.
/// Intenta un BACKUP DATABASE nativo (.bak); si el motor no lo soporta
/// (LocalDB, permisos limitados, etc.) cae a un .sql autocontenible que
/// incluye CREATE TABLE + INSERT para cada tabla (incluso las vacías).
/// </summary>
public class BackupService(IConfiguration config, IWebHostEnvironment env, ILogger<BackupService> logger)
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    public record Result(byte[] Bytes, string FileName, string Method);

    /// <summary>Automático: intenta .bak nativo, cae a .sql si el motor no lo soporta.</summary>
    public async Task<Result> GenerateAsync(CancellationToken ct)
    {
        try
        {
            return await GenerateBakAsync(ct);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex,
                "BACKUP DATABASE no disponible ({Msg}). Generando respaldo SQL ejecutable como fallback.",
                ex.Message);
            return await GenerateSqlAsync(ct);
        }
    }

    /// <summary>Fuerza el .bak nativo de SQL Server. Lanza excepción si el motor no lo soporta (ej. permisos limitados).</summary>
    public async Task<Result> GenerateBakAsync(CancellationToken ct)
    {
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        string database;
        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT DB_NAME()";
            database = (string)(await cmd.ExecuteScalarAsync(ct) ?? "PandoraDB");
        }

        var backupDir  = Path.Combine(env.ContentRootPath, "backups");
        Directory.CreateDirectory(backupDir);
        var backupFile = Path.Combine(backupDir, $"PandoraDB_{timestamp}.bak");
        var sqlPath    = backupFile.Replace("\\", "/");

        try
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandTimeout = 300;
            cmd.CommandText = $"""
                BACKUP DATABASE [{database}]
                TO DISK = N'{sqlPath.Replace("'", "''")}'
                WITH FORMAT, INIT,
                     NAME = N'Pandora Full Backup {timestamp}',
                     SKIP, NOREWIND, NOUNLOAD, STATS = 10
                """;

            await cmd.ExecuteNonQueryAsync(ct);

            if (!System.IO.File.Exists(backupFile))
                throw new FileNotFoundException("SQL Server no generó el archivo .bak en la ruta esperada.", backupFile);

            var bytes = await System.IO.File.ReadAllBytesAsync(backupFile, ct);
            System.IO.File.Delete(backupFile);

            logger.LogInformation("Backup .bak generado: {File} ({Kb} KB)",
                $"PandoraDB_{timestamp}.bak", bytes.Length / 1024);

            return new Result(bytes, $"PandoraDB_{timestamp}.bak", "bak");
        }
        catch
        {
            if (System.IO.File.Exists(backupFile))
                try { System.IO.File.Delete(backupFile); } catch { /* ignorar */ }
            throw;
        }
    }

    /// <summary>Fuerza el .sql autocontenible (CREATE TABLE + INSERT), sin intentar el .bak nativo primero.</summary>
    public async Task<Result> GenerateSqlAsync(CancellationToken ct)
    {
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        string database;
        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT DB_NAME()";
            database = (string)(await cmd.ExecuteScalarAsync(ct) ?? "PandoraDB");
        }

        var (bytes, fileName) = await GenerateSqlBackupAsync(conn, database, timestamp, ct);
        return new Result(bytes, fileName, "sql");
    }

    // ── Genera un .sql autocontenible (CREATE TABLE + INSERT) ────────────────
    private async Task<(byte[] Bytes, string FileName)> GenerateSqlBackupAsync(
        SqlConnection conn, string database, string timestamp, CancellationToken ct)
    {
        var sb = new StringBuilder();
        sb.AppendLine("-- ============================================================");
        sb.AppendLine($"-- Pandora DB Backup (SQL Script)");
        sb.AppendLine($"-- Base de datos: {database}");
        sb.AppendLine($"-- Generado: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine("-- Autocontenible: incluye CREATE TABLE + datos para cada tabla,");
        sb.AppendLine("-- incluso si está vacía. No replica FKs, índices ni identidades.");
        sb.AppendLine("-- ============================================================");
        sb.AppendLine();
        sb.AppendLine("SET NOCOUNT ON;");
        sb.AppendLine("BEGIN TRANSACTION;");
        sb.AppendLine();

        var tables = new List<string>();
        await using (var tCmd = conn.CreateCommand())
        {
            tCmd.CommandText = """
                SELECT TABLE_SCHEMA + '.' + TABLE_NAME
                FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_TYPE = 'BASE TABLE'
                ORDER BY TABLE_NAME
                """;
            await using var tReader = await tCmd.ExecuteReaderAsync(ct);
            while (await tReader.ReadAsync(ct))
                tables.Add(tReader.GetString(0));
        }

        foreach (var table in tables)
        {
            var parts    = table.Split('.');
            var safeName = string.Join(".", parts.Select(p => $"[{p}]"));

            sb.AppendLine($"-- ── {table} ──────────────────────────────");

            // CREATE TABLE (si no existe ya en el destino) a partir de INFORMATION_SCHEMA.COLUMNS
            try
            {
                var ddl = await BuildCreateTableAsync(conn, parts[0], parts[1], ct);
                sb.AppendLine($"IF OBJECT_ID(N'{table}', N'U') IS NULL");
                sb.AppendLine("BEGIN");
                sb.AppendLine(ddl);
                sb.AppendLine("END");
                sb.AppendLine();
            }
            catch (Exception ex)
            {
                sb.AppendLine($"-- ADVERTENCIA: no se pudo generar CREATE TABLE de {table}: {ex.Message}");
            }

            try
            {
                // Columna(s) de PK — permiten envolver cada INSERT en un
                // "IF NOT EXISTS" para que el script sea idempotente: al
                // restaurarlo solo repone filas que falten, sin duplicar ni
                // pisar lo que ya exista. Sin PK detectada, se inserta directo
                // (igual que antes) — sin forma segura de deduplicar.
                var pkCols = await GetPrimaryKeyColumnsAsync(conn, parts[0], parts[1], ct);

                await using var dCmd = conn.CreateCommand();
                dCmd.CommandText = $"SELECT * FROM {safeName}";
                await using var r = await dCmd.ExecuteReaderAsync(ct);

                var colNames = Enumerable.Range(0, r.FieldCount).Select(r.GetName).ToList();
                var cols     = colNames.Select(c => $"[{c}]").ToList();

                var rowCount = 0;
                while (await r.ReadAsync(ct))
                {
                    var vals = Enumerable.Range(0, r.FieldCount).Select(i =>
                    {
                        if (r.IsDBNull(i)) return "NULL";
                        return r.GetValue(i) switch
                        {
                            string s           => $"N'{s.Replace("'", "''")}'",
                            DateTime d         => $"'{d:yyyy-MM-ddTHH:mm:ss.fff}'",
                            DateTimeOffset dto => $"'{dto:yyyy-MM-ddTHH:mm:ss.fffzzz}'",
                            bool b             => b ? "1" : "0",
                            Guid g             => $"'{g}'",
                            byte[] arr         => $"0x{Convert.ToHexString(arr)}",
                            var v              => v?.ToString() ?? "NULL",
                        };
                    }).ToList();

                    var insertStmt =
                        $"INSERT INTO {table} ({string.Join(", ", cols)}) " +
                        $"VALUES ({string.Join(", ", vals)});";

                    if (pkCols.Count > 0)
                    {
                        var guard = string.Join(" AND ", pkCols.Select(pk =>
                        {
                            var idx = colNames.IndexOf(pk);
                            return idx < 0 ? "1=1" : $"[{pk}] = {vals[idx]}";
                        }));
                        sb.AppendLine($"IF NOT EXISTS (SELECT 1 FROM {table} WHERE {guard}) {insertStmt}");
                    }
                    else
                    {
                        sb.AppendLine(insertStmt);
                    }
                    rowCount++;
                }

                sb.AppendLine(rowCount == 0
                    ? $"-- (0 filas en {table})"
                    : $"-- ({rowCount} fila{(rowCount == 1 ? "" : "s")} en {table})");
                sb.AppendLine();
            }
            catch (Exception ex)
            {
                sb.AppendLine($"-- ADVERTENCIA: no se pudo exportar datos de {table}: {ex.Message}");
                sb.AppendLine();
            }
        }

        sb.AppendLine("COMMIT TRANSACTION;");
        sb.AppendLine($"-- Fin del respaldo — {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

        logger.LogInformation("Respaldo SQL autocontenible generado a las {Time} ({Tables} tablas)",
            DateTime.Now, tables.Count);

        var bytes = Encoding.UTF8.GetPreamble()
            .Concat(Encoding.UTF8.GetBytes(sb.ToString()))
            .ToArray();

        return (bytes, $"PandoraDB_{timestamp}.sql");
    }

    // ── Genera un CREATE TABLE simplificado a partir de INFORMATION_SCHEMA ───
    private static async Task<string> BuildCreateTableAsync(
        SqlConnection conn, string schema, string tableName, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH,
                   NUMERIC_PRECISION, NUMERIC_SCALE, IS_NULLABLE
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = @Schema AND TABLE_NAME = @Table
            ORDER BY ORDINAL_POSITION
            """;
        cmd.Parameters.AddWithValue("@Schema", schema);
        cmd.Parameters.AddWithValue("@Table", tableName);

        var lines = new List<string>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            var col      = r.GetString(0);
            var dataType = r.GetString(1);
            var maxLen   = r.IsDBNull(2) ? (int?)null : r.GetInt32(2);

            var typeSql = dataType switch
            {
                "nvarchar" or "varchar" or "nchar" or "char" =>
                    maxLen is null or -1 ? $"{dataType}(MAX)" : $"{dataType}({maxLen})",
                "decimal" or "numeric" => $"{dataType}(18,2)",
                _ => dataType
            };

            // Siempre NULL: esta reconstrucción es un fallback simplificado
            // (sin PKs/FKs/identity/defaults), forzar NOT NULL aquí solo
            // arriesgaría que el INSERT falle sin aportar fidelidad real.
            lines.Add($"    [{col}] {typeSql} NULL");
        }

        return $"    CREATE TABLE [{schema}].[{tableName}] (\n{string.Join(",\n", lines)}\n    );";
    }

    // ── Columna(s) de PRIMARY KEY de una tabla (vacío si no tiene) ────────────
    private static async Task<List<string>> GetPrimaryKeyColumnsAsync(
        SqlConnection conn, string schema, string tableName, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT kcu.COLUMN_NAME
            FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
            JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu
              ON tc.CONSTRAINT_NAME = kcu.CONSTRAINT_NAME AND tc.TABLE_SCHEMA = kcu.TABLE_SCHEMA
            WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
              AND tc.TABLE_SCHEMA = @Schema AND tc.TABLE_NAME = @Table
            ORDER BY kcu.ORDINAL_POSITION
            """;
        cmd.Parameters.AddWithValue("@Schema", schema);
        cmd.Parameters.AddWithValue("@Table", tableName);

        var cols = new List<string>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
            cols.Add(r.GetString(0));
        return cols;
    }
}
