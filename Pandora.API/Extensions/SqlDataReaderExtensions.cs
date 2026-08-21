using Microsoft.Data.SqlClient;

namespace Pandora.API.Extensions;

/// <summary>
/// Extensiones para leer columnas DATETIME2 de SQL Server marcándolas como UTC.
///
/// Todas las fechas en Pandora se guardan con GETUTCDATE(), pero
/// SqlDataReader.GetDateTime() regresa DateTimeKind.Unspecified. Eso hace que
/// System.Text.Json serialice el valor sin sufijo 'Z', y el navegador
/// (new Date(...)) lo interprete como si ya fuera hora local — sin restar el
/// offset de México — desfasando el horario que ve el usuario (~6 horas).
///
/// Usar siempre GetUtcDateTime/GetUtcDateTimeOrNull en vez de GetDateTime()
/// para cualquier columna DATETIME2 que se vaya a devolver en un JSON.
/// </summary>
public static class SqlDataReaderExtensions
{
    public static DateTime GetUtcDateTime(this SqlDataReader r, int ordinal) =>
        DateTime.SpecifyKind(r.GetDateTime(ordinal), DateTimeKind.Utc);

    public static DateTime GetUtcDateTime(this SqlDataReader r, string columnName) =>
        r.GetUtcDateTime(r.GetOrdinal(columnName));

    public static DateTime? GetUtcDateTimeOrNull(this SqlDataReader r, int ordinal) =>
        r.IsDBNull(ordinal) ? null : r.GetUtcDateTime(ordinal);

    public static DateTime? GetUtcDateTimeOrNull(this SqlDataReader r, string columnName) =>
        r.GetUtcDateTimeOrNull(r.GetOrdinal(columnName));
}
