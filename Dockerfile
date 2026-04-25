FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime

WORKDIR /app

COPY libs/ ./
COPY linux-libs/Microsoft.Data.SqlClient.dll ./

RUN mkdir -p /app/storage /app/biblioteca

RUN groupadd --system pandora \
 && useradd --system --gid pandora --no-create-home pandora \
 && chown -R pandora:pandora /app

USER pandora

ENV ASPNETCORE_ENVIRONMENT=Production \
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false

EXPOSE 8080

ENTRYPOINT ["dotnet", "Pandora.API.dll"]
