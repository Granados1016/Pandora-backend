# ── Stage 1: build + publish ──────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Fuerza rebuild completo sin caché
ARG BUILDTIME=20260429-cors
RUN echo "Build: $BUILDTIME"

# Copiar csproj y restaurar NuGet (capa cacheada mientras no cambie el csproj)
COPY Pandora.API/Pandora.API.csproj ./Pandora.API/

# Los DLLs pre-compilados deben estar donde el HintPath los busca
RUN mkdir -p ./Pandora.API/bin/Debug/net8.0/
COPY libs/ ./Pandora.API/bin/Debug/net8.0/

RUN dotnet restore ./Pandora.API/Pandora.API.csproj

# Copiar código fuente
COPY . .

# Asegurar que los DLLs de libs siguen en su lugar tras el COPY
COPY libs/ ./Pandora.API/bin/Debug/net8.0/

# Publicar — dotnet publish copia TODOS los assemblies NuGet al output
RUN dotnet publish ./Pandora.API/Pandora.API.csproj \
    -c Release \
    -o /app/publish \
    --no-restore

# ── Stage 2: runtime ──────────────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

RUN mkdir -p /app/storage /app/biblioteca \
    && groupadd --system pandora \
    && useradd --system --gid pandora --no-create-home pandora \
    && chown -R pandora:pandora /app

USER pandora

ENV ASPNETCORE_ENVIRONMENT=Production \
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false

EXPOSE 8080

COPY --from=build /app/publish .

ENTRYPOINT ["dotnet", "Pandora.API.dll"]
