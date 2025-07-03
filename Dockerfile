# ---------- Build stage ----------
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# Copy file csproj và khôi phục gói
COPY ToolsCTC.csproj ./
RUN dotnet restore

# Copy toàn bộ source code vào container
COPY . ./
RUN dotnet publish -c Release -o /out

# ---------- Runtime stage ----------
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /out ./

ENTRYPOINT ["dotnet", "ToolsCTC.dll"]
