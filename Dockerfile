# Giai đoạn build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# Copy csproj và restore
COPY *.sln .
COPY ToolsCTC.API/*.csproj ./ToolsCTC.API/
RUN dotnet restore

# Copy tất cả và build
COPY . .
WORKDIR /app/ToolsCTC.API
RUN dotnet publish -c Release -o /app/out

# Giai đoạn runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app/out .

ENTRYPOINT ["dotnet", "ToolsCTC.dll"]
