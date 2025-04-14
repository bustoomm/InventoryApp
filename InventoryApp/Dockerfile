# Base image for .NET runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 80

# Base image for building
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY . .

# Ganti dengan nama project lu yang .csproj
RUN dotnet restore "InventoryApp/InventoryApp.csproj"
RUN dotnet publish "InventoryApp/InventoryApp.csproj" -c Release -o /app/publish

# Final stage
FROM base AS final
WORKDIR /app
COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "InventoryApp.dll"]
