# syntax=docker/dockerfile:1
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src

# Copy project files and restore
COPY ["ExcelDataEntryCore/ExcelDataEntryCore.csproj", "ExcelDataEntryCore/"]
COPY ["ExcelDataEntryWeb/ExcelDataEntryWeb.csproj", "ExcelDataEntryWeb/"]
RUN dotnet restore "ExcelDataEntryWeb/ExcelDataEntryWeb.csproj"

# Copy full source and publish
COPY . .
WORKDIR "/src/ExcelDataEntryWeb"
RUN dotnet publish "ExcelDataEntryWeb.csproj" -c Release -o /app/publish /p:UseAppHost=false

# Runtime image
FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS final
WORKDIR /app
EXPOSE 5000
ENV ASPNETCORE_URLS=http://+:5000
COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "ExcelDataEntryWeb.dll"]
