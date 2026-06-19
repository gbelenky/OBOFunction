# SharePointMcp — multi-stage build for Azure Container Apps
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ["SharePointMcp.csproj", "./"]
RUN dotnet restore "SharePointMcp.csproj"
COPY . .
RUN dotnet publish "SharePointMcp.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app
COPY --from=build /app/publish .
# Container Apps injects PORT; default to 8089 for parity with local dev.
ENV PORT=8089
EXPOSE 8089
ENTRYPOINT ["dotnet", "SharePointMcp.dll"]
