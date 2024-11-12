FROM mcr.microsoft.com/dotnet/sdk:8.0-alpine AS base
WORKDIR /app
COPY *.csproj .
RUN dotnet restore

FROM base AS build
WORKDIR /app
COPY . .
RUN dotnet build -c Release --no-restore

FROM build AS publish
WORKDIR /app
RUN dotnet publish -c Release -o published --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:8.0-alpine AS final
EXPOSE 5000
WORKDIR /app
COPY --from=publish /app/published /app
ENTRYPOINT ["dotnet", "Contacts.dll"]

FROM mcr.microsoft.com/dotnet/sdk:8.0-alpine AS develop
WORKDIR app/
COPY . .
ENTRYPOINT ["dotnet", "run", "--no-launch-profile"]