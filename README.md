# Excel to PDF Generator (Docker Version)

This is a Dockerized version of the Excel to PDF generator application. It includes LibreOffice and Japanese fonts to ensure accurate PDF conversion and rendering.

## Prerequisites

- Docker installed on your machine.

## Build the Image

Run the following command in this directory to build the Docker image:

```bash
docker build -t nodenberg .
```

## Run the Container

Run the following command to start the application:

```bash
docker run -d -p 3000:3000 --name nodenberg-app nodenberg
```

The application will be available at `http://localhost:3000`.

## Features

- **LibreOffice Included**: The Docker image comes with LibreOffice pre-installed, so you don't need to install it on your host machine.
- **Japanese Fonts**: `fonts-noto-cjk` is included to support Japanese characters in PDF generation.
