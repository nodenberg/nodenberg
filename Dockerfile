# Nodenberg - Excel Report Generator

# Base image
FROM node:20-bullseye

# Install LibreOffice, Japanese fonts, and curl for PDF generation and healthcheck
RUN apt-get update && apt-get install -y --no-install-recommends --fix-missing \
    libreoffice-calc \
    libreoffice-l10n-ja \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    curl \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files
COPY package.json package-lock.json ./

# Install dependencies
RUN npm ci

# Copy source code
COPY . .

# Build the application
RUN npm run build

# Initialize LibreOffice user profile to prevent first-run issues
RUN node scripts/init-libreoffice.js || true

# Expose port
EXPOSE 3000

# Start the application
CMD ["npm", "start"]
