
# Base image
FROM node:20-bullseye

# Install LibreOffice and Japanese fonts
RUN apt-get update && apt-get install -y --no-install-recommends --fix-missing \
    libreoffice-calc \
    libreoffice-l10n-ja \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
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

# Expose port
EXPOSE 3000

# Start the application
CMD ["npm", "start"]
