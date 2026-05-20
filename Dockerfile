# Nodenberg - Excel Report Generator

ARG BASE_IMAGE=node:20-bookworm@sha256:abc255963bb4311b1f81bf45f2382df39a4041d2a975b1807d05809f1bc21bbe
FROM ${BASE_IMAGE}

ARG DEBIAN_SNAPSHOT=20260324T000000Z
ARG LO_VERSION=26.2.1
ARG LO_BASE_URL=https://download.documentfoundation.org/libreoffice/stable/${LO_VERSION}/deb/x86_64
ARG LO_MAIN_ARCHIVE=LibreOffice_${LO_VERSION}_Linux_x86-64_deb.tar.gz
ARG LO_MAIN_SHA256=9807363c8fabf79fc3562f606ce1673c4b29c3c53baca513f84e762845a093b2
ARG LO_LANGPACK_JA_ARCHIVE=LibreOffice_${LO_VERSION}_Linux_x86-64_deb_langpack_ja.tar.gz
ARG LO_LANGPACK_JA_SHA256=bf2f212eac17226fa15f6ac7553d01ab6f6a13a9309b8fb320dafff372e43811

# Pin Debian package sources to a snapshot for reproducible dependency resolution.
RUN set -eux; \
    rm -f /etc/apt/sources.list /etc/apt/sources.list.d/debian.sources; \
    printf 'Acquire::Check-Valid-Until "false";\n' > /etc/apt/apt.conf.d/99snapshot; \
    printf 'deb [check-valid-until=no] http://snapshot.debian.org/archive/debian/%s bookworm main\n' "${DEBIAN_SNAPSHOT}" > /etc/apt/sources.list; \
    printf 'deb [check-valid-until=no] http://snapshot.debian.org/archive/debian/%s bookworm-updates main\n' "${DEBIAN_SNAPSHOT}" >> /etc/apt/sources.list; \
    printf 'deb [check-valid-until=no] http://snapshot.debian.org/archive/debian-security/%s bookworm-security main\n' "${DEBIAN_SNAPSHOT}" >> /etc/apt/sources.list

# Install runtime dependencies from the same snapshot.
RUN apt-get update && apt-get install -y --no-install-recommends --fix-missing \
    ca-certificates \
    curl \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    libcairo2 \
    libcups2 \
    libdbus-1-3 \
    libfontconfig1 \
    libfreetype6 \
    libgbm1 \
    libglib2.0-0 \
    libgtk-3-0 \
    libnss3 \
    libsm6 \
    libx11-6 \
    libx11-xcb1 \
    libxdamage1 \
    libxext6 \
    libxi6 \
    libxinerama1 \
    libxrandr2 \
    libxrender1 \
    libxslt1.1 \
    libxt6 \
    xz-utils \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install official LibreOffice archives with SHA-256 verification.
RUN set -eux; \
    lo_series="$(printf '%s' "${LO_VERSION}" | cut -d. -f1-2)"; \
    mkdir -p /tmp/libreoffice/main /tmp/libreoffice/langpack; \
    curl -fsSL -o "/tmp/${LO_MAIN_ARCHIVE}" "${LO_BASE_URL}/${LO_MAIN_ARCHIVE}"; \
    echo "${LO_MAIN_SHA256}  /tmp/${LO_MAIN_ARCHIVE}" | sha256sum -c -; \
    tar -xzf "/tmp/${LO_MAIN_ARCHIVE}" -C /tmp/libreoffice/main --strip-components=1; \
    curl -fsSL -o "/tmp/${LO_LANGPACK_JA_ARCHIVE}" "${LO_BASE_URL}/${LO_LANGPACK_JA_ARCHIVE}"; \
    echo "${LO_LANGPACK_JA_SHA256}  /tmp/${LO_LANGPACK_JA_ARCHIVE}" | sha256sum -c -; \
    tar -xzf "/tmp/${LO_LANGPACK_JA_ARCHIVE}" -C /tmp/libreoffice/langpack --strip-components=1; \
    dpkg -i /tmp/libreoffice/main/DEBS/*.deb /tmp/libreoffice/langpack/DEBS/*.deb || true; \
    apt-get update; \
    apt-get install -y -f --no-install-recommends; \
    ln -sf "/opt/libreoffice${lo_series}/program/soffice" /usr/local/bin/soffice; \
    soffice --version; \
    apt-get clean; \
    rm -rf /var/lib/apt/lists/* /tmp/libreoffice /tmp/${LO_MAIN_ARCHIVE} /tmp/${LO_LANGPACK_JA_ARCHIVE}

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
