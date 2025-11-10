# Dockerfile — SmartEmailVerifier (2-Stage Email Verification)
# Base Image: Node.js 18 (Slim + Debian)
FROM node:18-slim

# Set working directory
WORKDIR /app

# Copy package files
COPY package.json package-lock.json ./

# Install only production dependencies
RUN npm ci --omit=dev

# Copy source code
COPY . .

# Create input/output directory
RUN mkdir -p /data

# Expose nothing (CLI tool)

# Default command: Run Stage 1 + Stage 2
CMD ["sh", "-c", "node finalfast.js && node smartRetryDoubleCheck.js"]

# Optional: Override with custom input
# docker run -v $(pwd)/data:/data smartemailverifier node finalfast.js