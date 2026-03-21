# Stage 1: Build
FROM gradle:8.5-jdk17 AS build
WORKDIR /app
COPY build.gradle settings.gradle ./
COPY src ./src
RUN gradle bootJar --no-daemon -x test

# Stage 2: Run
FROM eclipse-temurin:17-jre-alpine
WORKDIR /app

# Install fonts for JFreeChart rendering
RUN apk add --no-cache fontconfig ttf-dejavu

COPY --from=build /app/build/libs/medai-renderer.jar app.jar

# Railway sets PORT env var
EXPOSE 8080

ENTRYPOINT ["java", "-Xmx512m", "-jar", "app.jar"]
