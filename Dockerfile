# syntax=docker/dockerfile:1

FROM maven:3.9.9-eclipse-temurin-21 AS build
WORKDIR /workspace

# Copy Maven descriptor first to leverage Docker layer caching
COPY pom.xml .
RUN mvn -B dependency:go-offline

# Copy sources and build the runnable jar
COPY src ./src
RUN mvn -B clean package -DskipTests

FROM eclipse-temurin:21-jre
WORKDIR /app

COPY --from=build /workspace/target/dms-ptit-java-1.0-SNAPSHOT.jar app.jar

EXPOSE 8081
ENV PORT=8081

CMD ["sh", "-c", "java -jar app.jar --server.port=${PORT:-8081}"]
