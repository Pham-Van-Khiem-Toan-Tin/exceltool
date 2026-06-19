# Bước 1: Sử dụng Maven để build file JAR từ mã nguồn
FROM maven:3.9.6-eclipse-temurin-17 AS build
WORKDIR /app
COPY . .
RUN mvn clean package -DskipTests

# Bước 2: Sử dụng OpenJDK siêu nhẹ để chạy ứng dụng
FROM eclipse-temurin:17-jre-alpine
WORKDIR /app
COPY --from=build /app/target/*.jar app.jar

# Mở cổng 8080 (cổng mặc định của Spring Boot)
EXPOSE 8080

# Chạy ứng dụng kèm cấu hình tối ưu RAM cho gói Free của Render
ENTRYPOINT ["java", "-Xmx300m", "-Xss512k", "-jar", "app.jar"]