@echo off
:: C?u h\Uffffffff c?ng ch?y c?a Spring Boot (EXPOSE 8080)
SET APP_PORT=8080

echo ==========================================
echo 1. Dang cap nhat ma nguon moi nhat (Git Pull)...
echo ==========================================
git pull

echo.
echo ==========================================
echo 2. Dang quet va tat ung dung cu dang chay tren port %APP_PORT%...
echo ==========================================
:: Bu?c n\Uffffffffgi\Uffffffff?i ph\Uffffffffport 8080 tru?c khi ch?y b?n m?i
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :%APP_PORT% ^| findstr LISTENING') do (
    taskkill /f /pid %%a 2>nul
)

echo.
echo ==========================================
echo 3. Tuong duong Stage 'build': RUN mvn clean package -DskipTests
echo ==========================================
:: S? d?ng Maven Wrapper (mvnw) c\Uffffffffn trong d? \Uffffffffc?a b?n d? build
call mvnw clean package -DskipTests

echo.
echo ==========================================
echo 4. Tuong duong Stage 'runtime': Khoi chay file .jar voi cau hinh RAM
echo ==========================================
cd target
for %%i in (*.jar) do (
    set JAR_FILE=%%i
)

:: Kh?i ch?y ?ng d?ng ? m?t c?a s? CMD ri\Uffffffff (ch?y ng?m) k\Uffffffffc?u h\Uffffffff t?i uu RAM y h?t Dockerfile c?a b?n
start "SpringBoot-ExcelTool" java -Xmx300m -Xss512k -jar %JAR_FILE%

echo.
echo ==========================================
echo HOAN THANH! Ung dung dang khoi chay tren port %APP_PORT%.
echo ==========================================
pause