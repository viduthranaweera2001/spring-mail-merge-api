
FROM openjdk:17
LABEL authors="viduthranaweera"
WORKDIR /app
COPY target/mail-merge-api-0.0.1-SNAPSHOT.jar /app/app.jar
EXPOSE 8088
ENTRYPOINT ["java", "-jar", "app.jar","--server.port=8088"]


#./mvnw clean package