plugins {
    kotlin("jvm") version "1.9.0"
    application
}

group = "com.abomo"
version = "1.0"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi:5.2.3")
    implementation("org.apache.poi:poi-ooxml:5.2.3")
    implementation("com.itextpdf:kernel:7.2.5")
    implementation("com.itextpdf:layout:7.2.5")
}

application {
    mainClass.set("MainGUIKt")
}
