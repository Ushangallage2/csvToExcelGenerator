package com.example;

/**
 * Non-Application entry point so JavaFX can boot from a fat jar / jpackage image.
 * Extending javafx.application.Application as the JVM main class causes:
 * "JavaFX runtime components are missing, and are required to run this application"
 */
public final class Launcher {
    private Launcher() {}

    public static void main(String[] args) {
        CSVProcessorApp.main(args);
    }
}
