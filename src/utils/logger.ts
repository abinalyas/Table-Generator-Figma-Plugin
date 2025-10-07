// Simple logger utility to reduce verbose console.log statements
export class Logger {
    private static isDebugMode = false; // Set to true for detailed logging

    static debug(message: string, ...args: any[]): void {
        if (this.isDebugMode) {
            console.log(message, ...args);
        }
    }

    static info(message: string, ...args: any[]): void {
        console.log(message, ...args);
    }

    static warn(message: string, ...args: any[]): void {
        console.warn(message, ...args);
    }

    static error(message: string, ...args: any[]): void {
        console.error(message, ...args);
    }

    static success(message: string, ...args: any[]): void {
        console.log(`✅ ${message}`, ...args);
    }

    static progress(message: string, ...args: any[]): void {
        console.log(`🔄 ${message}`, ...args);
    }

    static warning(message: string, ...args: any[]): void {
        console.log(`⚠️ ${message}`, ...args);
    }
}