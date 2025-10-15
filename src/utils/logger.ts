// Simple logger utility for the plugin
export class Logger {
  static debug(message: string, ...args: any[]): void {
    console.log(`[DEBUG] ${message}`, ...args);
  }

  static info(message: string, ...args: any[]): void {
    console.log(`[INFO] ${message}`, ...args);
  }

  static progress(message: string, ...args: any[]): void {
    console.log(`[PROGRESS] ${message}`, ...args);
  }

  static success(message: string, ...args: any[]): void {
    console.log(`[SUCCESS] ${message}`, ...args);
  }

  static error(message: string, ...args: any[]): void {
    console.error(`[ERROR] ${message}`, ...args);
  }

  static warn(message: string, ...args: any[]): void {
    console.warn(`[WARN] ${message}`, ...args);
  }

  static warning(message: string, ...args: any[]): void {
    console.warn(`[WARNING] ${message}`, ...args);
  }
}
