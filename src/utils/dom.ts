// DOM utility functions - simple and safe

/**
 * Get an element by ID with type safety
 */
export function getElementById<T extends HTMLElement = HTMLElement>(id: string): T | null {
  return document.getElementById(id) as T | null;
}

/**
 * Create an element with optional attributes
 */
export function createElement<K extends keyof HTMLElementTagNameMap>(
  tagName: K,
  attributes?: Partial<HTMLElementTagNameMap[K]>,
  textContent?: string
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tagName);
  
  if (attributes) {
    Object.assign(element, attributes);
  }
  
  if (textContent) {
    element.textContent = textContent;
  }
  
  return element;
}

/**
 * Add CSS classes to an element
 */
export function addClass(element: HTMLElement, ...classes: string[]): void {
  element.classList.add(...classes);
}

/**
 * Remove CSS classes from an element
 */
export function removeClass(element: HTMLElement, ...classes: string[]): void {
  element.classList.remove(...classes);
}

/**
 * Toggle CSS class on an element
 */
export function toggleClass(element: HTMLElement, className: string): void {
  element.classList.toggle(className);
}

/**
 * Set multiple CSS styles on an element
 */
export function setStyles(element: HTMLElement, styles: Partial<CSSStyleDeclaration>): void {
  Object.assign(element.style, styles);
}