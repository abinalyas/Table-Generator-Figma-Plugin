// Simple utility functions - no complex dependencies

/**
 * Show a status message to the user
 */
export function showMessage(message: string, type: 'success' | 'error' | 'warning' | 'info' = 'info') {
  const statusMessage = document.getElementById('statusMessage');
  if (!statusMessage) return;
  
  statusMessage.textContent = message;
  statusMessage.className = `status ${type}`;
  statusMessage.style.display = 'block';
  
  requestAnimationFrame(() => {
    statusMessage.classList.add('show');
  });
  
  // Auto-hide success messages
  if (type === 'success') {
    setTimeout(() => {
      hideMessage();
    }, 5000);
  }
}

/**
 * Hide the status message
 */
export function hideMessage() {
  const statusMessage = document.getElementById('statusMessage');
  if (!statusMessage) return;
  
  statusMessage.classList.remove('show');
  setTimeout(() => {
    statusMessage.style.display = 'none';
  }, 250);
}

/**
 * Show loading indicator
 */
export function showLoader(message: string = 'Loading...') {
  const loader = document.getElementById('loader');
  const loaderText = document.getElementById('loaderText');
  
  if (!loader || !loaderText) return;
  
  loader.style.display = 'flex';
  loaderText.textContent = message;
  
  requestAnimationFrame(() => {
    loader.classList.add('show');
  });
}

/**
 * Hide loading indicator
 */
export function hideLoader() {
  const loader = document.getElementById('loader');
  const loaderText = document.getElementById('loaderText');
  
  if (!loader || !loaderText) return;
  
  loader.classList.remove('show');
  setTimeout(() => {
    loader.style.display = 'none';
    loaderText.textContent = '';
  }, 250);
}