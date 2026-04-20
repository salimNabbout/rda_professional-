document.addEventListener("DOMContentLoaded", () => {
  console.log("RDA profissional carregado.");
});

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('/static/sw.js')
    .then(() => console.log('Service Worker registrado.'))
    .catch(err => console.warn('SW falhou:', err));
}
