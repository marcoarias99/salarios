// Datos de la conexión con Microsoft 365. Se obtienen al registrar el sitio en Entra ID
// (App registrations → "Boletas de pago" → Overview). No son contraseñas: se pueden guardar aquí.
window.CONFIG_MICROSOFT = {
  // Application (client) ID de "Boletas de pago"
  clientId: '157e6d0c-2670-4c5e-8562-ebab74f3600d',
  // Directory (tenant) ID
  tenantId: '828971ca-4f6e-45ca-babc-10e584888f50',
  // Respaldo en SharePoint. Mientras "sitio" esté vacío, el respaldo se descarga en la computadora.
  sharepoint: {
    // Dirección del sitio, por ejemplo 'https://zodiac.sharepoint.com/sites/Planillas'
    sitio: '',
    // Carpeta dentro de "Documentos" donde se guardan los respaldos
    carpeta: 'Boletas de pago',
  },
};
