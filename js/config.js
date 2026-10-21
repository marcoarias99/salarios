// Datos de la conexión con Microsoft 365. Se obtienen al registrar el sitio en Entra ID
// (App registrations → "Boletas de pago" → Overview). No son contraseñas: se pueden guardar aquí.
window.CONFIG_MICROSOFT = {
  // Application (client) ID de "Boletas de pago"
  clientId: '157e6d0c-2670-4c5e-8562-ebab74f3600d',
  // Directory (tenant) ID
  tenantId: '828971ca-4f6e-45ca-babc-10e584888f50',
  // Respaldo en SharePoint. Mientras "sitio" esté vacío, el respaldo se descarga en la computadora.
  sharepoint: {
    // Dirección del sitio Administración
    sitio: 'https://zodiac353.sharepoint.com/sites/administracion',
    // Carpeta dentro de "Documents" donde se guardan los respaldos
    carpeta: 'Salarios/Boletas de pago',
    // Lista con una fila por boleta enviada (quincena en Title, Cedula, Nombre, Correo, Liquido, Fecha, EnviadoPor)
    lista: 'Envíos de boletas',
  },
};
