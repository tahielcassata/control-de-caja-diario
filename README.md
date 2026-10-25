# Sistema de cobros para barberías

Web app para registrar los cobros de una barbería desde una tablet en la caja y cerrar el día sin cuentas a mano. Reemplaza la planilla de papel donde se anotan los pagos en efectivo y por transferencia.

Funciona con **Google Apps Script + Google Sheets**: no necesita servidor ni base de datos, y no tiene costos de plataforma. Los datos quedan en una planilla de Google de la cuenta del local.

> **Demo:** abrí `docs/index.html` en el navegador (o la versión publicada en GitHub Pages). Funciona en modo demostración con datos de ejemplo. PIN de administración: `1234`.

## Funciones

**Pantalla de cobro**
- Botones grandes por **puesto de trabajo** (Puesto 1, 2, 3…) y por **servicio**, con el precio cargado automáticamente.
- Medio de pago: efectivo o transferencia.
- **Propina opcional**, con su propio medio de pago (puede ser distinto al del corte).
- Totales del día siempre a la vista: efectivo, transferencias, total y propinas a entregar.
- **Deshacer último** para corregir un error de carga al instante.
- Después de guardar se desmarca todo, para que el próximo cobro arranque de cero.

**Cierre del día**
- Resumen general y por puesto: cortes, plata de servicios y propinas a entregar.
- Botón **Cerrar día**: guarda el resumen y bloquea nuevas cargas. Al otro día la app arranca en cero.
- **Cierre automático** a la hora configurada, por si nadie lo cierra. Además, al abrir la app se cierra el día anterior si quedó abierto.
- **Reabrir día** con PIN, para cargar un cobro que quedó afuera. Al volver a cerrar se recalcula.
- Navegación por días anteriores.

**Administración (con PIN)**
- Agregar puestos de trabajo y activar o desactivar los existentes (el historial no se pierde).
- Agregar servicios, cambiar nombres y precios, o desactivarlos. Los cambios de precio no tocan los cobros ni los cierres anteriores.
- Configurar la hora del cierre automático.
- El **PIN lo asigna el instalador** desde la planilla (menú *Cobros → Asignar PIN del cliente*), y desde ahí también se resetea si el cliente se lo olvida. Se guarda con hash SHA-256 y sal, y se bloquea por 5 minutos después de 5 intentos fallidos.

**Diseño adaptativo**: el mismo código se ve bien en tablet, celular y PC.

## Arquitectura

```
Tablet / celu / PC  ──►  Web app (index.html)
                              │  google.script.run
                              ▼
                         Apps Script (Code.gs)
                              │
                              ▼
                         Google Sheets
                         ├─ Cobros     una fila por cobro
                         ├─ Cierres    resumen por día y por puesto
                         ├─ Puestos    puestos de trabajo
                         └─ Servicios  servicios y precios
```

- **Concurrencia:** cada escritura usa `LockService` para que dos cargas simultáneas no se pisen.
- **Rendimiento:** la hoja de cobros se lee de abajo hacia arriba en bloques y se corta al llegar a fechas anteriores. Con 360 cobros por día y un año de historial (más de 100.000 filas) la pantalla sigue respondiendo rápido.
- **Protección:** las hojas Cobros y Cierres tienen protección con aviso, para que no se editen sin querer.
- **Seguridad de acceso:** la web app se publica con acceso "Solo yo" y se usa con la tablet logueada en la cuenta del local. Las funciones de administración piden PIN.

## Estructura

```
src/
  Code.gs           backend (Apps Script)
  index.html        pantallas: cobro, cierre y administración (+ modo demo)
  appsscript.json   configuración del proyecto y de la web app
test/
  simulador.js      simulador de Apps Script para correr el backend en Node
  test.js           pruebas del backend
docs/
  index.html        copia para la demo en GitHub Pages
```

## Pruebas

El backend se prueba en Node con un simulador de las APIs de Google:

```bash
npm test
```

Cubre: carga y validación de cobros, propinas en distinto medio, deshacer, cierre manual y automático, reapertura con PIN, cambio de día, administración, bloqueo de PIN y rendimiento con un año de historial.

## Instalación para un cliente

Requisitos: Node.js y [clasp](https://github.com/google/clasp) (`npm install -g @google/clasp`).

1. Crear un Gmail para el local (o usar el que tengan).
2. En [script.google.com/home/usersettings](https://script.google.com/home/usersettings) activar la **API de Google Apps Script**.
3. Iniciar sesión con esa cuenta:
   ```bash
   clasp login
   ```
4. Crear la planilla con su proyecto de Apps Script:
   ```bash
   clasp create --type sheets --title "Cobros - Nombre del local" --rootDir src
   ```
   Esto crea `.clasp.json` (no se sube a Git: cada cliente tiene el suyo).
   clasp puede reemplazar `src/appsscript.json` por uno vacío; recuperá el del repo:
   ```bash
   git checkout -- src/appsscript.json
   ```
5. Subir el código:
   ```bash
   clasp push
   ```
6. Abrir el editor (`clasp open`), elegir la función `setup` y ejecutarla. Autorizar los permisos. Crea las hojas, 10 puestos, servicios de ejemplo y el cierre automático.
7. Volver a abrir la planilla y en el menú **Cobros → Asignar PIN del cliente** cargar el PIN del dueño.
8. **Implementar → Nueva implementación → Aplicación web.** Ejecutar como: *Yo*. Acceso: *Solo yo*.
9. Abrir el link de la web app en la tablet (logueada con la cuenta del local) y agregarla a la pantalla de inicio.
10. Entrar a Administración con el PIN y cargar los puestos, servicios y precios reales.

Para actualizar el código después: `clasp push` y en el editor *Implementar → Gestionar implementaciones → Editar → Nueva versión*.

## Tecnologías

Google Apps Script (V8), Google Sheets, HTML, CSS y JavaScript sin frameworks, clasp, Node.js para pruebas.

## Autor

Tahiel Cassata
