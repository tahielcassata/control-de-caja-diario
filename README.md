# Control de caja diario

**Sistema de cobros para negocios de servicios:** barberías, peluquerías, uñas, estética, spa y otros rubros. Web app para registrar los cobros desde una tablet en la caja y cerrar el día sin cuentas a mano. Reemplaza la planilla de papel donde se anotan los pagos en efectivo y por transferencia.

Funciona con **Google Apps Script + Google Sheets**: no necesita servidor ni base de datos, y no tiene costos de plataforma. Los datos quedan en una planilla de Google de la cuenta del local.

> **Demo:** abrí `docs/index.html` en el navegador (o la versión publicada en GitHub Pages). Funciona en modo demostración con datos de ejemplo. PIN de administración: `482916`. Tiene un botón para simular un corte de internet. Arriba se elige el rubro del ejemplo (también con `?rubro=Uñas`), y con `?vacio` se ve como recién instalado.

## Funciones

**Pantalla de cobro**
- Botones grandes por **puesto de trabajo** (Puesto 1, 2, 3…) y por **servicio**, con el precio cargado automáticamente.
- Medio de pago: efectivo o transferencia.
- **Propina opcional**, con su propio medio de pago (puede ser distinto al del corte).
- Totales del día siempre a la vista: efectivo, transferencias, total y propinas a entregar.
- **Deshacer último** para corregir un error de carga al instante.
- Después de guardar se desmarca todo, para que el próximo cobro arranque de cero.

**Cierre del día**
- Resumen general y por puesto: atenciones, plata cobrada y propinas a entregar.
- Botón **Cerrar día**: guarda el resumen y bloquea nuevas cargas. Al otro día la app arranca en cero.
- **Cierre automático** a la hora configurada, por si nadie lo cierra. Además, al abrir la app se cierra el día anterior si quedó abierto.
- **Reabrir día** con PIN, para cargar un cobro que quedó afuera. Al volver a cerrar se recalcula.
- Navegación por días anteriores.

**Administración (con PIN)**
- **Polirubro:** nombre del negocio propio y servicios del rubro que sea. Arranca sin servicios y trae **plantillas** para cargar de una vez los típicos de Barbería, Peluquería, Uñas, Estética o Masajes y spa (con precios de ejemplo para ajustar).
- Agregar puestos de trabajo y activar o desactivar los existentes (el historial no se pierde).
- Agregar servicios, cambiar nombres y precios, o desactivarlos. Los cambios de precio no tocan los cobros ni los cierres anteriores.
- Configurar la hora del cierre automático.
- El **PIN lo asigna el instalador** desde la planilla (menú *Control de caja → Asignar PIN del cliente*), y desde ahí también se resetea si el cliente se lo olvida. Tiene entre 6 y 8 números y no acepta PIN obvios (123456, 111111). Se guarda con hash SHA-256 y sal. Cada 5 intentos fallidos se bloquea, cada vez por más tiempo: 5 minutos, 1 hora y 24 horas. Administración se cierra sola a los 2 minutos sin uso.

**Funciona sin internet**
- Si se corta la conexión con la app abierta, se sigue cobrando normal: cada cobro se guarda en la tablet y se sube solo cuando vuelve internet.
- Un aviso muestra cuántos cobros quedan sin subir, y los totales del día ya los incluyen.
- Cada cobro lleva un código único y la hora real: si se reintenta no se duplica, y queda registrado a la hora en que se cobró.
- Si el cierre automático pasó durante el corte, al subir esos cobros el cierre de ese día se recalcula solo.
- Para abrir la app desde cero sí hace falta internet (la sirve Google).

**Diseño adaptativo**: el mismo código se ve bien en tablet, celular y PC.

**Licencias (control de pago del servicio)**
- Un **panel propio del desarrollador** (carpeta `licencias/`, en su cuenta de Google) con una fila por cliente: código, clave, fecha de pago y días de gracia.
- Cada sistema consulta su estado como mucho una vez por hora: **activo**, **aviso** (pago pendiente, muestra la fecha de suspensión) o **suspendido**.
- Suspendido solo bloquea la carga de cobros nuevos. Los datos del cliente siguen disponibles en «Cierre del día» y en su planilla, y los cobros que quedaron pendientes sin internet se guardan igual.
- Si el panel no responde, el sistema usa el último estado conocido y nunca se bloquea por un error. Sin configurar, el sistema anda normal.

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
- **Sin conexión:** cola de cobros en `localStorage`, subida en orden e idempotente (id único por cobro), con reintentos automáticos.
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
licencias/
  Code.gs           panel de licencias del desarrollador (proyecto aparte)
  appsscript.json
docs/
  index.html        copia para la demo en GitHub Pages
```

## Pruebas

El backend se prueba en Node con un simulador de las APIs de Google:

```bash
npm test
```

Cubre: licencias (estados por fecha, fallos del panel, bloqueo), carga y validación de cobros, cobros sin internet (duplicados, hora real, recálculo de cierres), propinas en distinto medio, deshacer, cierre manual y automático, reapertura con PIN, cambio de día, administración, bloqueo de PIN y rendimiento con un año de historial.

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
   clasp create --type sheets --title "Control de caja diario - Nombre del local" --rootDir src
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
6. Abrir el editor (`clasp open`), elegir la función `setup` y ejecutarla. Autorizar los permisos. Crea las hojas, 10 puestos y el cierre automático. Arranca sin servicios.
7. Volver a abrir la planilla y en el menú **Control de caja → Asignar PIN del cliente** cargar el PIN del dueño.
8. **Implementar → Nueva implementación → Aplicación web.** Ejecutar como: *Yo*. Acceso: *Solo yo*.
9. Abrir el link de la web app en la tablet (logueada con la cuenta del local) y agregarla a la pantalla de inicio.
10. Entrar a Administración con el PIN, cargar el nombre del negocio, ajustar los puestos y cargar los servicios (con la plantilla del rubro o uno por uno) con sus precios reales.

### Panel de licencias (una sola vez, en tu cuenta)

1. Crear el proyecto desde la carpeta `licencias`:
   ```bash
   cd licencias
   clasp create --type sheets --title "Licencias - Control de caja diario"
   git checkout -- appsscript.json
   clasp push
   ```
2. En el editor ejecutar `setup` y, en la hoja **Config**, escribir tu contacto (por ejemplo, tu WhatsApp).
3. **Implementar → Nueva implementación → Aplicación web.** Ejecutar como: *Yo*. Acceso: *Cualquier usuario*. Así los sistemas de los clientes pueden consultarla; sin el código y la clave de cada cliente solo responde "desconocido".
4. Por cada cliente: menú **Licencias → Nuevo cliente** y copiar la **línea de instalación**.
5. En la planilla del cliente: menú **Control de caja → Configurar licencia** y pegar esa línea.
6. Cuando el cliente paga: seleccionar su fila y usar **Licencias → Registrar pago (+1 mes)**. Para suspender o reactivar a mano, cambiar la columna **Modo**.

El acuerdo con el cliente (cuándo se suspende y qué incluye el servicio) tiene que estar por escrito antes de instalar.

Para actualizar el código después: `clasp push` y en el editor *Implementar → Gestionar implementaciones → Editar → Nueva versión*.

## Tecnologías

Google Apps Script (V8), Google Sheets, HTML, CSS y JavaScript sin frameworks, clasp, Node.js para pruebas.

## Autor

Tahiel Cassata

## Licencia

© 2026 Tahiel Cassata. **Todos los derechos reservados.** El código está publicado para que se pueda ver como parte de mi portfolio, pero no está permitido usarlo, copiarlo, instalarlo para terceros ni venderlo sin mi autorización por escrito. Ver [LICENSE](LICENSE).
