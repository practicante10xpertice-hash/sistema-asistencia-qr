# 📋 Sistema de Asistencia QR

Sistema web de control de asistencia de personal basado en **Google Apps Script** + **Google Sheets**, con interfaz moderna y registro mediante código QR.

---

## 🚀 Características

- **Login seguro** con contraseña hasheada en SHA-256
- **Dashboard** con estadísticas en tiempo real (asistencias, tardanzas, ausentes)
- **Gestión de personal** — crear, editar y eliminar empleados
- **Registro QR** — los empleados escanean un código con su celular para marcar asistencia
- **Registro manual** como alternativa al QR
- **Historial** con filtros por fecha, empleado y área
- **Exportar PDF** del reporte de asistencias
- **Detección de tardanzas** con tolerancia configurable (por defecto 15 min)
- **Doble turno** (mañana y tarde) por empleado
- **Modo demo** sin necesidad de conectar Apps Script

---

## 🗂️ Estructura del proyecto

```
sistema-asistencia-qr/
├── README.md
├── apps-script/
│   └── Code.gs          # Backend (Google Apps Script)
└── frontend/
    └── index.html       # Interfaz web completa
```

---

## ⚙️ Instalación

### 1. Crear la hoja de cálculo

1. Ve a [Google Sheets](https://sheets.google.com) y crea una nueva hoja
2. Nómbrala **SISTEMA QR**

### 2. Configurar Apps Script

1. Dentro de la hoja, ve a **Extensiones → Apps Script**
2. Borra el contenido del editor
3. Copia y pega el contenido de `apps-script/Code.gs`
4. Guarda el proyecto (Ctrl+S)

### 3. Subir el frontend

1. En Apps Script, crea un nuevo archivo HTML: **Archivo → Nuevo → Archivo HTML**
2. Nómbralo `index` (sin extensión)
3. Copia y pega el contenido de `frontend/index.html`
4. Guarda

### 4. Inicializar las hojas

1. En el editor de Apps Script, selecciona la función `initSheets` en el menú desplegable
2. Haz clic en ▶️ **Ejecutar**
3. Acepta los permisos que solicite Google

### 5. Publicar la aplicación

1. Haz clic en **Implementar → Nueva implementación**
2. Tipo: **Aplicación web**
3. Ejecutar como: **Yo**
4. Quién tiene acceso: **Cualquier persona**
5. Haz clic en **Implementar**
6. Copia la URL generada — esa es la URL de tu sistema

---

## 👤 Credenciales por defecto

| Usuario | Correo | Contraseña |
|---|---|---|
| Administrador | admin@empresa.com | admin123 |
| Usuario Demo | usuario@empresa.com | usuario123 |

> ⚠️ Cambia las contraseñas después de la primera instalación.

---

## 🗄️ Base de datos (Google Sheets)

El sistema crea automáticamente 4 hojas al ejecutar `initSheets()`:

| Hoja | Descripción |
|---|---|
| `Usuarios` | Credenciales de acceso al sistema |
| `Personal` | Empleados con horarios y días laborales |
| `Asistencias` | Registro de todas las marcaciones |
| `Config` | Token QR, nombre de empresa, tolerancia |

---

## 📱 Flujo de uso con QR

```
Administrador genera QR  →  Empleado escanea con celular
→  Sistema valida token  →  Registra entrada (Puntual / Tarde)
→  Aparece en dashboard e historial
```

---

## 🛠️ Tecnologías

- **Google Apps Script** — backend y base de datos
- **Google Sheets** — almacenamiento
- **HTML / CSS / JavaScript** — frontend (sin frameworks)
- **QRCode.js** — generación de códigos QR
- **Google Apps Script HtmlService** — hosting de la app

---

## 📄 Licencia

MIT — libre para uso personal y comercial.
