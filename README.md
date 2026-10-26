# NovaSheet

> Una hoja de cálculo web con actitud de producto: rápida, limpia y pensada para trabajar, no para pelear con menús.

![React](https://img.shields.io/badge/React-19-61DAFB?logo=react&logoColor=black)
![Vite](https://img.shields.io/badge/Vite-7-646CFF?logo=vite&logoColor=white)
![Tailwind](https://img.shields.io/badge/Tailwind-3-06B6D4?logo=tailwindcss&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-success)

---

## Por qué existe NovaSheet

NovaSheet nace para un caso muy concreto: abrir, editar y mover datos como en una spreadsheet sin cargar con toda la fricción de una suite gigantesca.

**No compite con Excel.**  
Compite con la pérdida de tiempo cuando solo quieres resolver algo en minutos.

---

## Qué puedes hacer hoy

- Editar celdas al vuelo.
- Agregar y eliminar filas/columnas.
- Pegar bloques de datos con expansión automática de la grilla.
- Mantener sesión gracias a persistencia en `localStorage`.
- Trabajar con rutas protegidas por login para flujo interno.

---

## Stack técnico

- **UI:** React 19
- **Build Tool:** Vite 7
- **Routing:** React Router
- **Estilos:** Tailwind CSS
- **Calidad:** ESLint (flat config)

---

## Arranque rápido

### Requisitos

- Node.js 20+
- pnpm 11+ (recomendado para respetar el lockfile)

### Desarrollo local

```bash
pnpm install
pnpm run dev
```

App local: `http://localhost:5173`

---

## Scripts del proyecto

```bash
pnpm run dev      # servidor de desarrollo
pnpm run build    # build de producción (carpeta dist/)
pnpm run preview  # sirve el build localmente
pnpm run lint     # valida estilo y reglas
```

---

## Estructura base

```text
src/
  componente/          # UI principal (header, admin, celdas, modales, etc.)
  componente/Auxiliares/
  view/                # vistas de login y spreadsheet
  lib/                 # utilidades compartidas
```

---

## Deploy en Netlify (importante)

Si en Netlify usas `pnpm run dev` como build command, el deploy **se queda colgado** porque `dev` levanta un servidor y no termina.

Este repositorio ya queda preparado para Netlify con:

- **Build command:** `pnpm run build`
- **Publish directory:** `dist`

Además, se configura redirección SPA para que rutas de React Router no den 404 al refrescar.

---

## Roadmap corto

- Atajos de teclado más completos
- Fórmulas básicas
- Export/import más robusto
- Mejoras visuales por tipo de celda

---

## Contribuciones

1. Haz fork
2. Crea rama: `feature/mi-mejora`
3. Ejecuta `pnpm run lint`
4. Abre PR explicando problema, solución y forma de prueba

---

## Licencia

MIT.

Si te fue útil, deja una estrella. Ayuda bastante a mantener el proyecto vivo. ⭐
