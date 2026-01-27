# Cómo Subir tu Proyecto a GitHub

Esta guía te mostrará los pasos y comandos básicos para inicializar un repositorio de Git en tu proyecto, conectarlo a un repositorio remoto en GitHub y subir tus archivos.

## Pasos

### 1. Inicializar un Repositorio de Git

Si aún no lo has hecho, el primer paso es inicializar un repositorio de Git en la carpeta de tu proyecto. Abre tu terminal en la raíz del proyecto y ejecuta:

```bash
git init
```

Este comando crea un subdirectorio oculto `.git` que contiene todos los archivos necesarios para el repositorio.

### 2. Añadir los Archivos al Área de Preparación (Staging)

A continuación, añade todos los archivos de tu proyecto al área de preparación para que Git pueda empezar a rastrearlos.

```bash
git add .
```

El `.` le indica a Git que quieres añadir todos los archivos y directorios del proyecto.

### 3. Realizar tu Primer Commit

Un "commit" es como una instantánea de tus archivos en un momento dado. Guarda los cambios que has preparado en el historial del repositorio.

```bash
git commit -m "Initial commit"
```

El mensaje (`-m`) es una breve descripción de los cambios que estás guardando. "Initial commit" (o "Commit inicial") es un mensaje estándar para el primer commit.

### 4. Crear y Conectar tu Repositorio en GitHub

Ahora, ve a [GitHub](https://github.com) y crea un nuevo repositorio. No lo inicialices con un `README`, `.gitignore` o `licencia`, ya que tu proyecto ya los tiene.

Una vez creado, GitHub te proporcionará una URL para tu repositorio. Cópiala y úsala en el siguiente comando para conectar tu repositorio local con el remoto:

```bash
git remote add origin https://github.com/<tu-usuario>/<tu-repositorio>.git
```

**Importante:** Reemplaza `<tu-usuario>` y `<tu-repositorio>` con tu nombre de usuario de GitHub y el nombre de tu repositorio, respectivamente.

### 5. (Opcional) Renombrar la Rama Principal

Por convención, la rama principal de muchos proyectos se llama `main`. Si tu rama principal se llama `master`, puedes renombrarla con el siguiente comando:

```bash
git branch -M main
```

### 6. Subir tus Cambios a GitHub

Finalmente, sube tu código al repositorio remoto en GitHub. El comando `push` envía tus commits al servidor.

```bash
git push -u origin main
```

*   `-u`: Esta opción establece una relación de seguimiento entre tu rama local `main` y la rama `main` en el repositorio remoto (`origin`). Esto te permitirá usar `git pull` y `git push` en el futuro sin tener que especificar la rama cada vez.
*   `origin`: Es el nombre predeterminado para tu repositorio remoto.
*   `main`: Es el nombre de la rama que estás subiendo.

¡Y eso es todo! Tu código ahora está en GitHub. Para futuras actualizaciones, solo necesitarás usar `git add .`, `git commit -m "Tu mensaje"` y `git push`.
