# Cómo Subir tu Proyecto a GitHub

Esta guía te muestra cómo subir tu código al repositorio `vinochelo/reportes-respaldo-2`.

## Opción 1: Usando el Script (Recomendado)

He preparado y configurado un script (`how-to-deploy.sh`) que automatiza todos los pasos.

**Ejecuta el Script desde tu terminal:**

```bash
# Primero, da permisos de ejecución al archivo
chmod +x how-to-deploy.sh

# Luego, ejecútalo
./how-to-deploy.sh
```

¡Y eso es todo! El script se encargará del resto.

## Opción 2: Comandos Manuales

Si prefieres ejecutar los comandos uno por uno, aquí los tienes listos para copiar y pegar en tu consola.

```bash
# 1. Inicializa un repositorio de Git (si no lo has hecho)
git init

# 2. Añade todos los archivos para ser rastreados
git add .

# 3. Guarda una "instantánea" de tus cambios
git commit -m "Initial commit"

# 4. (Opcional) Renombra la rama a 'main', que es la convención actual
git branch -M main

# 5. Conecta tu repositorio local con el que creaste en GitHub
git remote add origin https://github.com/vinochelo/reportes-respaldo-2.git

# 6. Sube tu código a GitHub
git push -u origin main
```

### Para futuras actualizaciones

Una vez que el proyecto está en GitHub, solo necesitarás estos tres comandos para subir nuevos cambios:

```bash
git add .
git commit -m "Describe tus nuevos cambios aquí"
git push
```
