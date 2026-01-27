#!/bin/bash

# -----------------------------------------------------------------------------
# Script para subir tu proyecto a GitHub
#
# INSTRUCCIONES:
# 1. Abre este archivo y reemplaza `<tu-usuario>` y `<tu-repositorio>`
#    con tu nombre de usuario de GitHub y el nombre de tu repositorio.
# 2. Guarda los cambios.
# 3. Abre tu terminal en la raíz del proyecto.
# 4. Da permisos de ejecución al script con el comando: chmod +x how-to-deploy.sh
# 5. Ejecuta el script con: ./how-to-deploy.sh
#
# NOTA: Este script asume que ya has creado un repositorio vacío en GitHub.
# -----------------------------------------------------------------------------

# Reemplaza con tus datos
GITHUB_USER="<tu-usuario>"
GITHUB_REPO="<tu-repositorio>"

# --- NO MODIFICAR DEBAJO DE ESTA LÍNEA ---

echo "Paso 1: Inicializando repositorio de Git..."
git init

echo "Paso 2: Añadiendo todos los archivos..."
git add .

echo "Paso 3: Creando el primer commit..."
git commit -m "Initial commit"

echo "Paso 4: Renombrando la rama principal a 'main'..."
git branch -M main

echo "Paso 5: Conectando con el repositorio remoto en GitHub..."
git remote add origin "https://github.com/${GITHUB_USER}/${GITHUB_REPO}.git"

echo "Paso 6: Subiendo los cambios a la rama 'main'..."
git push -u origin main

echo ""
echo "¡Proceso completado!"
echo "Tu código ha sido subido a https://github.com/${GITHUB_USER}/${GITHUB_REPO}"
echo "Para futuros cambios, solo necesitas usar 'git add .', 'git commit -m \"mensaje\"' y 'git push'."
