# FGWiki (v1.4 Release 95)
Esta es una Wiki local sumamente sencilla y liviana para organizar una pequeña base de conocimientos. Se puede combinar con el Sistema de Pedidos Lite para abrirla directamente desde el menú de la aplicación. Funciona con las versiones de 32 y 64 bits de Windows XP, Vista, 7, 8, 8.1, 10 y 11 así como sus variantes Windows Server.

![FGWiki](/Res/FGWiki.png)

## Modo de uso
Descargar el contenido de la carpeta /Bin/ y ejecutar el programa FGWiki.exe.

## Funcionamiento y primeros pasos
La FGWiki lista archivos de texto de una carpeta específica y permite crear nuevos archivos de texto o editar los ya existentes. Además, permite agregar *Tags* a los distintos archivos y buscar según varios criterios.

### Crear artículos
Para crear un artículo basta con hacer clic en Archivo > Nuevo Archivo, o presionar Ctrl + N. Se ingresa el nombre de archivo con su extensión (por lo general .txt) y automáticamente se activará el modo edición. Una vez finalizado, hacer clic en el botón Listo para guardar los cambios o descartarlos.

### Editar artículos
Para editar un artículo basta con hacer clic en Archivo > Editar archivo, hacer doble clic en el archivo o bien presionar Ctrl + E. Se activará el modo edición para modificar el archivo. Una vez finalizado, hacer clic en el botón Listo para guardar los cambios o descartarlos.

### Editar etiquetas
Para editar las etiquetas de un archivo, hacer clic en el icono que figura en la parte inferior derecha de la lista de archivos. Las etiquetas se deben separar por espacios, por ejemplo:
```
Windows Windows10 MiEtiqueta LimpiezaDeArchivos
```

### Asignar etiquetas masivamente
Haciendo clic en Opciones > Funciones experimentales > Añadir entrada Tags para todos los documentos, automatiza parte del proceso de creación de etiquetas. Esta función agregará la clave Tags= al archivo de etiquetas para todos los archivos existentes que no tengan etiquetas, mantendrá las existentes y quitará las de aquellos archivos que fueron borrados. Luego sugerirá abrir la aplicación de bloc de notas predeterminada para editar el archivo más rápidamente.

### Buscar artículos
Para filtrar entre las distintas entradas se debe utilizar el buscador de la esquina superior izquierda del programa. Se puede buscar por Contenido, por Etiquetas, por Nombre de archivo o por todos estos criterios a la vez, seleccionando la segunda lupa que aparece a la derecha de la búsqueda. Para quitar el filtro, se puede hacer clic en la lupa con un signo menos (-) que figura debajo de la lista de archivos, o borrar el contenido del campo de búsqueda y presionar Enter.

## Personalización
Se puede modificar la letra y el tamaño haciendo clic en Opciones > Cambiar tamaño de fuente. Allí se debe especificar el nombre de la fuente y seguido de un espacio y una coma, el tamaño de la fuente, por ejemplo:
```
Consolas, 12
```

## Limitaciones
El programa lee correctamente archivos de texto en formato ANSI. Por eso se recomienda utilizar la función "Nuevo archivo" (Ctrl + N) integrada en la Wiki. Si se opta por utilizar un editor de texto externo, se recomienda guardar el archivo en codificación ANSI.
