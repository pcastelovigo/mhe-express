# mhe-express
MHE-Express es un programa que automatiza la exportación de puntos de vista de Autodesk Navisworks, la conversión de los mismos a bloques de Autodesk AutoCAD y la generación posterior de un informe en Microsoft Word con imágenes, programado específicamente para agilizar el diseño de sistemas de CCTV.
MHE-Express es un programa (con un método) para automatizar
la exportación de puntos de vista de Navisworks a AutoCAD y la
generación de informes en Word.

Primer uso

Antes de utilizar la hoja de cálculo por primera vez, hay que
activar las referencias para que pueda leer datos de las otras
aplicaciones, se hace en Programador > VisualBasic >
Herramientas > Referencias
Las referencias a activar son AutoCAD 2019 Type ,Navisworks
Automation API 18, Navisworks Integrated API 18 y Microsoft
Word 16. No es necesario repetir este paso.
Modo de empleo
Abrir AutoCAD con el plano con los bloques que serán insertados
después.Pulsar el botón "BUSCAR BLOQUES A INSERTAR DESDE
AUTOCAD". El nombre de los bloques del plano saldrá en la
columna E.
Copiar los bloques en AutoCAD y pegarlos en el plano de destino
aparte.
Pulsar el botón "OBTENER VISTAS DE NAVISWORKS". Al pulsar el
botón, Navisworks se abre y da a escoger con que archivo. Una
vez escogido, Excel usará Navisworks para pasar
automáticamente por todas las vistas guardadas y copiará todos
los datos en las columnas de G a S.Aprovechar que Navisworks está abierto para obtener las
imágenes para el informe. Para ello, pulsar con el botón
derecho sobre las vistas y pulsar "EXPORTAR INFORME DE
PUNTOS DE VISTA". Una imagen de cada punto de vista se
generará en la carpeta donde esté guardado el archivo de
Navisworks.
Copiar el nombre del bloque deseado para cada fila desde la
columna E a la columna F, y pulsar "INSERTAR BLOQUES EN
AUTOCAD". Los bloques se insertarán en AutoCAD en las
posiciones obtenidas desde Navisworks.
Pulsar "GENERAR INFORME EN WORD" y escoger la carpeta
donde se encuentran las imagenes, si has seguido estas
instrucciones, es la misma donde se encuentra el archivo de
Navisworks de origen. El informe en Word se generará en 1-5
minutos.Opciones adicionales
• Escala
tamaño insercion: Cambia la escala del bloque
respecto al plano. Utilizar cuando el bloque sale demasiado
pequeño o demasiado grande. Valor por defecto: 1
• Escala plano: Cambia la escala del plano si el archivo
Navisworks original está en unidades diferentes a
milimetros. P. Ej.: 308.4 si el plano está en pies. Valor por
defecto: 1
• Altura planta: Define la cota de las alturas de cada planta,
para que el programa pueda colocarlas en la planta correcta
y definir la altura de montaje
• Desplazamiento plantas: Define el desplazamiento
vertical en AutoCAD entre diferentes plantas en la plantilla
utilizada. Valor por defecto: 508000
• Cambiar origen: Modifica el punto de origen respecto al
diseño de Navisworks en X,Y unidades. Valor por defecto: 0
Preguntas frecuentes
Me da error 5253 al generar el informe en Word
El número de archivos de imagen no concuerda con el número de
filas. Probablemente se han añadido puntos de vista después de
generar las imágenes o se ha escogido un directorio incorrecto.
Me da error "El portapapeles no está vacio" al generar el informe en Word
El programa utiliza el portapapeles para generar el informe, vacía
el portapapeles y cierra cualquier programa que pueda hacer uso
programático del mismo. Si el error continua, reinicia el
programa.
El FOV que obtiene el programa es incorrecto
El FOV cambia cada vez que se cambia el tamaño de la pantalla
de Navisworks, es una limitación de Navisworks al que Autodeskno da solución. Para solventarlo, asegurate de que obtienes los
puntos de vista usando el mismo tamaño de pantalla con el que
se definieron los puntos de vista, aprovechando el momento en el
que Excel te da a escoger el archivo a abrir.
Los bloques se insertan correctamente, pero fuera de lugar
Comprueba que el origen de coordenadas esté en el origen de
coordenadas universal. Seleccionar el origen de coordenadas,
pulsar botor derecho y pinchar en "Universal"
