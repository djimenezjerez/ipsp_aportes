# AUTOMATIZACIÓN DE LLENADO DE FORMULARIO DE APORTES

## ACERCA DEL SOFTWARE

Este software fue desarrollado para automatizar el llenado de los formularios de aportes de militantes del MAS-IPSP, para el mismo se precisa de las plantillas de [Recibo de Aporte en PDF](includes/PLANTILLA_FORMULARIO.pdf) y la plantilla de [Hoja de Excel de Aportantes](includes/PLANTILLA_APORTANTES.xlsx).

## DESARROLLO

Para continuar con el desarrollo es necesario contar con **Python 3.10.0** y el manejador de paquetes **pip** instalado.

Para instalar las dependencias en el entorno de producción se debe ejecutar:

```sh
pip install -r requirements/prod.txt
```

Para instalar las dependencias en el entorno de desarrollo se debe ejecutar:

```sh
pip install -r requirements/dev.txt
```

Para generar el archivo ejecutable se debe ejecutar el siguiente comando:

```sh
pyinstaller --onefile --windowed --icon=assets/mas_ipsp.ico main.py
```

## LICENCIA

* [LPG BOLIVIA versión 1](LICENSE.txt)
* [GPL versión 3.0](LICENSE)
