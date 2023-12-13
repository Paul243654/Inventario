# <h1 align="center"> Gestión del stock </h1>

## Introducción

Una empresa que sirve alimentos perecederos a sus clientes, recibe cada mañana en sus almacenes, la mercancía a granel enviada por sus proveedores. Una vez ingresada en el stock, se realiza la preparación de los pedidos y se envian a los clientes por la noche del mismo día.
Cada recepción de mercancía en el almacén, se aprovisiona para ir sirviendo a lo largo de toda la semana, dado que la mayor parte de la mercancía no puede estar mas de una semana almacenada, es muy importante tener controlado el stock, con la finalidad de enviar todo lo que pida el cliente y evitar perdidas, producto de los descuadres de stock.

<img align="left" width="400" height="400" src="https://github.com/Paul243654/Inventario/assets/112754073/c9e3c0c5-e10f-4881-b360-08387d635cf8">

<p align="center">
  <img width="400" height="400" src="https://github.com/Paul243654/Inventario/assets/112754073/feed961d-909c-4eed-9816-7a5b107ce92f">   
</p>



En la primera imagen tenemos el ejecutable que se encarga de realizar el cruce de archivos, ademas una vez acabado, tiene la opción de poder consultar el stock.
La segunda imagen nos muestra el formulario de entrada de los registros de errores, que se almacenan en una base de dato portable, para un posterior tratamiento de datos.


## Descripción

Extraemos ficheros "csv" de la base de datos de la empresa, la aplicación realiza el cruce de tablas de los ficheros, obteniendo un listado de stocks, ordenado por posiciones y código de artículo.

## Estado

![ready](https://github.com/Paul243654/Inventario/assets/112754073/5c545ff9-e225-48bb-9cbb-b6ad6300ea7f)


## Funcionalidades

El listado de stocks generado nos permitira:
- Realizar un inventario visual y poder detectar los artículos que presentan incidencia.
- Corregir la incidencia y actualizar el stock informáticamente.
- Generar un historial de incidencias, para sobre todo evaluar cada cierto tiempo los errores sistemáticos que se presentan.
  


![excel_fichero](https://github.com/Paul243654/Inventario/assets/112754073/055bf3cf-a77e-424e-b45a-d39e0c2daaf5)


Podemos observar en la imagen, el listado generado del stock de todos los artículos, con esta herramienta, el operario realizara un conteo visual y anotar la cantidad real así como tambien alguna observación si la hubiese.
Si hubiese descuadre de stock, se revisarían los artículos marcados y se procedería a solucionar la incidencia para finalmente registrala en la base de datos portable.


## Autores

Paul Nuñez

## Licencia
