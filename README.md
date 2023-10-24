# Excel

orden de operaciones en excel

0. Parentesis 


1. Exponentes


2. multiplicaciones y diviciones


3. sumas y restas



## Formulas basicas


```
=REDONDEAR.MAS(PROMEDIO(B5:G5),0)
```

## IF 

Para utilizar un if se utiliza un si es como el equivalente.

```
=PROMEDIO(SI(H5:H29<6,H5:H29,""))
```

## Formulas para mayusculas y minusculas

```
=MAYUSC(EXTRAE(B2,1,1))&MINUSC(DERECHA(B2,LARGO(B2)-1))
=EXTRAE(B2,9,6)
=LARGO(B2)
```


## TRIM de Excel

Para quitar todos los espacios

```
=espacios(rango)
```

## Escribir literales

por ejemplo si quieres escribir =1+2 entonces pues pon '=1+2

```
si pones el '1 lo trataria como texto y no numero
```


## Concatena
 
El  & se utiliza para concatenar sin usar formulas


## Referencias estructuradas tablas

Para hace esto:

```
=suma(ventas[columna])
```

Ahora suma todo

## Shotcuts

Para agregar filas

```
CTRL + +  # Permite instertar filas de manera rapida
CTRL + t # Para crear una tabla
```




