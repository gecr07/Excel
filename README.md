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

## Absoluto Relativos 

NOTA si no se usan estos conceptos Excel no sabe de donde sacar datos y va a empezar a decir que no encontro los datos y poenr NA###

![image](https://github.com/gecr07/Excel/assets/63270579/07402f20-89ce-40c9-b180-5eeaaf038dd3)

NO puedes no entender esto es un powerup.


## Explicacion Absoluto - Absoluto ( nunca cambia)

![image](https://github.com/gecr07/Excel/assets/63270579/14ea94ff-bc79-4c8d-b859-64f386b0562e)

Si nos damos cuenta la celda AG3 esta asi


```
AG3
la celda esta:
$AG3
Osea absoluto absoluto nunca cambiara.
```


## Explicacion grafica Absoluto relativo

Tenemos esta tabla (nombre tabla)

![image](https://github.com/gecr07/Excel/assets/63270579/30e6eee0-5d50-48d4-81e2-6cd04fe47acb)

Despues vamos a hacer una formula que busque dentro de esa tabla. ( absoluto osea poniendole el$ - relativo ¿no poniendo nada? ) que siempre se refiera a la primea celda de cada fila

![image](https://github.com/gecr07/Excel/assets/63270579/1d4bec40-bc15-42d7-b39a-5754703f05e7)

Quedaria asi 

![image](https://github.com/gecr07/Excel/assets/63270579/ace19012-84c1-4c08-89e2-86847abda611)

En resumen: ( Absoluto - Relativo ) Para que saque la informacion de Filas


## Explicacion Relativo absoluto


![image](https://github.com/gecr07/Excel/assets/63270579/db2ed5df-bc15-47b3-ba85-cb63c89ef0bf)


(Relativo- Absoluto) Para que saque la informacion de columnas  Queremos asegurarnos que la informacion siempre este en la primera fila de cada columna. Y ahora si da los valores correctos.

## Objetos, metodos y propiedades

En excel todo son metodos y en VBA los iconos son verdes los iconos grises son propiedades.

![image](https://github.com/gecr07/Excel/assets/63270579/1702db51-353c-4b8d-aad4-ad9ac949fccf)


## Jerarquias

Excel -> Hoja -> Celda si no se especifica libro y hoja excel por defecto trabaja con el libro y hoja actual.

```
# Para seleccionar unas celdas
Range("A5").Select

Workbooks("Libro1.xlsx").Sheets("Hoja2").Range("A5")
```

## Variables 

Ya no uses Dim solo pues declaralas asi

```
Mivaribale=Range("A5").Value

```


## Copiar y Pegar


Se copia con formato a diferencia del Value copia con formato. no copia al porta papeles.

## Argumentos de los metodos

Se usa el := se pasan asi y , para un parametro nuevo.

```
Workbooks.Open Filename:="C:\Users\Fede\Escritorio\Libro1.xlsx", ReadOnly:=True
```

## Paso por valor 

```
Sub uno ()

Dim var1 as string
var1 = "test"
dos(var1)
End Sub


Sub dos(var1)

var1 vale test1

End Sub
```









