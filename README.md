# Procesador de Horarios de Controladores Aéreos

## 📋 Descripción

Este programa en Python procesa archivos Excel de horarios unificados para controladores aéreos (ATCOs), aplicando análisis automático, conteo de turnos operativos y generación de estadísticas.

## 🚀 Funcionalidades

### **Procesamiento Principal**
- **Carga automática** del archivo `horioUnificado.xlsx`
- **Conteo dinámico** de turnos operativos por columna (días)
- **Coloreado inteligente** de celdas según el tipo de turno
- **Identificación automática** de domingos y aplicación de colores especiales
- **Preservación** de la estructura original de la tabla

### **Sistema de Conteo**
- **Lógica refinada**: Cuenta celdas vacías + celdas con contenido NO incluido en la lista de turnos no operativos
- **Valores calculados**: Conteo preciso basado en la lista de 47 turnos no operativos
- **Fila de resumen**: Agregada al final con el conteo por día

### **Coloreado Inteligente**
- **Turnos no operativos**: Amarillo (solo los de la lista oficial)
- **Encabezados de domingos**: Rojo claro
- **Conteo por rangos**:
  - ≤8: Rojo intenso con fuente blanca
  - =9: Rojo medio
  - =10: Azul clarito
  - =11: Verde clarito
  - =12: Verde intenso
  - ≥13: Sin relleno

### **Hoja de Estadísticas**
- **Estructura simplificada**: Solo 2 columnas
  - Columna A: SIGLA
  - Columna B: DESC (conteo unificado)
- **Fórmula unificada**: `=COUNTIF(HorarioUnificado!B2:AC2,"DESC")+COUNTIF(HorarioUnificado!B2:AC2,"TROP")`
- **Actualización automática**: Se actualiza al modificar la hoja principal
- **Formato profesional**: Encabezados con fondo gris y fuente en negrita
- **Ancho optimizado**: Columnas ajustadas al mínimo necesario para visualizar todos los valores

## 📊 Estructura del Archivo de Entrada

### **Archivo Original: `horioUnificado.xlsx`**
- **Columna A**: "SIGLA ATCO" (código del controlador)
- **Columnas B en adelante**: Fechas en formato "DÍA-DD" (ej: MON-04, TUE-05)
- **Filas 2-25**: 22 controladores con sus turnos asignados
- **Propósito**: Horario unificado para gestión de programación laboral

## 🎯 Lista de Turnos No Operativos

El programa reconoce **47 tipos** de turnos no operativos:

### **Turnos Básicos**
- `DESC`, `TROP`

### **Turnos Completos**
- `VACA`, `COME`, `COMT`, `COMS`

### **Formación, Instrucción y Entrenamiento**
- `SIND`, `CMED`, `CERT`
- `CAPA`, `MCAE`, `TCAE`, `MCHC`, `TCHC`, `NCHC`, `ACHC`
- `MENT`, `TENT`, `NENT`, `AENT`
- `MINS`, `TINS`, `NINS`, `AINS`

### **Gestión, Oficinas y Grupos de Trabajo**
- `MCOR`, `TCOR`, `MSMS`, `TSMS`
- `MDBM`, `TDBM`, `MDOC`, `TDOC`
- `MPRO`, `TPRO`, `MATF`, `TATF`
- `MGST`, `TGST`, `MOFI`, `TOFI`

### **Operativos y Asignaciones Especiales**
- `CET`, `ATC`, `KATC`, `XATC`, `YATC`, `ZATC`, `X`

## 🛠️ Instalación

1. **Clonar o descargar** el proyecto
2. **Instalar dependencias**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Asegurar** que el archivo `horioUnificado.xlsx` esté en el directorio del proyecto

## 📝 Uso

### **Ejecución Básica**
```bash
python procesador_horarios.py
```

### **Archivos Generados**
- **`horarioUnificado_procesado.xlsx`**: Archivo principal procesado
- **Hojas incluidas**:
  - `HorarioUnificado`: Datos originales con procesamiento
  - `Estadísticas`: Conteo unificado de DESC + TROP

### **Resultados Esperados**
1. **Conteo automático** de turnos operativos por día
2. **Coloreado inteligente** según el tipo de turno
3. **Identificación visual** de domingos
4. **Estadísticas dinámicas** de turnos DESC y TROP
5. **Preservación** de todos los datos originales

## 📁 Estructura de Archivos

```
generadorDescFiles/
├── horioUnificado.xlsx              # Archivo de entrada
├── procesador_horarios.py           # Programa principal
├── requirements.txt                  # Dependencias
├── README.md                        # Documentación
└── horarioUnificado_procesado.xlsx  # Archivo de salida
```

## 🔧 Características Técnicas

### **Librerías Utilizadas**
- **openpyxl**: Manipulación de archivos Excel
- **PatternFill**: Aplicación de colores de fondo
- **Font**: Configuración de fuentes

### **Manejo de Errores**
- **FileNotFoundError**: Archivo de entrada no encontrado
- **PermissionError**: Archivo de salida en uso
- **Validación**: Verificación de estructura de datos

### **Optimizaciones**
- **Limpieza de formato**: Elimina colores existentes antes de aplicar nuevos
- **Fórmulas dinámicas**: Para estadísticas automáticas
- **Cálculos precisos**: Conteo basado en lista oficial de turnos

## 📈 Funcionalidades Avanzadas

### **Conteo Dinámico**
- **Lógica refinada**: Solo cuenta turnos operativos (vacíos + no listados)
- **Precisión**: Basado en la lista oficial de 47 turnos no operativos
- **Flexibilidad**: Se adapta a cambios en la estructura de datos

### **Coloreado Condicional**
- **Rangos específicos**: 6 niveles de color según el conteo
- **Identificación visual**: Domingos marcados automáticamente
- **Consistencia**: Aplicación uniforme en toda la tabla

### **Estadísticas Automáticas**
- **Fórmula unificada**: DESC + TROP en una sola celda
- **Actualización automática**: Respuesta a cambios en datos
- **Interfaz limpia**: Solo 2 columnas esenciales
- **Encabezado optimizado**: "SIGLA" para mayor claridad
- **Ancho mínimo**: Columnas ajustadas para visualización óptima

## 🔄 Proceso de Actualización

### **Para Modificar Conteos**
1. **Editar** el archivo `horioUnificado.xlsx`
2. **Ejecutar** `python procesador_horarios.py`
3. **Verificar** los resultados en `horarioUnificado_procesado.xlsx`

### **Para Modificar Colores**
- **Editar** las variables de color en `procesador_horarios.py`
- **Ejecutar** el programa para aplicar cambios

### **Para Modificar Turnos No Operativos**
- **Actualizar** la lista `turnos_no_operativos` en el código
- **Reejecutar** para aplicar la nueva lógica

## 📊 Ejemplos de Salida

### **Conteo por Día**
- **Columna B (LUN-04)**: 15 turnos operativos
- **Columna C (MAR-05)**: 12 turnos operativos
- **Columna D (MIÉ-06)**: 18 turnos operativos

### **Estadísticas de Trabajador**
- **PHD**: 3 turnos DESC + TROP
- **HLG**: 2 turnos DESC + TROP
- **MEI**: 4 turnos DESC + TROP
- **Estructura**: Columna SIGLA + conteo unificado DESC+TROP

## 🎨 Esquema de Colores

| Rango de Conteo | Color de Fondo | Color de Fuente |
|------------------|----------------|-----------------|
| ≤8 | Rojo intenso | Blanco |
| =9 | Rojo medio | Negro |
| =10 | Azul clarito | Negro |
| =11 | Verde clarito | Negro |
| =12 | Verde intenso | Negro |
| ≥13 | Sin relleno | Negro |

## 🔍 Notas Importantes

- **Preservación de datos**: No se modifica el contenido original
- **Formato limpio**: Se eliminan colores existentes antes de aplicar nuevos
- **Fórmulas dinámicas**: Solo en la hoja de estadísticas
- **Conteo calculado**: En la hoja principal para precisión
- **Actualización manual**: Reejecutar programa para cambios en conteos principales

## Módulos de asignación de turnos (1T, 6RT, 6TT)

- Ejecución 1T: `python asignador_turnos_1t.py`
- Ejecución 6RT: `python asignador_turnos_6rt.py`
- Ejecución 6TT: `python asignador_turnos_6tt.py`

### Reglas para 6TT
- Elegibles: `['YIS','MAQ','DJO','AFG','JLF','JMV']` (CDT excluido)
- Fallback si no hay celdas libres en elegibles: `['FCE','JBV','HZG']`
- Decisión por día: si Turnos Operativos >=16 no se asigna; si <=15 se asigna
- Máximo un `6TT` por día
- Prioridad: evitar `1T`/`1`/`7` al día siguiente

### Rebalanceo 6RT+6TT (±1)
- Tras asignar 6TT, se reequilibran los totales 6RT+6TT entre elegibles moviendo `6TT` dentro del mismo día
- Se preservan todas las restricciones diarias y se prioriza receptor sin `1T`/`1`/`7` al día siguiente

### Estadísticas (asignador_turnos_6tt.py)
- Columnas: `SIGLA`, `DESC`, `1T` (1T+7), `6RT` (6RT+7), `6T` (solo 6TT), `6RT+6TT`
- Archivo de salida: `horarioUnificado_con_6tt.xlsx`

---

**Versión**: 2.1  
**Última actualización**: Asignación 6TT con rebalanceo (±1) y columna 6RT+6TT  
**Compatibilidad**: Excel 2016+  
**Python**: 3.7+ #   F e a t u r e :   I m p r o v e d   S c h e d u l e   C o l o r s  
 