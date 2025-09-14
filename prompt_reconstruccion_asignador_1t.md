# PROMPT: Reconstrucción Limpia y Optimizada del Asignador de Turnos 1T

## DESCRIPCIÓN GENERAL
Crear una clase `AsignadorTurnos` que asigne turnos "1T" (1 hora extra) o "7" (1 hora extra + 6 horas adicionales) sobre un archivo Excel procesado, siguiendo reglas específicas de asignación y equidad.

## FUNCIONALIDADES PRINCIPALES

### 1. GESTIÓN DE ARCHIVOS
- **Entrada**: Aceptar archivo Excel procesado (prioridad: archivo especificado → "horario_procesado_con_sabados_domingos.xlsx" → "horarioUnificado_procesado.xlsx")
- **Salida**: Generar "horarioUnificado_con_1t.xlsx"
- **Hoja de trabajo**: Usar hoja principal (excluir "Estadísticas" si existe)
- **Dependencias**: openpyxl, random, collections.defaultdict, typing

### 2. REGLAS DE ASIGNACIÓN POR PERSONAL DISPONIBLE
Buscar fila "TURNOS OPERATIVOS" en columna A y aplicar:
- **≤8 operativos**: No asignar turno
- **=9 operativos**: Asignar turno "7"
- **≥10 operativos**: Asignar turno "1T"

### 3. TRABAJADORES ELEGIBLES
Lista fija: `['GCE', 'YIS', 'MAQ', 'DJO', 'AFG', 'JLF', 'JMV']`

### 4. RESTRICCIONES DE ASIGNACIÓN

#### Restricciones Duras (NO asignar):
- **Día anterior**: Si tuvo BANTD, BLPTD, NLPRD, NANRD, 1T o 7
- **Día siguiente**: Si tiene BANTD, BLPTD, 1T o 7
- **Día actual**: Si ya existe 1T, 7, BLPTD o BANTD (no duplicar)

#### Restricciones Blandas (evitar si es posible):
- **Día anterior**: Si tuvo NANTD o NLPTD

#### Restricción Especial:
- **GCE + Torre**: Solo asignar 1T si conteo "Torre" ≤ 3

### 5. SISTEMA DE PRIORIDADES
Clasificar candidatos en 5 niveles:
1. **Nivel 1**: Prioridad (DESC/TROP/SIND ayer) + sin extra ayer + sin restricción blanda
2. **Nivel 2**: Sin extra ayer + sin restricción blanda
3. **Nivel 3**: Prioridad + restricción blanda
4. **Nivel 4**: Solo restricción blanda
5. **Nivel 5**: Resto de candidatos

### 6. SISTEMA DE EQUIDAD
Mantener contadores separados:
- **Grupo 1T**: Cuenta "1T" + "7" (toda persona con 1 hora extra)
- **Grupo 6RT**: Cuenta solo "7" (6 horas adicionales)

**Lógica de selección**:
- Para turno "1T": Balancear por grupo 1T
- Para turno "7": Balancear por grupo 1T; si empate, usar grupo 6RT como desempate
- Selección aleatoria entre candidatos con mínimo contador

### 7. INICIALIZACIÓN DE CONTADORES
Al iniciar, leer asignaciones existentes en la hoja y actualizar contadores:
- Recorrer todas las columnas (2 a max_col)
- Para cada trabajador elegible, contar 1T y 7 ya asignados

### 8. ALERTAS Y COMENTARIOS
Agregar comentarios en encabezados de días cuando:
- No se puede asignar por restricción dura del día anterior
- No se puede asignar por restricción dura del día siguiente
- Concatenar mensajes si ya existen comentarios

### 9. HOJA DE ESTADÍSTICAS
Crear/actualizar hoja "Estadísticas" con columnas:
- **SIGLA**: Nombre del trabajador
- **DESC**: Fórmula COUNTIF para DESC + TROP
- **1T**: Fórmula COUNTIF para 1T + 7
- **6RT**: Fórmula COUNTIF para solo 7
- **1D**: Fórmula COUNTIF para BANTD + BLPTD
- **3D**: Fórmula COUNTIF para 3 + 3D
- **6D**: Fórmula COUNTIF para NLPTD + NLPRD + NANTD + NANRD

**Formato**: Encabezados con fondo gris (E6E6E6) y negrita, ancho de columnas optimizado

## MÉTODOS REQUERIDOS

### Métodos de Inicialización:
- `__init__(archivo_procesado=None)`: Constructor principal
- `_resolver_archivo_entrada(preferido)`: Resolver archivo de entrada
- `_obtener_hoja_horario()`: Obtener hoja principal
- `_inicializar_contadores_desde_hoja()`: Inicializar contadores desde datos existentes

### Métodos de Búsqueda y Validación:
- `_obtener_fila_trabajador(trabajador)`: Encontrar fila del trabajador (2-25)
- `_obtener_conteo_operativos(col_dia)`: Buscar valor en fila "TURNOS OPERATIVOS"
- `_obtener_conteo_torre(col_dia)`: Buscar valor en fila "Torre"
- `_obtener_trabajadores_disponibles(col_dia)`: Lista de trabajadores con celda vacía

### Métodos de Validación de Restricciones:
- `_tiene_prioridad_dia_anterior(trabajador, col_dia)`: DESC/TROP/SIND ayer
- `_tuvo_extra_dia_anterior(trabajador, col_dia)`: 1T/7 ayer
- `_tuvo_restriccion_dura_ayer(trabajador, col_dia)`: BANTD/BLPTD/NLPRD/NANRD/1T/7 ayer
- `_tiene_restriccion_dura_manana(trabajador, col_dia)`: BANTD/BLPTD/1T/7 mañana
- `_tuvo_restriccion_blanda_ayer(trabajador, col_dia)`: NANTD/NLPTD ayer
- `_existe_turno_1t_o_7_en_dia(col_dia)`: Verificar duplicados en día actual

### Métodos de Lógica de Negocio:
- `_determinar_turno_por_personal(col_dia)`: Decidir 1T/7/None según operativos
- `_seleccionar_equitativo(candidatos, turno)`: Selección con equidad
- `_actualizar_contadores(trabajador, turno)`: Actualizar contadores 1T y 6RT

### Métodos de Asignación:
- `asignar_turno_en_dia(col_dia)`: Lógica principal de asignación
- `procesar_todos_los_dias()`: Procesar todas las columnas

### Métodos de Reporte:
- `_marcar_alerta_restriccion_dura(col_dia, mensaje)`: Agregar comentarios de alerta
- `_actualizar_hoja_estadisticas()`: Crear/actualizar hoja de estadísticas

## ESTRUCTURA DE DATOS

### Atributos de Clase:
- `TRABAJADORES_ELEGIBLES`: Lista constante de trabajadores
- `archivo_procesado`: Ruta del archivo de entrada
- `wb`: Workbook de openpyxl
- `ws`: Worksheet principal
- `contador_grupo_1t`: Dict[str, int] para contador 1T
- `contador_grupo_6rt`: Dict[str, int] para contador 6RT

### Constantes y Conjuntos:
- Turnos de prioridad: `{"DESC", "TROP", "SIND"}`
- Turnos extra: `{"1T", "7"}`
- Restricciones duras ayer: `{"BANTD", "BLPTD", "NLPRD", "NANRD", "1T", "7"}`
- Restricciones duras mañana: `{"BANTD", "BLPTD", "1T", "7"}`
- Restricciones blandas: `{"NANTD", "NLPTD"}`
- Turnos duplicados: `{"1T", "7", "BLPTD", "BANTD"}`

## OPTIMIZACIONES SUGERIDAS

### Rendimiento:
- Cachear búsquedas de filas de trabajadores
- Optimizar búsquedas de etiquetas (TURNOS OPERATIVOS, Torre)
- Reducir iteraciones innecesarias en validaciones

### Código Limpio:
- Separar lógica de negocio en métodos más pequeños
- Usar enums o constantes para tipos de turnos
- Implementar validación de entrada más robusta
- Agregar logging para debugging

### Mantenibilidad:
- Documentar métodos con docstrings completos
- Agregar tipos de datos más específicos
- Implementar manejo de errores más robusto
- Crear tests unitarios

### Funcionalidades Adicionales:
- Configuración externa de parámetros
- Reportes de asignación más detallados
- Validación de integridad de datos
- Backup automático antes de modificar

## EJECUCIÓN
```python
if __name__ == "__main__":
    asignador = AsignadorTurnos()
    asignador.procesar_todos_los_dias()
```

## NOTAS IMPORTANTES
- Usar `random.seed()` para reproducibilidad
- Manejar casos edge (archivos vacíos, hojas inexistentes)
- Validar que las fórmulas Excel sean correctas
- Asegurar compatibilidad con diferentes versiones de Excel
- Mantener la funcionalidad exacta del código original 