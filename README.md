# Condiciones Técnicas
 
## 1. Objeto del Proyecto
Definir el alcance, requisitos y entregables para el desarrollo de la biblioteca **ConvertHtmlPdfLib** (HTML→PDF, firma PAdES, concatenación masiva).
 
- **Código del proyecto:** [ConvertHtmlPdfLib]  
- **Responsable:** anieto  
- **Versión:** 1.0  
- **Fecha de publicación:** [2024-09-20]  
- **Fecha de Acualización:** [2025-05-15] 
 
---
 
## 2. Alcance y Objetivos
1. **Alcance funcional**  
   - Conversión de uno o varios ficheros HTML a PDF.  
   - Concatenación de múltiples PDFs en un documento maestro.  
   - Inserción de cabeceras, pies de página y marcas de agua (“Borrador”, “Copia”).  
   - Firma digital PAdES mediante archivos `.p12`.  
 
2. **Alcance no funcional**  
   - Alto rendimiento: procesar lotes de documentos simultáneos.  
   - Escalabilidad: uso en entornos on-premise y contenedores Docker.  
   - Compatibilidad .NET 6+
 
3. **Objetivos específicos**  
   - Entregar una libreria clara, con ejemplos de uso en C#.  
   - Documentación completa (README, ejemplos, changelog).  
   - Cumplimiento de licenciamiento AGPL-3.0 (iText).  
 
---
 
## 3. Requisitos Funcionales
| ID   | Requisito                                               |
|------|---------------------------------------------------------|
| RF-01| Convertir un HTML local a PDF con un solo método.      |
| RF-02| Recibir codigo HTML y generar un solo PDF unificado. |
| RF-03| Añadir encabezado/pie de página configurable.           |
| RF-04| Firmar digitalmente un PDF existente con P12+password.  |
 
---
 
## 4. Requisitos Técnicos
- **Lenguaje/Plataforma:**  
  - .NET 6+ (C#)  
- **Dependencias:**  
  - iText 7 AGPL (núcleo + pdfHTML si aplica)  
  - BouncyCastle (para firma PAdES)  
- **Licencia:**  
  - AGPL-3.0. Incluir cabeceras de licencia en cada fichero y `LICENSE` completo.  
- **Entorno de ejecución:**  
  - Windows, Linux, Docker  
- **Control de versiones:**  
  - GitHub → rama principal `main`, política de Pull Requests revisados.  
 
---
 
## 5. Entregables
1. **Código fuente** en GitHub (estructura de carpetas clara).  
2. **Archivo LICENSE** con AGPL-3.0 y cabeceras en cada fichero.  
3. **README.md** con secciones: instalación, ejemplos de uso, badges, contribuciones.  
4. **CHANGELOG.md** con historial de versiones.  
6. **Paquete NuGet** publicado.  

---
 
## 6. Soporte y Mantenimiento
- Período de garantía: 3 meses tras entrega de v1.0.  
- Canal de reporte de incidencias: GitHub Issues.  
- Tiempo de respuesta SLAs:  
  - Crítico (producción caída): 24 h  
  - Alto (fallo grave): 48 h  
  - Medio/Bajo: 5 días hábiles  
 
---
 
## 7. Criterios de Aceptación
- Cumple todos los RF y RNF listados.  
- Documentación revisada y aprobada.  
- Pruebas unitarias verdes y cobertura ≥ 80%.  
- Ejemplos de uso ejecutables sin errores.  
- Licencia AGPL-3.0 correctamente aplicada.  
 
---
 
## 8. Anexos
- Enlace al repositorio: https://github.com/anietorpgye/ConvertHtmlPdfLib  
- Enlace a la AGPL-3.0: https://www.gnu.org/licenses/agpl-3.0.txt, https://itextpdf.com/how-buy/legal/agpl-gnu-affero-general-public-license
- Contacto: adrian.nieto.art@gmail.com  
 
---
 
*Este pliego puede adaptarse y expandirse según necesidades específicas de tu proyecto o requerimientos de tu organización.*
