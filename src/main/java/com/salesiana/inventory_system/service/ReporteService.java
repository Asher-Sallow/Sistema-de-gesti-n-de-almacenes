package com.salesiana.inventory_system.service;

import com.salesiana.inventory_system.entity.Producto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ReporteService {
    
    @Autowired
    private ProductoService productoService;
    
    @Autowired
    private MovimientoService movimientoService;
    
    // ============================================
    // REPORTES DE INVENTARIO - EXCEL
    // ============================================
    
    /**
     * Genera reporte de inventario completo en formato Excel CON FORMATO PROFESIONAL
     */
    public byte[] generarReporteInventarioExcel() throws IOException {
        List<Producto> productos = productoService.obtenerTodosProductos();
        
        Workbook workbook = null;
        ByteArrayOutputStream out = null;
        
        try {
            workbook = new XSSFWorkbook();
            out = new ByteArrayOutputStream();
            
            Sheet sheet = workbook.createSheet("Inventario");
            
            // 🎨 ESTILOS PROFESIONALES
            CellStyle tituloStyle = crearEstiloTitulo(workbook);
            CellStyle fechaStyle = crearEstiloFecha(workbook);
            CellStyle headerStyle = crearEstiloEncabezado(workbook, IndexedColors.DARK_GREEN);
            CellStyle datosStyle = crearEstiloDatos(workbook);
            CellStyle monedaStyle = crearEstiloMoneda(workbook);
            CellStyle resumenStyle = crearEstiloResumen(workbook);
            CellStyle estadoAgotadoStyle = crearEstiloEstado(workbook, IndexedColors.RED);
            CellStyle estadoCriticoStyle = crearEstiloEstado(workbook, IndexedColors.ORANGE);
            CellStyle estadoNormalStyle = crearEstiloEstado(workbook, IndexedColors.LIGHT_GREEN);
            
            int rowNum = 0;
            
            // 📌 TÍTULO DEL REPORTE
            Row tituloRow = sheet.createRow(rowNum++);
            Cell tituloCell = tituloRow.createCell(0);
            tituloCell.setCellValue("REPORTE DE INVENTARIO - DROGUERÍA INTI");
            tituloCell.setCellStyle(tituloStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));
            tituloRow.setHeight((short) 600);
            
            // 📅 FECHA DE GENERACIÓN
            Row fechaRow = sheet.createRow(rowNum++);
            Cell fechaCell = fechaRow.createCell(0);
            fechaCell.setCellValue("Fecha de Generación: " + 
                LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
            fechaCell.setCellStyle(fechaStyle);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 8));
            
            // Línea en blanco
            rowNum++;
            
            // 🎯 ENCABEZADOS
            Row headerRow = sheet.createRow(rowNum++);
            String[] headers = {
                "Código", "Nombre", "Categoría", "Stock Actual", "Stock Mínimo", 
                "Precio Compra", "Precio Venta", "Valor Inventario", "Estado"
            };
            
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }
            headerRow.setHeight((short) 400);
            
            // 📊 DATOS DE PRODUCTOS
            double valorTotalInventario = 0;
            int totalAgotados = 0;
            int totalStockBajo = 0;
            
            for (Producto producto : productos) {
                Row row = sheet.createRow(rowNum++);
                
                // Código
                Cell cellCodigo = row.createCell(0);
                cellCodigo.setCellValue(producto.getCodigo());
                cellCodigo.setCellStyle(datosStyle);
                
                // Nombre
                Cell cellNombre = row.createCell(1);
                cellNombre.setCellValue(producto.getNombre());
                cellNombre.setCellStyle(datosStyle);
                
                // Categoría
                Cell cellCategoria = row.createCell(2);
                cellCategoria.setCellValue(
                    producto.getCategoria() != null ? producto.getCategoria().getNombre() : "Sin categoría"
                );
                cellCategoria.setCellStyle(datosStyle);
                
                // Stock Actual
                Cell cellStock = row.createCell(3);
                cellStock.setCellValue(producto.getStockActual());
                cellStock.setCellStyle(datosStyle);
                
                // Stock Mínimo
                Cell cellStockMin = row.createCell(4);
                cellStockMin.setCellValue(producto.getStockMinimo());
                cellStockMin.setCellStyle(datosStyle);
                
                // Precio Compra
                Cell cellPrecioCompra = row.createCell(5);
                if (producto.getPrecioCompra() != null) {
                    cellPrecioCompra.setCellValue(producto.getPrecioCompra().doubleValue());
                } else {
                    cellPrecioCompra.setCellValue(0);
                }
                cellPrecioCompra.setCellStyle(monedaStyle);
                
                // Precio Venta
                Cell cellPrecioVenta = row.createCell(6);
                if (producto.getPrecioVenta() != null) {
                    cellPrecioVenta.setCellValue(producto.getPrecioVenta().doubleValue());
                } else {
                    cellPrecioVenta.setCellValue(0);
                }
                cellPrecioVenta.setCellStyle(monedaStyle);
                
                // Valor Inventario
                double valorInventario = producto.getPrecioCompra() != null ? 
                    producto.getStockActual() * producto.getPrecioCompra().doubleValue() : 0;
                Cell cellValor = row.createCell(7);
                cellValor.setCellValue(valorInventario);
                cellValor.setCellStyle(monedaStyle);
                valorTotalInventario += valorInventario;
                
                // Estado con COLORES
                String estado = determinarEstadoStock(producto);
                Cell cellEstado = row.createCell(8);
                cellEstado.setCellValue(estado);
                
                if (estado.equals("AGOTADO")) {
                    cellEstado.setCellStyle(estadoAgotadoStyle);
                    totalAgotados++;
                } else if (estado.equals("CRÍTICO")) {
                    cellEstado.setCellStyle(estadoCriticoStyle);
                    totalStockBajo++;
                } else {
                    cellEstado.setCellStyle(estadoNormalStyle);
                }
            }
            
            // 📈 RESUMEN
            rowNum++; // Línea en blanco
            Row resumenTituloRow = sheet.createRow(rowNum++);
            Cell resumenTituloCell = resumenTituloRow.createCell(0);
            resumenTituloCell.setCellValue("RESUMEN DEL INVENTARIO");
            resumenTituloCell.setCellStyle(headerStyle);
            sheet.addMergedRegion(new CellRangeAddress(rowNum-1, rowNum-1, 0, 3));
            
            // Total productos
            Row totalProductosRow = sheet.createRow(rowNum++);
            totalProductosRow.createCell(0).setCellValue("Total de Productos:");
            totalProductosRow.getCell(0).setCellStyle(resumenStyle);
            totalProductosRow.createCell(1).setCellValue(productos.size());
            totalProductosRow.getCell(1).setCellStyle(datosStyle);
            
            // Stock bajo
            Row stockBajoRow = sheet.createRow(rowNum++);
            stockBajoRow.createCell(0).setCellValue("Productos con Stock Bajo:");
            stockBajoRow.getCell(0).setCellStyle(resumenStyle);
            stockBajoRow.createCell(1).setCellValue(totalStockBajo);
            stockBajoRow.getCell(1).setCellStyle(datosStyle);
            
            // Agotados
            Row agotadosRow = sheet.createRow(rowNum++);
            agotadosRow.createCell(0).setCellValue("Productos Agotados:");
            agotadosRow.getCell(0).setCellStyle(resumenStyle);
            agotadosRow.createCell(1).setCellValue(totalAgotados);
            agotadosRow.getCell(1).setCellStyle(datosStyle);
            
            // Valor total
            Row valorTotalRow = sheet.createRow(rowNum++);
            valorTotalRow.createCell(0).setCellValue("Valor Total del Inventario:");
            valorTotalRow.getCell(0).setCellStyle(resumenStyle);
            Cell valorTotalCell = valorTotalRow.createCell(1);
            valorTotalCell.setCellValue(valorTotalInventario);
            valorTotalCell.setCellStyle(monedaStyle);
            
            // Autoajustar columnas
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 500);
            }
            
            // CRÍTICO: Escribir al stream ANTES de cerrar
            workbook.write(out);
            return out.toByteArray();
            
        } finally {
            // Cerrar recursos en orden inverso
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    System.err.println("Error al cerrar workbook: " + e.getMessage());
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    System.err.println("Error al cerrar output stream: " + e.getMessage());
                }
            }
        }
    }
    
    /**
     * Genera reporte de stock bajo en formato Excel con formato profesional
     */
    public byte[] generarReporteStockBajoExcel() throws IOException {
        List<Producto> productos = productoService.obtenerProductosStockBajo();
        
        Workbook workbook = null;
        ByteArrayOutputStream out = null;
        
        try {
            workbook = new XSSFWorkbook();
            out = new ByteArrayOutputStream();
            
            Sheet sheet = workbook.createSheet("Stock Bajo");
            
            // Estilos
            CellStyle tituloStyle = crearEstiloTitulo(workbook);
            CellStyle fechaStyle = crearEstiloFecha(workbook);
            CellStyle headerStyle = crearEstiloEncabezado(workbook, IndexedColors.ORANGE);
            CellStyle datosStyle = crearEstiloDatos(workbook);
            CellStyle alertaStyle = crearEstiloEstado(workbook, IndexedColors.RED);
            
            int rowNum = 0;
            
            // Título
            Row tituloRow = sheet.createRow(rowNum++);
            Cell tituloCell = tituloRow.createCell(0);
            tituloCell.setCellValue("⚠️ REPORTE DE STOCK BAJO - DROGUERÍA INTI");
            tituloCell.setCellStyle(tituloStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));
            tituloRow.setHeight((short) 600);
            
            // Fecha
            Row fechaRow = sheet.createRow(rowNum++);
            Cell fechaCell = fechaRow.createCell(0);
            fechaCell.setCellValue("Fecha: " + 
                LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
            fechaCell.setCellStyle(fechaStyle);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 5));
            
            // Alerta
            Row alertaRow = sheet.createRow(rowNum++);
            Cell alertaCell = alertaRow.createCell(0);
            alertaCell.setCellValue("⚠️ ALERTA: Estos productos requieren reposición inmediata");
            alertaCell.setCellStyle(alertaStyle);
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 5));
            
            rowNum++; // Línea en blanco
            
            // Encabezados
            Row headerRow = sheet.createRow(rowNum++);
            String[] headers = {"Código", "Nombre", "Categoría", "Stock Actual", "Stock Mínimo", "Déficit"};
            
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }
            
            // Datos
            for (Producto producto : productos) {
                Row row = sheet.createRow(rowNum++);
                
                row.createCell(0).setCellValue(producto.getCodigo());
                row.createCell(1).setCellValue(producto.getNombre());
                row.createCell(2).setCellValue(
                    producto.getCategoria() != null ? producto.getCategoria().getNombre() : "Sin categoría"
                );
                row.createCell(3).setCellValue(producto.getStockActual());
                row.createCell(4).setCellValue(producto.getStockMinimo());
                row.createCell(5).setCellValue(producto.getStockMinimo() - producto.getStockActual());
                
                for (int i = 0; i < 6; i++) {
                    row.getCell(i).setCellStyle(datosStyle);
                }
            }
            
            // Autoajustar columnas
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 500);
            }
            
            workbook.write(out);
            return out.toByteArray();
            
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    System.err.println("Error al cerrar workbook: " + e.getMessage());
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    System.err.println("Error al cerrar output stream: " + e.getMessage());
                }
            }
        }
    }
    
    // ============================================
    // REPORTES DE INVENTARIO - PDF (HTML)
    // ============================================
    
    /**
     * Genera reporte de inventario en formato HTML (para imprimir como PDF)
     */
    public byte[] generarReporteInventarioPdf() throws IOException {
        List<Producto> productos = productoService.obtenerTodosProductos();
        
        StringBuilder html = new StringBuilder();
        
        // Encabezado HTML y estilos
        html.append("<!DOCTYPE html><html><head><meta charset='UTF-8'>");
        html.append("<title>Reporte de Inventario</title>");
        html.append("<style>");
        html.append("body { font-family: Arial, sans-serif; margin: 20px; }");
        html.append("h1 { color: #2c3e50; text-align: center; margin-bottom: 10px; border-bottom: 3px solid #27ae60; padding-bottom: 10px; }");
        html.append(".info { text-align: center; color: #7f8c8d; margin-bottom: 20px; font-size: 14px; }");
        
        // NUEVO: Estilos para el botón de descarga
        html.append(".download-container { position: fixed; top: 20px; right: 20px; z-index: 1000; }");
        html.append(".download-btn { background: linear-gradient(135deg, #27ae60 0%, #229954 100%); color: white; padding: 14px 28px; ");
        html.append("border: none; border-radius: 10px; cursor: pointer; font-size: 16px; font-weight: bold; ");
        html.append("box-shadow: 0 4px 15px rgba(39, 174, 96, 0.3); transition: all 0.3s ease; ");
        html.append("display: flex; align-items: center; gap: 10px; text-decoration: none; }");
        html.append(".download-btn:hover { background: linear-gradient(135deg, #229954 0%, #1e8449 100%); ");
        html.append("transform: translateY(-2px); box-shadow: 0 6px 20px rgba(39, 174, 96, 0.4); }");
        html.append(".download-btn:active { transform: translateY(0); }");
        html.append(".download-icon { font-size: 20px; }");
        
        // Estilos para ocultar el botón al imprimir
        html.append("@media print { .download-container { display: none; } body { margin: 10px; } @page { size: landscape; } }");
        
        html.append("table { width: 100%; border-collapse: collapse; margin-top: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }");
        html.append("th, td { border: 1px solid #bdc3c7; padding: 12px; text-align: left; font-size: 12px; }");
        html.append("th { background-color: #27ae60; color: white; font-weight: bold; text-transform: uppercase; }");
        html.append("tr:nth-child(even) { background-color: #ecf0f1; }");
        html.append("tr:hover { background-color: #d5dbdb; }");
        html.append(".agotado { background-color: #e74c3c; color: white; font-weight: bold; padding: 4px 8px; border-radius: 4px; }");
        html.append(".critico { background-color: #e67e22; color: white; font-weight: bold; padding: 4px 8px; border-radius: 4px; }");
        html.append(".normal { background-color: #27ae60; color: white; padding: 4px 8px; border-radius: 4px; }");
        html.append(".resumen { margin-top: 30px; padding: 20px; background-color: #ecf0f1; border-left: 5px solid #27ae60; border-radius: 5px; }");
        html.append(".resumen h3 { color: #2c3e50; margin-top: 0; }");
        html.append(".resumen p { margin: 8px 0; color: #34495e; }");
        html.append("</style></head><body>");
        
        // NUEVO: Botón de descarga flotante
        html.append("<div class='download-container'>");
        html.append("<button class='download-btn' onclick='descargarPDF()'>");
        html.append("<span class='download-icon'>📥</span>");
        html.append("<span>Descargar PDF</span>");
        html.append("</button>");
        html.append("</div>");
        
        // Título y fecha
        html.append("<h1>📦 REPORTE DE INVENTARIO - DROGUERÍA INTI</h1>");
        html.append("<div class='info'>");
        html.append("<strong>Fecha de Generación:</strong> ");
        html.append(LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
        html.append(" | <strong>Total de productos:</strong> ").append(productos.size());
        html.append("</div>");
        
        // Tabla de productos
        html.append("<table>");
        html.append("<thead><tr>");
        html.append("<th>Código</th><th>Nombre</th><th>Categoría</th><th>Stock Actual</th>");
        html.append("<th>Stock Mínimo</th><th>Precio Compra</th><th>Precio Venta</th>");
        html.append("<th>Valor Inventario</th><th>Estado</th>");
        html.append("</tr></thead>");
        html.append("<tbody>");
        
        double valorTotal = 0;
        int agotados = 0;
        int criticos = 0;
        int normales = 0;
        
        for (Producto p : productos) {
            String estado = determinarEstadoStock(p);
            String claseEstado = estado.equals("AGOTADO") ? "agotado" : 
                                (estado.equals("CRÍTICO") ? "critico" : "normal");
            
            if (estado.equals("AGOTADO")) agotados++;
            else if (estado.equals("CRÍTICO")) criticos++;
            else normales++;
            
            double valorInv = p.getPrecioCompra() != null ? 
                p.getStockActual() * p.getPrecioCompra().doubleValue() : 0;
            valorTotal += valorInv;
            
            html.append("<tr>");
            html.append("<td>").append(escaparHtml(p.getCodigo())).append("</td>");
            html.append("<td>").append(escaparHtml(p.getNombre())).append("</td>");
            html.append("<td>").append(
                p.getCategoria() != null ? escaparHtml(p.getCategoria().getNombre()) : "Sin categoría"
            ).append("</td>");
            html.append("<td style='text-align: center;'>").append(p.getStockActual()).append("</td>");
            html.append("<td style='text-align: center;'>").append(p.getStockMinimo()).append("</td>");
            html.append("<td style='text-align: right;'>S/ ").append(
                p.getPrecioCompra() != null ? String.format("%.2f", p.getPrecioCompra()) : "0.00"
            ).append("</td>");
            html.append("<td style='text-align: right;'>S/ ").append(
                p.getPrecioVenta() != null ? String.format("%.2f", p.getPrecioVenta()) : "0.00"
            ).append("</td>");
            html.append("<td style='text-align: right;'><strong>S/ ").append(String.format("%.2f", valorInv)).append("</strong></td>");
            html.append("<td style='text-align: center;'><span class='").append(claseEstado).append("'>")
                .append(estado).append("</span></td>");
            html.append("</tr>");
        }
        
        html.append("</tbody></table>");
        
        // Resumen
        html.append("<div class='resumen'>");
        html.append("<h3>📊 Resumen del Inventario</h3>");
        html.append("<p><strong>Total de Productos:</strong> ").append(productos.size()).append("</p>");
        html.append("<p><strong>Productos con Stock Bajo:</strong> ").append(criticos).append("</p>");
        html.append("<p><strong>Productos Agotados:</strong> ").append(agotados).append("</p>");
        html.append("<p><strong>Productos con Stock Normal:</strong> ").append(normales).append("</p>");
        html.append("<p><strong>Valor Total del Inventario:</strong> S/ ").append(String.format("%.2f", valorTotal)).append("</p>");
        html.append("</div>");
        
        // NUEVO: JavaScript para descargar el PDF
        html.append("<script>");
        html.append("function descargarPDF() {");
        html.append("  const button = document.querySelector('.download-btn');");
        html.append("  const originalContent = button.innerHTML;");
        html.append("  button.innerHTML = '<span class=\"download-icon\">⏳</span><span>Preparando descarga...</span>';");
        html.append("  button.disabled = true;");
        html.append("  setTimeout(() => { ");
        html.append("    window.print(); ");
        html.append("    setTimeout(() => { button.innerHTML = originalContent; button.disabled = false; }, 1000);");
        html.append("  }, 500);");
        html.append("}");
        html.append("</script>");
        
        html.append("</body></html>");
        
        return html.toString().getBytes("UTF-8");
    }
    
    /**
     * Genera reporte de stock bajo en formato HTML
     */
    public byte[] generarReporteStockBajoPdf() throws IOException {
        List<Producto> productos = productoService.obtenerProductosStockBajo();
        
        StringBuilder html = new StringBuilder();
        
        html.append("<!DOCTYPE html><html><head><meta charset='UTF-8'>");
        html.append("<title>Reporte de Stock Bajo</title>");
        html.append("<style>");
        html.append("body { font-family: Arial, sans-serif; margin: 20px; }");
        html.append("h1 { color: #e67e22; text-align: center; margin-bottom: 10px; border-bottom: 3px solid #e67e22; padding-bottom: 10px; }");
        html.append(".info { text-align: center; color: #7f8c8d; margin-bottom: 20px; }");
        
        // Botón de descarga
        html.append(".download-container { position: fixed; top: 20px; right: 20px; z-index: 1000; }");
        html.append(".download-btn { background: linear-gradient(135deg, #e67e22 0%, #d35400 100%); color: white; padding: 14px 28px; ");
        html.append("border: none; border-radius: 10px; cursor: pointer; font-size: 16px; font-weight: bold; ");
        html.append("box-shadow: 0 4px 15px rgba(230, 126, 34, 0.3); transition: all 0.3s ease; ");
        html.append("display: flex; align-items: center; gap: 10px; }");
        html.append(".download-btn:hover { background: linear-gradient(135deg, #d35400 0%, #ba4a00 100%); ");
        html.append("transform: translateY(-2px); box-shadow: 0 6px 20px rgba(230, 126, 34, 0.4); }");
        html.append("@media print { .download-container { display: none; } body { margin: 10px; } }");
        
        html.append("table { width: 100%; border-collapse: collapse; margin-top: 20px; }");
        html.append("th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }");
        html.append("th { background-color: #e67e22; color: white; }");
        html.append("tr:nth-child(even) { background-color: #f2f2f2; }");
        html.append(".alert { background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0; }");
        html.append(".resumen { margin-top: 20px; padding: 15px; background-color: #f9f9f9; border-left: 4px solid #e67e22; }");
        html.append("</style></head><body>");
        
        // Botón de descarga
        html.append("<div class='download-container'>");
        html.append("<button class='download-btn' onclick='window.print()'>");
        html.append("<span>📥</span><span>Descargar PDF</span>");
        html.append("</button>");
        html.append("</div>");
        
        html.append("<h1>⚠️ REPORTE DE STOCK BAJO - DROGUERÍA INTI</h1>");
        html.append("<div class='info'>");
        html.append("<strong>Fecha:</strong> ");
        html.append(LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
        html.append("</div>");
        
        html.append("<div class='alert'>");
        html.append("<strong>⚠️ ALERTA:</strong> Estos productos requieren atención inmediata para reposición");
        html.append("</div>");
        
        html.append("<table>");
        html.append("<thead><tr>");
        html.append("<th>Código</th><th>Nombre</th><th>Categoría</th>");
        html.append("<th>Stock Actual</th><th>Stock Mínimo</th><th>Déficit</th>");
        html.append("</tr></thead>");
        html.append("<tbody>");
        
        for (Producto p : productos) {
            int deficit = p.getStockMinimo() - p.getStockActual();
            
            html.append("<tr>");
            html.append("<td>").append(escaparHtml(p.getCodigo())).append("</td>");
            html.append("<td>").append(escaparHtml(p.getNombre())).append("</td>");
            html.append("<td>").append(
                p.getCategoria() != null ? escaparHtml(p.getCategoria().getNombre()) : "Sin categoría"
            ).append("</td>");
            html.append("<td style='text-align: center;'><strong>").append(p.getStockActual()).append("</strong></td>");
            html.append("<td style='text-align: center;'>").append(p.getStockMinimo()).append("</td>");
            html.append("<td style='text-align: center; color: red;'><strong>").append(deficit).append("</strong></td>");
            html.append("</tr>");
        }
        
        html.append("</tbody></table>");
        
        html.append("<div class='resumen'>");
        html.append("<h3>Resumen</h3>");
        html.append("<p><strong>Total de Productos con Stock Bajo:</strong> ").append(productos.size()).append("</p>");
        html.append("<p><em>Se recomienda realizar pedidos de reposición lo antes posible.</em></p>");
        html.append("</div>");
        
        html.append("</body></html>");
        
        return html.toString().getBytes("UTF-8");
    }
    
    // ============================================
    // REPORTES ADICIONALES (Temporales - devuelven reportes base)
    // ============================================
    
    public byte[] generarReporteVencimientosExcel() throws IOException {
        return generarReporteStockBajoExcel();
    }
    
    public byte[] generarReporteVencimientosPdf() throws IOException {
        return generarReporteStockBajoPdf();
    }
    
    public byte[] generarReporteRotacionExcel() throws IOException {
        return generarReporteInventarioExcel();
    }
    
    public byte[] generarReporteRotacionPdf() throws IOException {
        return generarReporteInventarioPdf();
    }
    
    public byte[] generarReporteComprasProveedorExcel() throws IOException {
        return generarReporteInventarioExcel();
    }
    
    public byte[] generarReporteComprasProveedorPdf() throws IOException {
        return generarReporteInventarioPdf();
    }
    
    public byte[] generarReporteMovimientosSemanalesExcel() throws IOException {
        return generarReporteInventarioExcel();
    }
    
    public byte[] generarReporteMovimientosSemanalesPdf() throws IOException {
        return generarReporteInventarioPdf();
    }
    
    public byte[] generarReporteMovimientosAnualesExcel() throws IOException {
        return generarReporteInventarioExcel();
    }
    
    public byte[] generarReporteMovimientosAnualesPdf() throws IOException {
        return generarReporteInventarioPdf();
    }
    
    // ============================================
    // MÉTODOS AUXILIARES - ESTILOS EXCEL
    // ============================================
    
    private CellStyle crearEstiloTitulo(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.DARK_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        return style;
    }
    
    private CellStyle crearEstiloFecha(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setItalic(true);
        font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }
    
    private CellStyle crearEstiloEncabezado(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        return style;
    }
    
    private CellStyle crearEstiloDatos(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }
    
    private CellStyle crearEstiloMoneda(Workbook workbook) {
        CellStyle style = crearEstiloDatos(workbook);
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("S/ #,##0.00"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }
    
    private CellStyle crearEstiloResumen(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
    
    private CellStyle crearEstiloEstado(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
    
    // ============================================
    // MÉTODOS AUXILIARES - LÓGICA
    // ============================================
    
    /**
     * Determina el estado del stock de un producto
     */
    private String determinarEstadoStock(Producto producto) {
        if (producto.getStockActual() == 0) {
            return "AGOTADO";
        } else if (producto.getStockActual() <= producto.getStockMinimo()) {
            return "CRÍTICO";
        } else {
            return "NORMAL";
        }
    }
    
    /**
     * Escapa caracteres especiales HTML para prevenir XSS
     */
    private String escaparHtml(String texto) {
        if (texto == null) {
            return "";
        }
        return texto.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace("\"", "&quot;")
                    .replace("'", "&#39;");
    }
}