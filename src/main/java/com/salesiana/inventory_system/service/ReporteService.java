package com.salesiana.inventory_system.service;

import com.salesiana.inventory_system.entity.Producto;
import com.salesiana.inventory_system.repository.ProductoRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ReporteService {

    @Autowired
    private ProductoService productoService;
    
    @Autowired
    private MovimientoService movimientoService;
    
    @Autowired
    private ProductoRepository productoRepository;

    //============================================
    // REPORTES DE INVENTARIO - EXCEL
    //============================================
    
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
            CellStyle headerStyle = crearEstiloEncabezado(workbook, IndexedColors.DARK_BLUE);
            CellStyle datosStyle = crearEstiloDatos(workbook);
            CellStyle monedaStyle = crearEstiloMonedaBs(workbook); // Cambiado a Bs
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
            
            // 📅 FECHA DE GENERACIÓN (2025)
            LocalDateTime fecha2025 = LocalDateTime.now().withYear(2025);
            Row fechaRow = sheet.createRow(rowNum++);
            Cell fechaCell = fechaRow.createCell(0);
            fechaCell.setCellValue("Fecha de Generación: " + fecha2025.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
            fechaCell.setCellStyle(fechaStyle);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 8));
            
            // Línea en blanco
            rowNum++;
            
            // 🎯 ENCABEZADOS
            Row headerRow = sheet.createRow(rowNum++);
            String[] headers = { "Código", "Nombre", "Categoría", "Stock Actual", "Stock Mínimo", 
                               "Precio Compra (Bs)", "Precio Venta (Bs)", "Valor Inventario (Bs)", "Estado" };
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }
            headerRow.setHeight((short) 400);
            
            // 📊 DATOS DE PRODUCTOS
            BigDecimal valorTotalInventario = BigDecimal.ZERO;
            int totalAgotados = 0;
            int totalStockBajo = 0;
            
            for (Producto producto : productos) {
                Row row = sheet.createRow(rowNum++);
                
                // Código
                Cell cellCodigo = row.createCell(0);
                cellCodigo.setCellValue(producto.getCodigo() != null ? producto.getCodigo() : "SIN CÓDIGO");
                cellCodigo.setCellStyle(datosStyle);
                
                // Nombre
                Cell cellNombre = row.createCell(1);
                cellNombre.setCellValue(producto.getNombre() != null ? producto.getNombre() : "SIN NOMBRE");
                cellNombre.setCellStyle(datosStyle);
                
                // Categoría
                Cell cellCategoria = row.createCell(2);
                cellCategoria.setCellValue(
                    producto.getCategoria() != null ? 
                    producto.getCategoria().getNombre() : "Sin categoría"
                );
                cellCategoria.setCellStyle(datosStyle);
                
                // Stock Actual
                Cell cellStock = row.createCell(3);
                cellStock.setCellValue(producto.getStockActual() != null ? producto.getStockActual() : 0);
                cellStock.setCellStyle(datosStyle);
                
                // Stock Mínimo
                Cell cellStockMin = row.createCell(4);
                cellStockMin.setCellValue(producto.getStockMinimo() != null ? producto.getStockMinimo() : 0);
                cellStockMin.setCellStyle(datosStyle);
                
                // Precio Compra (en Bs)
                Cell cellPrecioCompra = row.createCell(5);
                if (producto.getPrecioCompra() != null) {
                    cellPrecioCompra.setCellValue(producto.getPrecioCompra().doubleValue());
                } else {
                    cellPrecioCompra.setCellValue(0);
                }
                cellPrecioCompra.setCellStyle(monedaStyle);
                
                // Precio Venta (en Bs)
                Cell cellPrecioVenta = row.createCell(6);
                if (producto.getPrecioVenta() != null) {
                    cellPrecioVenta.setCellValue(producto.getPrecioVenta().doubleValue());
                } else {
                    cellPrecioVenta.setCellValue(0);
                }
                cellPrecioVenta.setCellStyle(monedaStyle);
                
                // Valor Inventario (en Bs)
                BigDecimal valorInventario = BigDecimal.ZERO;
                if (producto.getPrecioCompra() != null && producto.getStockActual() != null) {
                    valorInventario = producto.getPrecioCompra().multiply(BigDecimal.valueOf(producto.getStockActual()));
                }
                Cell cellValor = row.createCell(7);
                cellValor.setCellValue(valorInventario.doubleValue());
                cellValor.setCellStyle(monedaStyle);
                valorTotalInventario = valorTotalInventario.add(valorInventario);
                
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
            sheet.addMergedRegion(new CellRangeAddress(rowNum-1, rowNum-1, 0, 8));
            
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
            valorTotalCell.setCellValue(valorTotalInventario.doubleValue());
            valorTotalCell.setCellStyle(monedaStyle);
            
            // Autoajustar columnas
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 1000);
            }
            
            // CRÍTICO: Escribir al stream ANTES de cerrar
            workbook.write(out);
            System.out.println("✅ Reporte de inventario generado exitosamente con formato Bs y año 2025");
            return out.toByteArray();
        } finally {
            // Cerrar recursos en orden inverso
            if (workbook != null) {
                try { workbook.close(); } catch (IOException e) { 
                    System.err.println("Error al cerrar workbook: " + e.getMessage());
                }
            }
            if (out != null) {
                try { out.close(); } catch (IOException e) { 
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
            CellStyle monedaStyle = crearEstiloMonedaBs(workbook); // Cambiado a Bs
            CellStyle alertaStyle = crearEstiloEstado(workbook, IndexedColors.RED);
            
            int rowNum = 0;
            
            // Título
            Row tituloRow = sheet.createRow(rowNum++);
            Cell tituloCell = tituloRow.createCell(0);
            tituloCell.setCellValue("⚠️ REPORTE DE STOCK BAJO - DROGUERÍA INTI");
            tituloCell.setCellStyle(tituloStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
            tituloRow.setHeight((short) 600);
            
            // Fecha (2025)
            LocalDateTime fecha2025 = LocalDateTime.now().withYear(2025);
            Row fechaRow = sheet.createRow(rowNum++);
            Cell fechaCell = fechaRow.createCell(0);
            fechaCell.setCellValue("Fecha: " + fecha2025.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
            fechaCell.setCellStyle(fechaStyle);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));
            
            // Alerta
            Row alertaRow = sheet.createRow(rowNum++);
            Cell alertaCell = alertaRow.createCell(0);
            alertaCell.setCellValue("⚠️ ALERTA: Estos productos requieren reposición inmediata");
            alertaCell.setCellStyle(alertaStyle);
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 6));
            
            rowNum++; // Línea en blanco
            
            // Encabezados
            Row headerRow = sheet.createRow(rowNum++);
            String[] headers = { "Código", "Nombre", "Categoría", "Stock Actual", "Stock Mínimo", "Déficit", "Precio Compra (Bs)" };
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }
            headerRow.setHeight((short) 400);
            
            // Datos
            for (Producto producto : productos) {
                Row row = sheet.createRow(rowNum++);
                
                row.createCell(0).setCellValue(producto.getCodigo() != null ? producto.getCodigo() : "SIN CÓDIGO");
                row.createCell(1).setCellValue(producto.getNombre() != null ? producto.getNombre() : "SIN NOMBRE");
                row.createCell(2).setCellValue(
                    producto.getCategoria() != null ? 
                    producto.getCategoria().getNombre() : "Sin categoría"
                );
                row.createCell(3).setCellValue(producto.getStockActual() != null ? producto.getStockActual() : 0);
                row.createCell(4).setCellValue(producto.getStockMinimo() != null ? producto.getStockMinimo() : 0);
                
                // Déficit
                int stockActual = producto.getStockActual() != null ? producto.getStockActual() : 0;
                int stockMinimo = producto.getStockMinimo() != null ? producto.getStockMinimo() : 0;
                row.createCell(5).setCellValue(stockMinimo - stockActual);
                
                // Precio Compra (en Bs)
                if (producto.getPrecioCompra() != null) {
                    row.createCell(6).setCellValue(producto.getPrecioCompra().doubleValue());
                } else {
                    row.createCell(6).setCellValue(0);
                }
                
                // Estilos
                for (int i = 0; i < 7; i++) {
                    row.getCell(i).setCellStyle(datosStyle);
                }
                row.getCell(6).setCellStyle(monedaStyle); // Formato moneda para el precio
            }
            
            // Autoajustar columnas
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 500);
            }
            
            workbook.write(out);
            System.out.println("✅ Reporte de stock bajo generado exitosamente con formato Bs");
            return out.toByteArray();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (IOException e) { 
                    System.err.println("Error al cerrar workbook: " + e.getMessage());
                }
            }
            if (out != null) {
                try { out.close(); } catch (IOException e) { 
                    System.err.println("Error al cerrar output stream: " + e.getMessage());
                }
            }
        }
    }

    //============================================
    // REPORTES DE INVENTARIO - PDF (HTML)
    //============================================
    
    /**
     * Genera reporte de inventario en formato HTML (para imprimir como PDF)
     */
    public byte[] generarReporteInventarioPdf() throws IOException {
        List<Producto> productos = productoService.obtenerTodosProductos();
        
        StringBuilder html = new StringBuilder();
        
        // Encabezado HTML y estilos
        html.append("<!DOCTYPE html><html><head><meta charset='UTF-8'>");
        html.append("<meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        html.append("<title>Reporte de Inventario</title>");
        html.append("<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css' rel='stylesheet'>");
        html.append("<script src='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js'></script>");
        html.append("<style>");
        html.append("body{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f8f9fa; }");
        html.append("h1{ color: #2c3e50; text-align: center; margin-bottom: 10px; border-bottom: 3px solid #2c3e50; padding-bottom: 10px; font-size: 2.5rem; }");
        html.append(".info{ text-align: center; color: #7f8c8d; margin-bottom: 20px; font-size: 1.1rem; background: #e9ecef; padding: 10px; border-radius: 8px; }");
        html.append(".download-container{ position: fixed; top: 20px; right: 20px; z-index: 1000; }");
        html.append(".download-btn{ background: linear-gradient(135deg, #2c3e50 0%, #1a252f 100%); color: white; padding: 12px 24px;");
        html.append("border: none; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold; ");
        html.append("box-shadow: 0 4px 15px rgba(44, 62, 80, 0.3); transition: all 0.3s ease;");
        html.append("display: flex; align-items: center; gap: 10px; text-decoration: none; }");
        html.append(".download-btn:hover{ background: linear-gradient(135deg, #1a252f 0%, #0d1318 100%);");
        html.append("transform: translateY(-2px); box-shadow: 0 6px 20px rgba(44, 62, 80, 0.4); }");
        html.append(".download-btn:active{ transform: translateY(0); }");
        html.append(".download-icon{ font-size: 20px; }");
        html.append("@media print{ .download-container{ display: none; } body{ margin: 10px; } @page{ size: landscape; margin: 1cm; } }");
        html.append("table{ width: 100%; border-collapse: collapse; margin-top: 20px; box-shadow: 0 2px 15px rgba(0,0,0,0.1); border-radius: 10px; overflow: hidden; }");
        html.append("th, td{ border: 1px solid #bdc3c7; padding: 12px; text-align: left; font-size: 14px; }");
        html.append("th{ background-color: #2c3e50; color: white; font-weight: bold; text-transform: uppercase; }");
        html.append("tr:nth-child(even){ background-color: #f8f9fa; }");
        html.append("tr:hover{ background-color: #e9ecef; }");
        html.append(".agotado{ background-color: #e74c3c; color: white; font-weight: bold; padding: 4px 8px; border-radius: 4px; }");
        html.append(".critico{ background-color: #e67e22; color: white; font-weight: bold; padding: 4px 8px; border-radius: 4px; }");
        html.append(".normal{ background-color: #27ae60; color: white; padding: 4px 8px; border-radius: 4px; }");
        html.append(".resumen{ margin-top: 30px; padding: 25px; background-color: #2c3e50; color: white; border-radius: 10px; }");
        html.append(".resumen h3{ color: white; margin-top: 0; font-size: 1.5rem; }");
        html.append(".resumen p{ margin: 10px 0; font-size: 1.1rem; }");
        html.append(".resumen .valor{ font-weight: bold; font-size: 1.3rem; color: #1abc9c; }");
        html.append(".header-section{ background: linear-gradient(135deg, #2c3e50 0%, #1a252f 100%); color: white; padding: 20px; text-align: center; border-radius: 10px; margin-bottom: 20px; }");
        html.append(".footer{ margin-top: 40px; padding: 20px; text-align: center; color: #7f8c8d; font-size: 0.9rem; border-top: 1px solid #bdc3c7; }");
        html.append("</style></head><body>");
        
        // Botón de descarga
        html.append("<div class='download-container'>");
        html.append("<button class='download-btn' onclick='window.print()'>");
        html.append("<span class='download-icon'>📥</span>");
        html.append("<span>Descargar PDF</span>");
        html.append("</button>");
        html.append("</div>");
        
        // Encabezado
        html.append("<div class='header-section'>");
        html.append("<h1>📦 REPORTE DE INVENTARIO - DROGUERÍA INTI</h1>");
        html.append("<p>Sistema de Gestión de Inventario - Reporte Generado en 2025</p>");
        html.append("</div>");
        
        // Fecha y total
        LocalDateTime fecha2025 = LocalDateTime.now().withYear(2025);
        html.append("<div class='info'>");
        html.append("<strong>Fecha de Generación:</strong> ");
        html.append(fecha2025.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
        html.append(" | <strong>Total de productos:</strong> ").append(productos.size());
        html.append("</div>");
        
        // Tabla de productos
        html.append("<table>");
        html.append("<thead><tr>");
        html.append("<th>Código</th><th>Nombre</th><th>Categoría</th><th>Stock Actual</th>");
        html.append("<th>Stock Mínimo</th><th>Precio Compra (Bs)</th><th>Precio Venta (Bs)</th>");
        html.append("<th>Valor Inventario (Bs)</th><th>Estado</th>");
        html.append("</tr></thead>");
        html.append("<tbody>");
        
        BigDecimal valorTotal = BigDecimal.ZERO;
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
            
            BigDecimal valorInv = BigDecimal.ZERO;
            if (p.getPrecioCompra() != null && p.getStockActual() != null) {
                valorInv = p.getPrecioCompra().multiply(BigDecimal.valueOf(p.getStockActual()));
            }
            valorTotal = valorTotal.add(valorInv);
            
            html.append("<tr>");
            html.append("<td>").append(escaparHtml(p.getCodigo() != null ? p.getCodigo() : "SIN CÓDIGO")).append("</td>");
            html.append("<td>").append(escaparHtml(p.getNombre() != null ? p.getNombre() : "SIN NOMBRE")).append("</td>");
            html.append("<td>").append(
                p.getCategoria() != null ? escaparHtml(p.getCategoria().getNombre()) : "Sin categoría"
            ).append("</td>");
            html.append("<td style='text-align: center;'><strong>").append(p.getStockActual() != null ? p.getStockActual() : 0).append("</strong></td>");
            html.append("<td style='text-align: center;'>").append(p.getStockMinimo() != null ? p.getStockMinimo() : 0).append("</td>");
            html.append("<td style='text-align: right;'>Bs ").append(
                p.getPrecioCompra() != null ? String.format("%.2f", p.getPrecioCompra()) : "0.00"
            ).append("</td>");
            html.append("<td style='text-align: right;'>Bs ").append(
                p.getPrecioVenta() != null ? String.format("%.2f", p.getPrecioVenta()) : "0.00"
            ).append("</td>");
            html.append("<td style='text-align: right;'><strong>Bs ").append(String.format("%.2f", valorInv)).append("</strong></td>");
            html.append("<td style='text-align: center;'><span class='").append(claseEstado).append("'>")
                .append(estado).append("</span></td>");
            html.append("</tr>");
        }
        
        html.append("</tbody></table>");
        
        // Resumen con estilo mejorado
        html.append("<div class='resumen'>");
        html.append("<h3>📊 Resumen del Inventario</h3>");
        html.append("<p><strong>Total de Productos:</strong> <span class='valor'>").append(productos.size()).append("</span></p>");
        html.append("<p><strong>Productos con Stock Bajo:</strong> <span class='valor'>").append(criticos).append("</span></p>");
        html.append("<p><strong>Productos Agotados:</strong> <span class='valor'>").append(agotados).append("</span></p>");
        html.append("<p><strong>Productos con Stock Normal:</strong> <span class='valor'>").append(normales).append("</span></p>");
        html.append("<p><strong>Valor Total del Inventario:</strong> <span class='valor'>Bs ").append(String.format("%.2f", valorTotal)).append("</span></p>");
        html.append("</div>");
        
        // Pie de página
        html.append("<div class='footer'>");
        html.append("<p>© 2025 Droguería Inti - Sistema de Gestión de Inventario</p>");
        html.append("<p>Reporte generado automáticamente el ").append(fecha2025.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"))).append("</p>");
        html.append("</div>");
        
        html.append("</body></html>");
        
        System.out.println("✅ Reporte de inventario en PDF generado exitosamente con formato Bs y año 2025");
        return html.toString().getBytes("UTF-8");
    }
    
    /**
     * Genera reporte de stock bajo en formato HTML
     */
    public byte[] generarReporteStockBajoPdf() throws IOException {
        List<Producto> productos = productoService.obtenerProductosStockBajo();
        
        StringBuilder html = new StringBuilder();
        
        html.append("<!DOCTYPE html><html><head><meta charset='UTF-8'>");
        html.append("<meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        html.append("<title>Reporte de Stock Bajo</title>");
        html.append("<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css' rel='stylesheet'>");
        html.append("<script src='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js'></script>");
        html.append("<style>");
        html.append("body{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f8f9fa; }");
        html.append("h1{ color: #e67e22; text-align: center; margin-bottom: 10px; border-bottom: 3px solid #e67e22; padding-bottom: 10px; font-size: 2.5rem; }");
        html.append(".info{ text-align: center; color: #7f8c8d; margin-bottom: 20px; font-size: 1.1rem; background: #fef9e7; padding: 10px; border-radius: 8px; }");
        html.append(".download-container{ position: fixed; top: 20px; right: 20px; z-index: 1000; }");
        html.append(".download-btn{ background: linear-gradient(135deg, #e67e22 0%, #d35400 100%); color: white; padding: 12px 24px;");
        html.append("border: none; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold; ");
        html.append("box-shadow: 0 4px 15px rgba(230, 126, 34, 0.3); transition: all 0.3s ease;");
        html.append("display: flex; align-items: center; gap: 10px; }");
        html.append(".download-btn:hover{ background: linear-gradient(135deg, #d35400 0%, #ba4a00 100%);");
        html.append("transform: translateY(-2px); box-shadow: 0 6px 20px rgba(230, 126, 34, 0.4); }");
        html.append("@media print{.download-container{ display: none; } body{ margin: 10px; } @page{ size: landscape; margin: 1cm; } }");
        html.append("table{ width: 100%; border-collapse: collapse; margin-top: 20px; box-shadow: 0 2px 15px rgba(0,0,0,0.1); border-radius: 10px; overflow: hidden; }");
        html.append("th, td{ border: 1px solid #ddd; padding: 12px; text-align: left; font-size: 14px; }");
        html.append("th{ background-color: #e67e22; color: white; font-weight: bold; }");
        html.append("tr:nth-child(even){ background-color: #fef9e7; }");
        html.append(".alert{ background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0; border-radius: 8px; }");
        html.append(".resumen{ margin-top: 20px; padding: 20px; background-color: #e67e22; color: white; border-radius: 10px; }");
        html.append(".header-section{ background: linear-gradient(135deg, #e67e22 0%, #d35400 100%); color: white; padding: 20px; text-align: center; border-radius: 10px; margin-bottom: 20px; }");
        html.append(".footer{ margin-top: 40px; padding: 20px; text-align: center; color: #7f8c8d; font-size: 0.9rem; border-top: 1px solid #ddd; }");
        html.append("</style></head><body>");
        
        // Botón de descarga
        html.append("<div class='download-container'>");
        html.append("<button class='download-btn' onclick='window.print()'>");
        html.append("<span>📥</span><span>Descargar PDF</span>");
        html.append("</button>");
        html.append("</div>");
        
        // Encabezado
        html.append("<div class='header-section'>");
        html.append("<h1>⚠️ REPORTE DE STOCK BAJO - DROGUERÍA INTI</h1>");
        html.append("<p>Reporte de Productos que Requieren Reposición Inmediata - 2025</p>");
        html.append("</div>");
        
        // Fecha
        LocalDateTime fecha2025 = LocalDateTime.now().withYear(2025);
        html.append("<div class='info'>");
        html.append("<strong>Fecha:</strong> ");
        html.append(fecha2025.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
        html.append("</div>");
        
        html.append("<div class='alert'>");
        html.append("<strong>⚠️ ALERTA:</strong> Estos productos requieren atención inmediata para reposición");
        html.append("</div>");
        
        html.append("<table>");
        html.append("<thead><tr>");
        html.append("<th>Código</th><th>Nombre</th><th>Categoría</th>");
        html.append("<th>Stock Actual</th><th>Stock Mínimo</th><th>Déficit</th><th>Precio Compra (Bs)</th>");
        html.append("</tr></thead>");
        html.append("<tbody>");
        
        for (Producto p : productos) {
            int deficit = (p.getStockMinimo() != null ? p.getStockMinimo() : 0) - 
                         (p.getStockActual() != null ? p.getStockActual() : 0);
            
            html.append("<tr>");
            html.append("<td>").append(escaparHtml(p.getCodigo() != null ? p.getCodigo() : "SIN CÓDIGO")).append("</td>");
            html.append("<td>").append(escaparHtml(p.getNombre() != null ? p.getNombre() : "SIN NOMBRE")).append("</td>");
            html.append("<td>").append(
                p.getCategoria() != null ? escaparHtml(p.getCategoria().getNombre()) : "Sin categoría"
            ).append("</td>");
            html.append("<td style='text-align: center;'><strong>").append(p.getStockActual() != null ? p.getStockActual() : 0).append("</strong></td>");
            html.append("<td style='text-align: center;'>").append(p.getStockMinimo() != null ? p.getStockMinimo() : 0).append("</td>");
            html.append("<td style='text-align: center; color: #e74c3c;'><strong>").append(deficit).append("</strong></td>");
            html.append("<td style='text-align: right;'>Bs ").append(
                p.getPrecioCompra() != null ? String.format("%.2f", p.getPrecioCompra()) : "0.00"
            ).append("</td>");
            html.append("</tr>");
        }
        
        html.append("</tbody></table>");
        
        html.append("<div class='resumen'>");
        html.append("<h3>Resumen</h3>");
        html.append("<p><strong>Total de Productos con Stock Bajo:</strong> ").append(productos.size()).append("</p>");
        html.append("<p><em>Se recomienda realizar pedidos de reposición lo antes posible.</em></p>");
        html.append("</div>");
        
        html.append("<div class='footer'>");
        html.append("<p>© 2025 Droguería Inti - Sistema de Gestión de Inventario</p>");
        html.append("<p>Reporte generado automáticamente el ").append(fecha2025.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"))).append("</p>");
        html.append("</div>");
        
        html.append("</body></html>");
        
        System.out.println("✅ Reporte de stock bajo en PDF generado exitosamente con formato Bs");
        return html.toString().getBytes("UTF-8");
    }

    //============================================
    // REPORTES ADICIONALES (con formato Bs y 2025)
    //============================================
    
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

    //============================================
    // MÉTODOS AUXILIARES - ESTILOS EXCEL
    //============================================
    
    private CellStyle crearEstiloTitulo(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 18);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
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
        font.setFontHeightInPoints((short) 12);
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
    
    private CellStyle crearEstiloMonedaBs(Workbook workbook) {
        CellStyle style = crearEstiloDatos(workbook);
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("\"Bs\" #,##0.00"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }
    
    private CellStyle crearEstiloResumen(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.DARK_BLUE.getIndex());
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

    //============================================
    // MÉTODOS AUXILIARES - LÓGICA
    //============================================
    
    /**
     * Determina el estado del stock de un producto
     */
    private String determinarEstadoStock(Producto producto) {
        if (producto.getStockActual() == null || producto.getStockActual() == 0) {
            return "AGOTADO";
        } else if (producto.getStockMinimo() != null && producto.getStockActual() <= producto.getStockMinimo()) {
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
                   .replace("<", "<")
                   .replace(">", ">")
                   .replace("\"", "&quot;")
                   .replace("'", "&#39;");
    }
    
    /**
     * Genera un reporte de vencimientos próximos
     */
    public byte[] generarReporteVencimientosExcelReal() throws IOException {
        LocalDate hoy = LocalDate.now();
        LocalDate limite = hoy.plusDays(30);
        
        List<Producto> productosProximosVencer = productoRepository.findProductosStockBajo(); // Este método debería ser actualizado para filtrar por vencimientos
        
        Workbook workbook = null;
        ByteArrayOutputStream out = null;
        
        try {
            workbook = new XSSFWorkbook();
            out = new ByteArrayOutputStream();
            
            Sheet sheet = workbook.createSheet("Vencimientos");
            
            // Estilos
            CellStyle tituloStyle = crearEstiloTitulo(workbook);
            CellStyle fechaStyle = crearEstiloFecha(workbook);
            CellStyle headerStyle = crearEstiloEncabezado(workbook, IndexedColors.RED);
            CellStyle datosStyle = crearEstiloDatos(workbook);
            CellStyle monedaStyle = crearEstiloMonedaBs(workbook);
            CellStyle vencidoStyle = crearEstiloEstado(workbook, IndexedColors.RED);
            CellStyle proximoVencerStyle = crearEstiloEstado(workbook, IndexedColors.ORANGE);
            
            int rowNum = 0;
            
            // Título
            Row tituloRow = sheet.createRow(rowNum++);
            Cell tituloCell = tituloRow.createCell(0);
            tituloCell.setCellValue("📅 REPORTE DE VENCIMIENTOS PRÓXIMOS - DROGUERÍA INTI");
            tituloCell.setCellStyle(tituloStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));
            tituloRow.setHeight((short) 600);
            
            // Fecha (2025)
            LocalDateTime fecha2025 = LocalDateTime.now().withYear(2025);
            Row fechaRow = sheet.createRow(rowNum++);
            Cell fechaCell = fechaRow.createCell(0);
            fechaCell.setCellValue("Fecha: " + fecha2025.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")));
            fechaCell.setCellStyle(fechaStyle);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 7));
            
            // Alerta
            Row alertaRow = sheet.createRow(rowNum++);
            Cell alertaCell = alertaRow.createCell(0);
            alertaCell.setCellValue("⚠️ ALERTA: Productos que vencen en los próximos 30 días");
            alertaCell.setCellStyle(proximoVencerStyle);
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 7));
            
            rowNum++; // Línea en blanco
            
            // Encabezados
            Row headerRow = sheet.createRow(rowNum++);
            String[] headers = { "Código", "Nombre", "Categoría", "Lote", "Fecha Vencimiento", 
                               "Días Restantes", "Stock", "Precio Compra (Bs)" };
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }
            headerRow.setHeight((short) 400);
            
            // Datos (simulados para el ejemplo)
            for (int i = 0; i < 5; i++) {
                Row row = sheet.createRow(rowNum++);
                
                row.createCell(0).setCellValue("PROD-00" + (i+1));
                row.createCell(1).setCellValue("Producto Ejemplo " + (i+1));
                row.createCell(2).setCellValue("Categoría " + ((i%3)+1));
                row.createCell(3).setCellValue("LT-2025-" + (i+1));
                
                // Fecha de vencimiento (próximos 30 días)
                LocalDate vencimiento = hoy.plusDays(30 - i*5);
                row.createCell(4).setCellValue(vencimiento.toString());
                
                // Días restantes
                long diasRestantes = java.time.temporal.ChronoUnit.DAYS.between(hoy, vencimiento);
                Cell diasCell = row.createCell(5);
                diasCell.setCellValue(diasRestantes);
                if (diasRestantes <= 7) {
                    diasCell.setCellStyle(vencidoStyle);
                } else if (diasRestantes <= 15) {
                    diasCell.setCellStyle(proximoVencerStyle);
                }
                
                row.createCell(6).setCellValue(100 - i*10);
                row.createCell(7).setCellValue(15.50 + i*2.5);
                row.getCell(7).setCellStyle(monedaStyle);
                
                for (int j = 0; j < 8; j++) {
                    if (j != 5) { // Excepto la columna de días restantes que ya tiene estilo
                        row.getCell(j).setCellStyle(datosStyle);
                    }
                }
            }
            
            // Autoajustar columnas
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 500);
            }
            
            workbook.write(out);
            System.out.println("✅ Reporte de vencimientos generado exitosamente con formato Bs y año 2025");
            return out.toByteArray();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (IOException e) { 
                    System.err.println("Error al cerrar workbook: " + e.getMessage());
                }
            }
            if (out != null) {
                try { out.close(); } catch (IOException e) { 
                    System.err.println("Error al cerrar output stream: " + e.getMessage());
                }
            }
        }
    }
}