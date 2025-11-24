package com.salesiana.inventory_system.controller;

import com.salesiana.inventory_system.entity.TransferenciaUbicacion;
import com.salesiana.inventory_system.service.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.format.annotation.DateTimeFormat;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import java.time.LocalDateTime;
import java.util.Collections;
import java.util.List;

@Controller
@RequestMapping("/transferencias-ubicacion")
public class TransferenciaUbicacionController {

    @Autowired
    private TransferenciaUbicacionService transferenciaService;
    
    @Autowired
    private ProductoService productoService;
    
    @Autowired
    private LoteService loteService;
    
    @Autowired
    private UbicacionAlmacenService ubicacionService;
    
    @Autowired
    private UsuarioService usuarioService;

    @GetMapping
public String listarTransferencias(
        @RequestParam(required = false) Integer productoId,
        @RequestParam(required = false) @DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME) LocalDateTime fechaInicio,
        @RequestParam(required = false) @DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME) LocalDateTime fechaFin,
        Model model) {
    try {
        List<TransferenciaUbicacion> transferencias;
        
        if (productoId != null) {
            transferencias = transferenciaService.obtenerPorProducto(productoId);
            model.addAttribute("productoSeleccionado", 
                productoService.obtenerProductoPorId(productoId).orElse(null));
        } else if (fechaInicio != null && fechaFin != null) {
            transferencias = obtenerTransferenciasPorRangoFechas(fechaInicio, fechaFin);
        } else {
            transferencias = transferenciaService.obtenerUltimasTransferencias();
        }

        // CALCULAR ESTADÍSTICAS EN JAVA (sin usar estado que no existe)
        long productosUnicos = 0;

        if (transferencias != null && !transferencias.isEmpty()) {
            // Contar productos únicos usando Set
            java.util.Set<Integer> productosSet = new java.util.HashSet<>();
            for (TransferenciaUbicacion t : transferencias) {
                if (t.getProducto() != null && t.getProducto().getId() != null) {
                    productosSet.add(t.getProducto().getId());
                }
            }
            productosUnicos = productosSet.size();
        }

        // Pasar estadísticas calculadas al modelo
        model.addAttribute("transferencias", transferencias != null ? transferencias : Collections.emptyList());
        model.addAttribute("productosUnicos", productosUnicos);
        model.addAttribute("productos", productoService.obtenerTodosProductos());
        model.addAttribute("ubicaciones", ubicacionService.obtenerTodasUbicaciones());

        System.out.println("✅ Transferencias cargadas: " + (transferencias != null ? transferencias.size() : 0));
        System.out.println("📊 Estadísticas - Productos únicos: " + productosUnicos);
    } catch (Exception e) {
        System.err.println("❌ Error al listar transferencias: " + e.getMessage());
        e.printStackTrace();
        model.addAttribute("error", "Error al cargar las transferencias: " + e.getMessage());
        model.addAttribute("transferencias", Collections.emptyList());
        model.addAttribute("productosUnicos", 0L);
        model.addAttribute("productos", Collections.emptyList());
        model.addAttribute("ubicaciones", Collections.emptyList());
    }
    return "transferencias/lista";
}
    
    private List<TransferenciaUbicacion> obtenerTransferenciasPorRangoFechas(LocalDateTime fechaInicio, LocalDateTime fechaFin) {
        // TODO: Implementar método en el servicio para filtrar por rango de fechas
        return transferenciaService.obtenerUltimasTransferencias();
    }

    @GetMapping("/nuevo")
    public String mostrarFormularioNuevo(
            @RequestParam(required = false) Integer productoId,
            @RequestParam(required = false) Integer loteId,
            Model model) {
        
        try {
            TransferenciaUbicacion transferencia = new TransferenciaUbicacion();
            
            // Si viene productoId, pre-seleccionarlo
            if (productoId != null) {
                productoService.obtenerProductoPorId(productoId).ifPresent(transferencia::setProducto);
            }
            
            // Si viene loteId, pre-seleccionarlo
            if (loteId != null) {
                loteService.obtenerLotePorId(loteId).ifPresent(transferencia::setLote);
            }
            
            model.addAttribute("transferencia", transferencia);
            model.addAttribute("productos", productoService.obtenerTodosProductos());
            model.addAttribute("ubicaciones", ubicacionService.obtenerTodasUbicaciones());
            
            System.out.println("✅ Formulario de transferencia cargado");
        } catch (Exception e) {
            System.err.println("❌ Error al cargar formulario de transferencia: " + e.getMessage());
            e.printStackTrace();
            model.addAttribute("error", "Error al cargar el formulario: " + e.getMessage());
        }
        
        return "transferencias/form";
    }

    @PostMapping("/guardar")
    public String guardarTransferencia(@ModelAttribute TransferenciaUbicacion transferencia, RedirectAttributes redirectAttributes) {
        try {
            System.out.println("=== GUARDANDO TRANSFERENCIA ===");
            System.out.println("Producto ID: " + (transferencia.getProducto() != null ? transferencia.getProducto().getId() : "null"));
            System.out.println("Ubicación Origen ID: " + (transferencia.getUbicacionOrigen() != null ? transferencia.getUbicacionOrigen().getId() : "null"));
            System.out.println("Ubicación Destino ID: " + (transferencia.getUbicacionDestino() != null ? transferencia.getUbicacionDestino().getId() : "null"));
            System.out.println("Cantidad: " + transferencia.getCantidad());
            
            // Validar datos básicos
            if (transferencia.getProducto() == null || transferencia.getProducto().getId() == null) {
                redirectAttributes.addFlashAttribute("error", "Debe seleccionar un producto");
                return "redirect:/transferencias-ubicacion/nuevo";
            }
            
            if (transferencia.getUbicacionDestino() == null || transferencia.getUbicacionDestino().getId() == null) {
                redirectAttributes.addFlashAttribute("error", "Debe seleccionar una ubicación destino");
                return "redirect:/transferencias-ubicacion/nuevo";
            }
            
            if (transferencia.getCantidad() == null || transferencia.getCantidad() <= 0) {
                redirectAttributes.addFlashAttribute("error", "La cantidad debe ser mayor a 0");
                return "redirect:/transferencias-ubicacion/nuevo";
            }
            
            // Validar que ubicación origen y destino sean diferentes
            if (transferencia.getUbicacionOrigen() != null && 
                transferencia.getUbicacionOrigen().getId() != null && 
                transferencia.getUbicacionOrigen().getId().equals(transferencia.getUbicacionDestino().getId())) {
                redirectAttributes.addFlashAttribute("error", "La ubicación de origen y destino no pueden ser iguales");
                return "redirect:/transferencias-ubicacion/nuevo";
            }
            
            // Verificar si el lote pertenece al producto
            if (transferencia.getLote() != null && transferencia.getLote().getId() != null) {
                if (!transferencia.getLote().getProducto().getId().equals(transferencia.getProducto().getId())) {
                    redirectAttributes.addFlashAttribute("error", "El lote seleccionado no pertenece al producto seleccionado");
                    return "redirect:/transferencias-ubicacion/nuevo";
                }
            }
            
            // Registrar transferencia
            transferenciaService.registrarTransferencia(transferencia);
            
            System.out.println("✅ Transferencia registrada exitosamente");
            redirectAttributes.addFlashAttribute("success", "Transferencia registrada exitosamente");
            return "redirect:/transferencias-ubicacion";
        } catch (Exception e) {
            System.err.println("❌ Error al registrar transferencia: " + e.getMessage());
            e.printStackTrace();
            redirectAttributes.addFlashAttribute("error", "Error al registrar transferencia: " + e.getMessage());
            return "redirect:/transferencias-ubicacion/nuevo";
        }
    }
    
    @PostMapping("/buscar")
    public String buscarTransferencias(
            @RequestParam(required = false) @DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME) LocalDateTime fechaInicio,
            @RequestParam(required = false) @DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME) LocalDateTime fechaFin,
            @RequestParam(required = false) Integer productoId,
            Model model) {
        
        try {
            List<TransferenciaUbicacion> transferencias;
            
            if (fechaInicio != null && fechaFin != null) {
                // Buscar por rango de fechas
                transferencias = obtenerTransferenciasPorRangoFechas(fechaInicio, fechaFin);
                System.out.println("🔍 Búsqueda por fechas: " + transferencias.size() + " resultados");
            } else if (productoId != null) {
                // Buscar por producto
                transferencias = transferenciaService.obtenerPorProducto(productoId);
                System.out.println("🔍 Búsqueda por producto: " + transferencias.size() + " resultados");
            } else {
                // Obtener todas las transferencias
                transferencias = transferenciaService.obtenerUltimasTransferencias();
                System.out.println("🔍 Todas las transferencias: " + transferencias.size());
            }
            
            model.addAttribute("transferencias", transferencias != null ? transferencias : Collections.emptyList());
            model.addAttribute("productos", productoService.obtenerTodosProductos());
            model.addAttribute("ubicaciones", ubicacionService.obtenerTodasUbicaciones());
            
            if (productoId != null) {
                model.addAttribute("productoSeleccionado", productoService.obtenerProductoPorId(productoId).orElse(null));
            }
        } catch (Exception e) {
            System.err.println("❌ Error en búsqueda de transferencias: " + e.getMessage());
            e.printStackTrace();
            model.addAttribute("error", "Error en la búsqueda: " + e.getMessage());
            model.addAttribute("transferencias", Collections.emptyList());
            model.addAttribute("productos", Collections.emptyList());
            model.addAttribute("ubicaciones", Collections.emptyList());
        }
        
        return "transferencias/lista";
    }
}