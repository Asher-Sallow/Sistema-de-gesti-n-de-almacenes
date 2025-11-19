package com.salesiana.inventory_system.controller;

import com.salesiana.inventory_system.entity.Producto;
import com.salesiana.inventory_system.service.ProductoService;
import com.salesiana.inventory_system.service.CategoriaService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Optional;

@Controller
@RequestMapping("/productos")
public class ProductoController {
    
    @Autowired
    private ProductoService productoService;
    
    @Autowired
    private CategoriaService categoriaService;
    
    @GetMapping
    public String listarProductos(@RequestParam(required = false) String stock, Model model) {
        try {
            List<Producto> productos;
            
            if (stock != null) {
                switch (stock) {
                    case "bajo":
                        productos = productoService.obtenerProductosStockBajo();
                        break;
                    case "critico":
                        productos = productoService.obtenerProductosStockBajo();
                        break;
                    case "agotado":
                        productos = productoService.obtenerProductosAgotados();
                        break;
                    default:
                        productos = productoService.obtenerTodosProductos();
                }
            } else {
                productos = productoService.obtenerTodosProductos();
            }
            
            model.addAttribute("productos", productos);
            model.addAttribute("productosStockBajo", productoService.obtenerProductosStockBajo().size());
            model.addAttribute("productosAgotados", productoService.obtenerProductosAgotados().size());
            
            System.out.println("Productos cargados: " + productos.size());
            
        } catch (Exception e) {
            System.err.println("Error al cargar productos: " + e.getMessage());
            e.printStackTrace();
            model.addAttribute("error", "Error al cargar los productos: " + e.getMessage());
        }
        
        return "productos/lista";
    }
    
    @GetMapping("/nuevo")
    public String mostrarFormularioNuevo(Model model) {
        model.addAttribute("producto", new Producto());
        model.addAttribute("categorias", categoriaService.obtenerTodasCategorias());
        return "productos/form";
    }
    
    @PostMapping("/guardar")
    public String guardarProducto(@ModelAttribute Producto producto) {
        try {
            productoService.guardarProducto(producto);
            return "redirect:/productos?success=Producto guardado correctamente";
        } catch (Exception e) {
            System.err.println("Error al guardar producto: " + e.getMessage());
            return "redirect:/productos?error=Error al guardar producto: " + e.getMessage();
        }
    }
    
    @GetMapping("/editar/{id}")
    public String mostrarFormularioEditar(@PathVariable Integer id, Model model) {
        try {
            Optional<Producto> producto = productoService.obtenerProductoPorId(id);
            if (producto.isPresent()) {
                model.addAttribute("producto", producto.get());
                model.addAttribute("categorias", categoriaService.obtenerTodasCategorias());
                return "productos/form";
            } else {
                return "redirect:/productos?error=Producto no encontrado";
            }
        } catch (Exception e) {
            System.err.println("Error al cargar producto para editar: " + e.getMessage());
            return "redirect:/productos?error=Error al cargar producto";
        }
    }
    
    @GetMapping("/eliminar/{id}")
    public String eliminarProducto(@PathVariable Integer id) {
        try {
            productoService.eliminarProducto(id);
            return "redirect:/productos?success=Producto eliminado correctamente";
        } catch (Exception e) {
            System.err.println("Error al eliminar producto: " + e.getMessage());
            return "redirect:/productos?error=Error al eliminar producto: " + e.getMessage();
        }
    }
    
    @GetMapping("/buscar")
    public String buscarProductos(@RequestParam String q, Model model) {
        try {
            List<Producto> productos = productoService.buscarProductos(q);
            model.addAttribute("productos", productos);
            model.addAttribute("terminoBusqueda", q);
            model.addAttribute("productosStockBajo", productoService.obtenerProductosStockBajo().size());
            model.addAttribute("productosAgotados", productoService.obtenerProductosAgotados().size());
            
            System.out.println("Búsqueda: '" + q + "' - Resultados: " + productos.size());
            
        } catch (Exception e) {
            System.err.println("Error en búsqueda: " + e.getMessage());
            model.addAttribute("error", "Error en la búsqueda: " + e.getMessage());
        }
        
        return "productos/lista";
    }
}