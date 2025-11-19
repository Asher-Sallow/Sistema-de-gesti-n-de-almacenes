package com.salesiana.inventory_system.service;

import com.salesiana.inventory_system.entity.*;
import com.salesiana.inventory_system.repository.*;
import com.salesiana.inventory_system.repository.UbicacionAlmacenRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;

@Service
@Transactional
public class TransferenciaUbicacionService {

    @Autowired
    private TransferenciaUbicacionRepository transferenciaRepository;
    
    @Autowired
    private UbicacionAlmacenRepository ubicacionRepository;
    
    @Autowired
    private ProductoRepository productoRepository;
    
    @Autowired
    private LoteRepository loteRepository;
    
    @Autowired
    private UsuarioRepository usuarioRepository;

    public TransferenciaUbicacion registrarTransferencia(TransferenciaUbicacion transferencia) {
        // Validar capacidad destino
        UbicacionAlmacen destino = ubicacionRepository.findById(transferencia.getUbicacionDestino().getId())
                .orElseThrow(() -> new RuntimeException("Ubicación destino no encontrada"));
        
        if (!destino.tieneCapacidadDisponible(transferencia.getCantidad())) {
            throw new RuntimeException("La ubicación destino no tiene capacidad suficiente. Capacidad disponible: " + 
                    (destino.getCapacidadMaxima() - destino.getCapacidadActual()));
        }
        
        // Asignar usuario actual
        transferencia.setUsuario(obtenerUsuarioActual());
        
        // Actualizar stock del producto
        Producto producto = productoRepository.findById(transferencia.getProducto().getId())
                .orElseThrow(() -> new RuntimeException("Producto no encontrado"));
        
        // Si hay ubicación origen, validar y actualizar
        if (transferencia.getUbicacionOrigen() != null && transferencia.getUbicacionOrigen().getId() != null) {
            Integer origenId = transferencia.getUbicacionOrigen().getId();
            int stockDisponible = producto.getStockActual();
            
            // Verificar que haya suficiente stock en la ubicación origen
            if (stockDisponible < transferencia.getCantidad()) {
                throw new RuntimeException("Stock insuficiente en la ubicación de origen. Stock disponible: " + stockDisponible);
            }
            
            // Actualizar capacidad de la ubicación origen
            UbicacionAlmacen origen = ubicacionRepository.findById(origenId)
                    .orElseThrow(() -> new RuntimeException("Ubicación origen no encontrada"));
            origen.setCapacidadActual(origen.getCapacidadActual() - transferencia.getCantidad());
            ubicacionRepository.save(origen);
        }
        
        // Actualizar capacidad de la ubicación destino
        destino.setCapacidadActual(destino.getCapacidadActual() + transferencia.getCantidad());
        ubicacionRepository.save(destino);
        
        // Actualizar ubicación del lote si existe
        if (transferencia.getLote() != null && transferencia.getLote().getId() != null) {
            Optional<Lote> loteOpt = loteRepository.findById(transferencia.getLote().getId());
            loteOpt.ifPresent(lote -> {
                lote.setUbicacion(destino);
                loteRepository.save(lote);
            });
        }
        
        // Guardar la transferencia
        return transferenciaRepository.save(transferencia);
    }

    public List<TransferenciaUbicacion> obtenerPorProducto(Integer productoId) {
        return transferenciaRepository.findByProductoIdOrderByFechaTransferenciaDesc(productoId);
    }

    public List<TransferenciaUbicacion> obtenerPorLote(Integer loteId) {
        return transferenciaRepository.findByLoteId(loteId);
    }

    public List<TransferenciaUbicacion> obtenerPorUbicacion(Integer ubicacionId) {
        return transferenciaRepository.findByUbicacionOrigenIdOrUbicacionDestinoIdOrderByFechaTransferenciaDesc(
                ubicacionId, ubicacionId);
    }
    
    public List<TransferenciaUbicacion> obtenerUltimasTransferencias() {
        return transferenciaRepository.findUltimasTransferencias();
    }

    private Usuario obtenerUsuarioActual() {
        Authentication auth = SecurityContextHolder.getContext().getAuthentication();
        if (auth != null && auth.isAuthenticated() && !"anonymousUser".equals(auth.getName())) {
            return usuarioRepository.findByUsername(auth.getName()).orElse(null);
        }
        return null;
    }
}