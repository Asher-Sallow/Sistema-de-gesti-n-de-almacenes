package com.salesiana.inventory_system.repository;

import com.salesiana.inventory_system.entity.TransferenciaUbicacion;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;
import java.time.LocalDateTime;
import java.util.List;

@Repository
public interface TransferenciaUbicacionRepository extends JpaRepository<TransferenciaUbicacion, Integer> {
    List<TransferenciaUbicacion> findByProductoIdOrderByFechaTransferenciaDesc(Integer productoId);
    
    List<TransferenciaUbicacion> findByUbicacionOrigenIdOrUbicacionDestinoIdOrderByFechaTransferenciaDesc(
            Integer origenId, Integer destinoId);
    
    List<TransferenciaUbicacion> findByFechaTransferenciaBetween(LocalDateTime inicio, LocalDateTime fin);
    
    @Query("SELECT t FROM TransferenciaUbicacion t WHERE t.lote.id = ?1 ORDER BY t.fechaTransferencia DESC")
    List<TransferenciaUbicacion> findByLoteId(Integer loteId);
    
    @Query("SELECT t FROM TransferenciaUbicacion t ORDER BY t.fechaTransferencia DESC LIMIT 10")
    List<TransferenciaUbicacion> findUltimasTransferencias();
}