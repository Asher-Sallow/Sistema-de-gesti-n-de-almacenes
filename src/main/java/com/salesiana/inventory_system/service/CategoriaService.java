/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.salesiana.inventory_system.service;

/**
 *
 * @author Andrei
 */

import com.salesiana.inventory_system.entity.Categoria;
import com.salesiana.inventory_system.repository.CategoriaRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Optional;

@Service
public class CategoriaService {
    
    @Autowired
    private CategoriaRepository categoriaRepository;
    
    public List<Categoria> obtenerTodasCategorias() {
        return categoriaRepository.findByActivaTrue();
    }
    
    public Optional<Categoria> obtenerCategoriaPorId(Integer id) {
        return categoriaRepository.findById(id);
    }
    
    public Categoria guardarCategoria(Categoria categoria) {
        return categoriaRepository.save(categoria);
    }
    
    public void eliminarCategoria(Integer id) {
        categoriaRepository.findById(id).ifPresent(categoria -> {
            categoria.setActiva(false);
            categoriaRepository.save(categoria);
        });
    }
    
    public boolean existeCategoriaConNombre(String nombre) {
        return categoriaRepository.existsByNombre(nombre);
    }
}
