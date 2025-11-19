/* 
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Other/javascript.js to edit this template
 */

// Funciones generales de la aplicación
document.addEventListener('DOMContentLoaded', function() {
    // Inicializar tooltips de Bootstrap
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });

    // Confirmación para eliminar
    const deleteButtons = document.querySelectorAll('.btn-delete');
    deleteButtons.forEach(button => {
        button.addEventListener('click', function(e) {
            if (!confirm('¿Está seguro de que desea eliminar este registro?')) {
                e.preventDefault();
            }
        });
    });

    // Auto-ocultar alertas después de 5 segundos
    const alerts = document.querySelectorAll('.alert.auto-dismiss');
    alerts.forEach(alert => {
        setTimeout(() => {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        }, 5000);
    });

    // Validación de formularios
    const forms = document.querySelectorAll('form.needs-validation');
    forms.forEach(form => {
        form.addEventListener('submit', function(event) {
            if (!form.checkValidity()) {
                event.preventDefault();
                event.stopPropagation();
            }
            form.classList.add('was-validated');
        });
    });

    // Actualizar contadores en tiempo real
    actualizarContadores();
    setInterval(actualizarContadores, 60000); // Actualizar cada minuto
});

// Función para actualizar contadores
function actualizarContadores() {
    // Aquí podrías hacer llamadas AJAX para actualizar contadores en tiempo real
    console.log('Actualizando contadores...');
}

// Función para formatear números como moneda
function formatCurrency(amount) {
    return new Intl.NumberFormat('es-PE', {
        style: 'currency',
        currency: 'PEN'
    }).format(amount);
}

// Función para formatear fechas
function formatDate(dateString) {
    const date = new Date(dateString);
    return date.toLocaleDateString('es-PE', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
    });
}

// Función para formatear fecha y hora
function formatDateTime(dateTimeString) {
    const date = new Date(dateTimeString);
    return date.toLocaleString('es-PE', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    });
}

// Función para mostrar loading en botones
function showLoading(button) {
    const originalText = button.innerHTML;
    button.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Procesando...';
    button.disabled = true;
    
    return function() {
        button.innerHTML = originalText;
        button.disabled = false;
    };
}

// Función para buscar productos en tiempo real
function buscarProductos(query) {
    if (query.length < 2) return;
    
    fetch(`/inventory/api/productos/buscar?q=${encodeURIComponent(query)}`)
        .then(response => response.json())
        .then(data => {
            mostrarResultadosBusqueda(data);
        })
        .catch(error => console.error('Error en búsqueda:', error));
}

// Función para mostrar resultados de búsqueda
function mostrarResultadosBusqueda(productos) {
    const resultados = document.getElementById('resultados-busqueda');
    if (resultados) {
        resultados.innerHTML = productos.map(producto => `
            <div class="list-group-item">
                <div class="d-flex w-100 justify-content-between">
                    <h6 class="mb-1">${producto.nombre}</h6>
                    <small>Stock: ${producto.stockActual}</small>
                </div>
                <p class="mb-1">${producto.codigo} - ${producto.categoria}</p>
                <small>Precio: ${formatCurrency(producto.precioVenta)}</small>
            </div>
        `).join('');
    }
}

// Exportar funciones para uso global
window.app = {
    formatCurrency,
    formatDate,
    formatDateTime,
    showLoading,
    buscarProductos
};
