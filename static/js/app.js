// ===================================================
// SST GESTOR PROFESIONAL - JavaScript Frontend
// ===================================================

// --- Estado global ---
let clienteSeleccionado = null;
let propuestaActual = null;
let listaClientes = [];

// --- DOM Elements ---
const form = document.getElementById('clienteForm');
const btnGenerar = document.getElementById('btnGenerarPropuesta');
const btnDescargar = document.getElementById('btnDescargarPdf');
const clienteSelect = document.getElementById('clienteSelect');
const resultadoDiv = document.getElementById('resultadoPropuesta');
const footerPdf = document.getElementById('footerPdf');
const alertContainer = document.getElementById('alertContainer');

// --- API URL (ajustar si es necesario) ---
const API_BASE = 'http://127.0.0.1:5000';

// ===================================================
// 1. CARGAR ACTIVIDADES ECONÓMICAS (CIIU)
// ===================================================
async function cargarActividades() {
    try {
        // Por ahora usamos datos estáticos (luego se pueden cargar desde la API)
        const actividades = [
            { codigo: 'A0111.11', descripcion: 'Cultivo de trigo.', nivel: 'Bajo' },
            { codigo: 'A0111.12', descripcion: 'Cultivo de maíz.', nivel: 'Bajo' },
            { codigo: 'B0510.00', descripcion: 'Extracción de carbón de piedra', nivel: 'Alto' },
            { codigo: 'B0610.00', descripcion: 'Extracción de petróleo', nivel: 'Alto' },
            { codigo: 'C1010.11', descripcion: 'Explotación de mataderos', nivel: 'Alto' },
            { codigo: 'C1610.01', descripcion: 'Aserrado de madera', nivel: 'Medio' },
            { codigo: 'F4100.10', descripcion: 'Construcción de edificios', nivel: 'Alto' },
            { codigo: 'G4510.01', descripcion: 'Venta de vehículos', nivel: 'Bajo' },
            { codigo: 'J5811.01', descripcion: 'Publicación de libros', nivel: 'Medio' },
        ];

        const select = document.getElementById('actividadCodigo');
        select.innerHTML = '<option value="">Seleccionar...</option>';
        actividades.forEach(a => {
            const opt = document.createElement('option');
            opt.value = a.codigo;
            opt.textContent = `${a.codigo} - ${a.descripcion.substring(0, 40)}... (${a.nivel})`;
            select.appendChild(opt);
        });
    } catch (error) {
        console.error('Error cargando actividades:', error);
    }
}

// ===================================================
// 2. CREAR CLIENTE
// ===================================================
form.addEventListener('submit', async (e) => {
    e.preventDefault();

    const data = {
        razon_social: document.getElementById('razonSocial').value.trim(),
        ruc: document.getElementById('ruc').value.trim(),
        representante: document.getElementById('representante').value.trim(),
        email: document.getElementById('email').value.trim(),
        telefono: document.getElementById('telefono').value.trim(),
        sector: document.getElementById('sector').value,
        actividad_codigo: document.getElementById('actividadCodigo').value,
        numero_trabajadores: parseInt(document.getElementById('numeroTrabajadores').value),
        tiene_grupos_prioritarios: document.getElementById('gruposPrioritarios').checked,
        tiene_responsable_previo: document.getElementById('responsablePrev').checked
    };

    // Validaciones básicas
    if (!data.razon_social || !data.ruc || !data.actividad_codigo || !data.numero_trabajadores) {
        mostrarAlerta('Por favor, complete todos los campos obligatorios.', 'warning');
        return;
    }

    if (data.ruc.length !== 13 || !/^\d+$/.test(data.ruc)) {
        mostrarAlerta('El RUC debe tener 13 dígitos numéricos.', 'warning');
        return;
    }

    mostrarLoading(true);

    try {
        const response = await fetch(`${API_BASE}/clientes`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });

        const result = await response.json();

        if (response.ok) {
            mostrarAlerta(`✅ Cliente "${data.razon_social}" creado con ID: ${result.id}`, 'success');
            form.reset();
            await cargarListaClientes(); // Recargar lista
        } else {
            mostrarAlerta(`❌ Error: ${result.error || 'No se pudo crear el cliente'}`, 'danger');
        }
    } catch (error) {
        mostrarAlerta('❌ Error de conexión con el servidor.', 'danger');
        console.error(error);
    }

    mostrarLoading(false);
});

// ===================================================
// 3. CARGAR LISTA DE CLIENTES
// ===================================================
async function cargarListaClientes() {
    try {
        // Como no tenemos endpoint GET /clientes, usamos un workaround:
        // Por ahora, cargamos desde el select que mantenemos manualmente.
        // En producción, añadir un endpoint GET /clientes
        // Por ahora, mantenemos la lista en el cliente.
        // Simulamos con datos locales
        if (listaClientes.length === 0) {
            // Si no hay datos, usamos algunos de ejemplo
            listaClientes = [
                { id: 1, razon_social: 'Constructora XYZ', ruc: '1790012345001' }
            ];
        }

        clienteSelect.innerHTML = '<option value="">Seleccionar cliente...</option>';
        listaClientes.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.id;
            opt.textContent = `${c.razon_social} (${c.ruc})`;
            clienteSelect.appendChild(opt);
        });

        // Si hay clientes, activamos el botón de generar
        btnGenerar.disabled = listaClientes.length === 0;

    } catch (error) {
        console.error('Error cargando clientes:', error);
    }
}

// ===================================================
// 4. GENERAR PROPUESTA
// ===================================================
btnGenerar.addEventListener('click', async () => {
    const clienteId = clienteSelect.value;
    if (!clienteId) {
        mostrarAlerta('Por favor, seleccione un cliente.', 'warning');
        return;
    }

    mostrarLoading(true);

    try {
        const response = await fetch(`${API_BASE}/propuestas/generar`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ cliente_id: parseInt(clienteId) })
        });

        const result = await response.json();

        if (response.ok) {
            propuestaActual = result;
            mostrarPropuesta(result);
            footerPdf.classList.remove('d-none');
            mostrarAlerta('✅ Propuesta generada exitosamente.', 'success');
        } else {
            mostrarAlerta(`❌ Error: ${result.error || 'No se pudo generar la propuesta'}`, 'danger');
        }
    } catch (error) {
        mostrarAlerta('❌ Error de conexión con el servidor.', 'danger');
        console.error(error);
    }

    mostrarLoading(false);
});

// ===================================================
// 5. MOSTRAR PROPUESTA EN PANTALLA
// ===================================================
function mostrarPropuesta(data) {
    const servicios = data.servicios || [];
    let html = `
        <div class="table-responsive">
            <table class="table table-striped table-bordered">
                <thead>
                    <tr>
                        <th>Descripción</th>
                        <th>Tipo</th>
                        <th>Cant.</th>
                        <th>Precio</th>
                        <th>Subtotal</th>
                    </tr>
                </thead>
                <tbody>
    `;

    servicios.forEach(s => {
        const tipoClass = s.tipo === 'Setup' ? 'badge bg-info' :
                          s.tipo === 'Mensual' ? 'badge bg-primary' : 'badge bg-warning';
        html += `
            <tr>
                <td>${s.descripcion}</td>
                <td><span class="${tipoClass}">${s.tipo}</span></td>
                <td>${s.cantidad}</td>
                <td>$${s.precio_unitario.toFixed(2)}</td>
                <td>$${s.subtotal.toFixed(2)}</td>
            </tr>
        `;
    });

    html += `
            <tr class="total-row">
                <td colspan="3"></td>
                <td><strong>Total Mensual:</strong></td>
                <td><strong>$${data.total_mensual.toFixed(2)}</strong></td>
            </tr>
            <tr class="total-row">
                <td colspan="3"></td>
                <td><strong>Total Anual:</strong></td>
                <td><strong>$${data.total_anual.toFixed(2)}</strong></td>
            </tr>
        </tbody>
        </table>
        </div>
        <div class="mt-3">
            <p class="text-muted small">
                <i class="fas fa-info-circle"></i> Propuesta basada en Decreto Ejecutivo 255 y Acuerdo Ministerial 196.
            </p>
            <p class="text-muted small">
                <i class="fas fa-file-pdf"></i> ID de propuesta: ${data.propuesta_id}
            </p>
        </div>
    `;

    resultadoDiv.innerHTML = html;

    // Guardar ID para descargar PDF
    propuestaActual = data;
}

// ===================================================
// 6. DESCARGAR PDF
// ===================================================
btnDescargar.addEventListener('click', () => {
    if (!propuestaActual || !propuestaActual.propuesta_id) {
        mostrarAlerta('No hay propuesta para descargar.', 'warning');
        return;
    }

    const url = `${API_BASE}/propuestas/${propuestaActual.propuesta_id}/pdf`;
    window.open(url, '_blank');
});

// ===================================================
// 7. UTILIDADES
// ===================================================

// --- Mostrar alerta ---
function mostrarAlerta(mensaje, tipo = 'info') {
    const colores = {
        success: 'success',
        danger: 'danger',
        warning: 'warning',
        info: 'info'
    };

    alertContainer.innerHTML = `
        <div class="alert alert-${colores[tipo] || 'info'} alert-dismissible fade show" role="alert">
            <i class="fas fa-${tipo === 'success' ? 'check-circle' : tipo === 'danger' ? 'exclamation-circle' : 'info-circle'} me-2"></i>
            ${mensaje}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;

    // Auto cerrar después de 5 segundos
    setTimeout(() => {
        const alert = alertContainer.querySelector('.alert');
        if (alert) alert.remove();
    }, 8000);
}

// --- Mostrar/ocultar loading ---
function mostrarLoading(show) {
    let overlay = document.querySelector('.spinner-overlay');
    if (!overlay) {
        overlay = document.createElement('div');
        overlay.className = 'spinner-overlay';
        overlay.innerHTML = `
            <div class="text-center">
                <div class="spinner-border text-primary" role="status" style="width: 4rem; height: 4rem;">
                    <span class="visually-hidden">Cargando...</span>
                </div>
                <p class="mt-3 fw-bold">Procesando...</p>
            </div>
        `;
        document.body.appendChild(overlay);
    }

    if (show) {
        overlay.classList.add('active');
    } else {
        overlay.classList.remove('active');
    }
}

// ===================================================
// 8. INICIALIZAR
// ===================================================
document.addEventListener('DOMContentLoaded', () => {
    cargarActividades();
    cargarListaClientes();

    // Agregar cliente de ejemplo si no hay (solo para pruebas)
    if (listaClientes.length === 0) {
        // Simular un cliente de ejemplo
        listaClientes = [
            { id: 1, razon_social: 'Constructora XYZ', ruc: '1790012345001' }
        ];
        cargarListaClientes();
    }
});

// ===================================================
// 9. BONUS: Guardar cliente con tecla Enter
// ===================================================
document.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.target.closest('#clienteForm')) {
        e.preventDefault();
        document.getElementById('btnCrearCliente').click();
    }
});