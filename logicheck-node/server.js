const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());

// Memoria del servidor para las órdenes
let baseDatosOrdenes = {};

// Ruta principal para mostrar la interfaz web directamente
app.get('/', (req, res) => {
    res.send(`
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>LogiCheck - Servidor Node.js</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; color: #333; }
            .container { max-width: 900px; margin: 0 auto; background: white; padding: 25px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
            h1 { color: #2c3e50; text-align: center; margin-bottom: 25px; }
            .section { margin-bottom: 20px; padding: 15px; border: 1px solid #ddd; border-radius: 6px; background: #fafafa; }
            h3 { margin-top: 0; color: #34495e; }
            input[type="text"], input[type="file"] { width: 100%; padding: 10px; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; margin-top: 5px; }
            button { padding: 10px 15px; font-size: 15px; background-color: #3498db; color: white; border: none; border-radius: 4px; cursor: pointer; }
            button:hover { background-color: #2980b9; }
            .operador-btn { background-color: #9b59b6; margin: 5px; }
            .operador-btn:hover { background-color: #8e44ad; }
            table { width: 100%; border-collapse: collapse; margin-top: 15px; }
            th, td { border: 1px solid #ddd; padding: 10px; text-align: center; }
            th { background-color: #2c3e50; color: white; }
            .success { background-color: #d4edda; color: #155724; padding: 15px; border-radius: 4px; margin-top: 15px; text-align: center; font-weight: bold; display: none; }
            .error { background-color: #f8d7da; color: #721c24; padding: 15px; border-radius: 4px; margin-top: 15px; text-align: center; font-weight: bold; display: none; }
            #btnFinalizar { background-color: #2ecc71; width: 100%; margin-top: 20px; display: none; font-size: 18px; }
            #btnFinalizar:hover { background-color: #27ae60; }
            .badge { background: #e67e22; color: white; padding: 3px 8px; border-radius: 12px; font-size: 12px; }
        </style>
    </head>
    <body>

    <div class="container">
        <h1>LogiCheck - Servidor Node.js</h1>

        <div class="section" style="border-left: 5px solid #e67e22;">
            <h3>Panel de Supervisor: Cargar Archivo de Profit</h3>
            <input type="file" id="archivoProfit" accept=".xlsx, .xls, .csv">
            <button onclick="subirArchivo()" style="margin-top: 10px;">Subir y Sincronizar</button>
        </div>

        <div class="section" style="border-left: 5px solid #3498db;">
            <h3>Área de Picking: Selecciona tu Usuario</h3>
            <div>
                <button class="operador-btn" onclick="seleccionarOperador('Carlos')">Operador: Carlos</button>
                <button class="operador-btn" onclick="seleccionarOperador('Mariana')">Operador: Mariana</button>
                <button class="operador-btn" onclick="seleccionarOperador('José')">Operador: José</button>
                <button class="operador-btn" onclick="seleccionarOperador('Ana')">Operador: Ana</button>
            </div>
            <div id="infoOperadorActivo" style="margin-top: 10px; font-weight: bold; color: #2980b9;"></div>
        </div>

        <div class="section" id="seccionTrabajo" style="display:none; border-left: 5px solid #2ecc71;">
            <h3 id="tituloAreaTrabajo">Órdenes Pendientes</h3>
            
            <div id="listaPedidosPendientes">
                <p>Selecciona una nota para empezar:</p>
                <div id="botonesPedidosContainer"></div>
            </div>

            <div id="interfazEscaneo" style="display:none; margin-top: 15px; border-top: 1px dashed #ccc; padding-top: 15px;">
                <h4 id="detallePedidoActivo" style="color: #27ae60;"></h4>
                <p>Escanea los productos uno a uno con la pistola:</p>
                <input type="text" id="txtEscanerItem" placeholder="Escanea código SKU..." onkeypress="manejarEscaneo(event)">
                
                <table>
                    <thead>
                        <tr>
                            <th>SKU</th>
                            <th>Descripción</th>
                            <th>Esperado</th>
                            <th>Escaneado</th>
                            <th>Estado</th>
                        </tr>
                    </thead>
                    <tbody id="tablaItems"></tbody>
                </table>

                <div id="mensajeAlerta" class="error"></div>
                <div id="mensajeExito" class="success">¡Verificación exitosa! Conteo exacto completado.</div>

                <button id="btnFinalizar" onclick="finalizarOrden()">Guardar y Cerrar Nota</button>
                <br>
                <button onclick="cancelarOrdenActual()" style="background-color: #95a5a6; margin-top: 10px;">Regresar a la lista</button>
            </div>
        </div>
    </div>

    <script>
        let operadorActual = null;
        let pedidoEnCursoKey = null;
        let ordenesGlobales = {};

        setInterval(obtenerOrdenes, 3000);

        async function obtenerOrdenes() {
            try {
                let res = await fetch('/api/ordenes');
                ordenesGlobales = await res.json();
                if (operadorActual) {
                    actualizarListaPedidos();
                    if (pedidoEnCursoKey && ordenesGlobales[pedidoEnCursoKey]) {
                        actualizarTablaEscaneo();
                        verificarEstadoCompletado();
                    }
                }
            } catch (e) { console.error(e); }
        }

        async function subirArchivo() {
            let input = document.getElementById('archivoProfit');
            if (!input.files[0]) return alert('Selecciona un archivo primero');

            let formData = new FormData();
            formData.append('archivo', input.files[0]);

            let res = await fetch('/api/cargar-excel', { method: 'POST', body: formData });
            let data = await res.json();
            if (data.success) {
                alert(\`¡Carga exitosa! Se procesaron \${data.total} notas.\`);
                obtenerOrdenes();
            } else {
                alert('Error: ' + data.error);
            }
        }

        function seleccionarOperador(nombre) {
            operadorActual = nombre;
            document.getElementById("infoOperadorActivo").innerText = \`Operador activo: \${nombre}\`;
            document.getElementById("seccionTrabajo").style.display = "block";
            document.getElementById("interfazEscaneo").style.display = "none";
            obtenerOrdenes();
        }

        function actualizarListaPedidos() {
            const container = document.getElementById("botonesPedidosContainer");
            container.innerHTML = "";
            let hayDisponibles = false;

            for (let key in ordenesGlobales) {
                let orden = ordenesGlobales[key];
                if (orden.estado === "PENDIENTE" || (orden.estado === "EN_PROCESO" && orden.operadorAsignado === operadorActual)) {
                    hayDisponibles = true;
                    let div = document.createElement("div");
                    div.style.cssText = "background: white; padding: 10px; margin-bottom: 8px; border: 1px solid #ddd; border-radius: 4px; display: flex; justify-content: space-between; align-items: center;";
                    let badge = orden.operadorAsignado === operadorActual ? '<span class="badge">Tu Orden Activa</span>' : '';
                    div.innerHTML = \`<div><strong>\${orden.titulo}</strong> (\${orden.items.length} productos) \${badge}</div><button onclick="iniciarPicking('\${key}')">\${orden.operadorAsignado === operadorActual ? 'Continuar' : 'Tomar Nota'}</button>\`;
                    container.appendChild(div);
                }
            }
            if (!hayDisponibles) container.innerHTML = "<p style='color: #7f8c8d;'>No hay notas disponibles.</p>";
        }

        async function iniciarPicking(key) {
            let res = await fetch('/api/tomar-orden', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ key, operador: operadorActual })
            });
            let data = await res.json();
            if (data.success) {
                pedidoEnCursoKey = key;
                document.getElementById("listaPedidosPendientes").style.display = "none";
                document.getElementById("interfazEscaneo").style.display = "block";
                document.getElementById("detallePedidoActivo").innerText = \`Armando: \${ordenesGlobales[key].titulo}\`;
                document.getElementById("txtEscanerItem").value = "";
                document.getElementById("txtEscanerItem").focus();
                obtenerOrdenes();
            } else {
                alert(data.error);
                obtenerOrdenes();
            }
        }

        async function cancelarOrdenActual() {
            if (pedidoEnCursoKey) {
                await fetch('/api/finalizar-orden', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ key: pedidoEnCursoKey, estado: "PENDIENTE" })
                });
            }
            pedidoEnCursoKey = null;
            document.getElementById("interfazEscaneo").style.display = "none";
            document.getElementById("listaPedidosPendientes").style.display = "block";
            obtenerOrdenes();
        }

        async function manejarEscaneo(event) {
            if (event.key === "Enter" && pedidoEnCursoKey) {
                let sku = document.getElementById("txtEscanerItem").value.trim();
                document.getElementById("txtEscanerItem").value = "";

                let res = await fetch('/api/escanear', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ key: pedidoEnCursoKey, sku })
                });
                let data = await res.json();
                if (data.success) {
                    ordenesGlobales[pedidoEnCursoKey].items = data.items;
                    ocultarMensajes();
                    actualizarTablaEscaneo();
                    verificarEstadoCompletado();
                } else {
                    mostrarAlerta(data.error);
                }
            }
        }

        function actualizarTablaEscaneo() {
            if (!pedidoEnCursoKey || !ordenesGlobales[pedidoEnCursoKey]) return;
            const tbody = document.getElementById("tablaItems");
            tbody.innerHTML = "";
            let orden = ordenesGlobales[pedidoEnCursoKey];

            orden.items.forEach(item => {
                let estadoTexto = "Pendiente", color = "orange";
                if (item.escaneado === item.esperado) { estadoTexto = "Completo"; color = "green"; }
                else if (item.escaneado > item.esperado) { estadoTexto = "¡Sobrante!"; color = "red"; }
                let tr = document.createElement("tr");
                tr.innerHTML = \`<td>\${item.sku}</td><td>\${item.nombre}</td><td><b>\${item.esperado}</b></td><td><b>\${item.escaneado}</b></td><td style="color:\${color}; font-weight:bold;">\${estadoTexto}</td>\`;
                tbody.appendChild(tr);
            });
        }

        function verificarEstadoCompletado() {
            if (!pedidoEnCursoKey || !ordenesGlobales[pedidoEnCursoKey]) return;
            let orden = ordenesGlobales[pedidoEnCursoKey];
            let todoExacto = true, haySobrantes = false;
            orden.items.forEach(item => {
                if (item.escaneado !== item.esperado) todoExacto = false;
                if (item.escaneado > item.esperado) haySobrantes = true;
            });
            document.getElementById("btnFinalizar").style.display = (todoExacto && !haySobrantes) ? "block" : "none";
            document.getElementById("mensajeExito").style.display = (todoExacto && !haySobrantes) ? "block" : "none";
        }

        async function finalizarOrden() {
            if (!pedidoEnCursoKey) return;
            await fetch('/api/finalizar-orden', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ key: pedidoEnCursoKey, estado: "COMPLETADO" })
            });
            alert(\`¡Nota empacada con éxito por \${operadorActual}!\`);
            pedidoEnCursoKey = null;
            document.getElementById("interfazEscaneo").style.display = "none";
            document.getElementById("listaPedidosPendientes").style.display = "block";
            obtenerOrdenes();
        }

        function mostrarAlerta(texto) {
            let alerta = document.getElementById("mensajeAlerta");
            alerta.innerText = texto;
            alerta.style.display = "block";
        }
        function ocultarMensajes() {
            document.getElementById("mensajeAlerta").style.display = "none";
            document.getElementById("mensajeExito").style.display = "none";
        }
    </script>
    </body>
    </html>
    `);
});

// Endpoints de la API
app.post('/api/cargar-excel', upload.single('archivo'), (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No se subió ningún archivo' });

        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (rows.length === 0) return res.status(400).json({ error: 'El archivo está vacío' });

        let nuevasOrdenes = {};
        let notaContador = 1;
        let itemsNotaActual = [];
        let clienteActual = "Cliente General";

        for (let i = 0; i < rows.length; i++) {
            let fila = rows[i];
            let colA = String(fila[0] || "").trim();
            let esFilaTotal = fila.some(celda => String(celda).toLowerCase().includes("total"));

            if (esFilaTotal) {
                for (let c = 2; c < fila.length; c++) {
                    let valorCelda = String(fila[c] || "").trim();
                    if (valorCelda.length > 3 && !valorCelda.toLowerCase().includes("total") && !valorCelda.includes("RECEPTOR") && !valorCelda.includes("REMITENTE")) {
                        clienteActual = valorCelda;
                        break;
                    }
                }

                if (itemsNotaActual.length > 0) {
                    let idKey = `nota_${notaContador}`;
                    nuevasOrdenes[idKey] = {
                        id: idKey,
                        titulo: `Nota #${notaContador} - ${clienteActual}`,
                        cliente: clienteActual,
                        estado: "PENDIENTE",
                        operadorAsignado: null,
                        items: itemsNotaActual
                    };
                    notaContador++;
                    itemsNotaActual = [];
                    clienteActual = "Cliente General";
                }
                continue;
            }

            if (colA !== "" && colA.toLowerCase() !== "código" && colA.toLowerCase() !== "codigo") {
                let colB = String(fila[1] || "").trim();
                let cantidadVal = parseFloat(fila[6]) || parseFloat(fila[fila.length - 1]) || 1;

                itemsNotaActual.push({
                    sku: colA,
                    nombre: colB || "Sin descripción",
                    esperado: isNaN(cantidadVal) ? 1 : cantidadVal,
                    escaneado: 0
                });
            }
        }

        if (itemsNotaActual.length > 0) {
            let idKey = `nota_${notaContador}`;
            nuevasOrdenes[idKey] = {
                id: idKey,
                titulo: `Nota #${notaContador} - ${clienteActual}`,
                cliente: clienteActual,
                estado: "PENDIENTE",
                operadorAsignado: null,
                items: itemsNotaActual
            };
        }

        baseDatosOrdenes = nuevasOrdenes;
        res.json({ success: true, total: Object.keys(baseDatosOrdenes).length });

    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/ordenes', (req, res) => {
    res.json(baseDatosOrdenes);
});

app.post('/api/tomar-orden', (req, res) => {
    const { key, operador } = req.body;
    if (baseDatosOrdenes[key] && baseDatosOrdenes[key].estado === "PENDIENTE") {
        baseDatosOrdenes[key].estado = "EN_PROCESO";
        baseDatosOrdenes[key].operadorAsignado = operador;
        return res.json({ success: true });
    }
    res.status(400).json({ error: 'La orden ya no está disponible' });
});

app.post('/api/escanear', (req, res) => {
    const { key, sku } = req.body;
    let orden = baseDatosOrdenes[key];
    if (!orden) return res.status(404).json({ error: 'Orden no encontrada' });

    let item = orden.items.find(i => i.sku.toLowerCase() === sku.toLowerCase());
    if (item) {
        if (item.escaneado < item.esperado) {
            item.escaneado++;
            return res.json({ success: true, items: orden.items });
        } else {
            return res.status(400).json({ error: `Sobrante detectado para ${item.nombre}` });
        }
    }
    res.status(400).json({ error: 'El código no pertenece a esta nota' });
});

app.post('/api/finalizar-orden', (req, res) => {
    const { key, estado } = req.body;
    if (baseDatosOrdenes[key]) {
        baseDatosOrdenes[key].estado = estado;
        if (estado === "PENDIENTE") baseDatosOrdenes[key].operadorAsignado = null;
        return res.json({ success: true });
    }
    app.status(404).json({ error: 'Orden no encontrada' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor LogiCheck corriendo en puerto ${PORT}`);
});
