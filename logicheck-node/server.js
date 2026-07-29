const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Memoria del servidor para las órdenes y estados
let baseDatosOrdenes = {};

// 1. Endpoint para que el supervisor suba el archivo de Profit
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
                        estado: "PENDIENTE", // PENDIENTE, EN_PROCESO, COMPLETADO
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

// 2. Endpoint para obtener el estado actual de las órdenes
app.get('/api/ordenes', (req, res) => {
    res.json(baseDatosOrdenes);
});

// 3. Endpoint para que un operador tome una orden
app.post('/api/tomar-orden', (req, res) => {
    const { key, operador } = req.body;
    if (baseDatosOrdenes[key] && baseDatosOrdenes[key].estado === "PENDIENTE") {
        baseDatosOrdenes[key].estado = "EN_PROCESO";
        baseDatosOrdenes[key].operadorAsignado = operador;
        return res.json({ success: true });
    }
    res.status(400).json({ error: 'La orden ya no está disponible' });
});

// 4. Endpoint para actualizar el escaneo de un producto
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

// 5. Endpoint para liberar o finalizar una orden
app.post('/api/finalizar-orden', (req, res) => {
    const { key, estado } = req.body; // estado: "PENDIENTE" (si cancela) o "COMPLETADO"
    if (baseDatosOrdenes[key]) {
        baseDatosOrdenes[key].estado = estado;
        if (estado === "PENDIENTE") baseDatosOrdenes[key].operadorAsignado = null;
        return res.json({ success: true });
    }
    res.status(404).json({ error: 'Orden no encontrada' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor LogiCheck corriendo en puerto ${PORT}`);
});