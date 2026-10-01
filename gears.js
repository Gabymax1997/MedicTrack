(() => {
    "use strict";

    const CLAVE_FILAS = "registro-insumos-filas";
    const CLAVE_SESION = "medictrack-sesion";
    const CLAVE_USUARIOS = "medictrack-usuarios";
    const USUARIOS_INICIALES = [
        { id: 1, nombre: "Administrador", usuario: "admin", password: "admin123", rol: "admin" },
        { id: 2, nombre: "Usuario", usuario: "usuario", password: "usuario123", rol: "usuario" }
    ];
    let filas = JSON.parse(localStorage.getItem(CLAVE_FILAS) || "[]");
    let seriesExcel = new Set();
    let lector = null;
    let libreriasListas = false;

    function obtenerSesion() {
        try {
            return JSON.parse(localStorage.getItem(CLAVE_SESION) || "null");
        } catch {
            return null;
        }
    }

    function guardarSesion(usuario) {
        localStorage.setItem(CLAVE_SESION, JSON.stringify({
            usuario: usuario.usuario,
            nombre: usuario.nombre,
            rol: usuario.rol
        }));
    }

    function cerrarSesion() {
        localStorage.removeItem(CLAVE_SESION);
        mostrarLogin();
    }

    function obtenerUsuarios() {
        try {
            const guardados = JSON.parse(localStorage.getItem(CLAVE_USUARIOS) || "null");
            if (Array.isArray(guardados) && guardados.length) return guardados;
        } catch {
            // ignore
        }

        localStorage.setItem(CLAVE_USUARIOS, JSON.stringify(USUARIOS_INICIALES));
        return [...USUARIOS_INICIALES];
    }

    function autenticarUsuario(usuario, password) {
        const usuarios = obtenerUsuarios();
        const usuarioNormalizado = String(usuario || "").trim();
        const passwordNormalizado = String(password || "");

        return usuarios.find((item) => item.usuario.toLowerCase() === usuarioNormalizado.toLowerCase()
            && item.password === passwordNormalizado);
    }

    function guardarUsuarios(usuarios) {
        localStorage.setItem(CLAVE_USUARIOS, JSON.stringify(usuarios));
    }

    function crearUsuario({ nombre, usuario, password, rol }) {
        const usuarios = obtenerUsuarios();
        const nombreLimpio = String(nombre || "").trim();
        const usuarioLimpio = String(usuario || "").trim();
        const passwordLimpio = String(password || "").trim();
        const rolLimpio = String(rol || "usuario").trim();

        if (!nombreLimpio || !usuarioLimpio || !passwordLimpio) {
            throw new Error("Faltan datos del usuario.");
        }

        if (!['admin', 'usuario'].includes(rolLimpio)) {
            throw new Error("Rol inválido.");
        }

        if (usuarios.some(item => item.usuario.toLowerCase() === usuarioLimpio.toLowerCase())) {
            throw new Error("Ya existe un usuario con ese nombre.");
        }

        const nuevoUsuario = {
            id: Date.now(),
            nombre: nombreLimpio,
            usuario: usuarioLimpio,
            password: passwordLimpio,
            rol: rolLimpio
        };

        usuarios.push(nuevoUsuario);
        guardarUsuarios(usuarios);
        return nuevoUsuario;
    }

    function borrarUsuario(id) {
        const usuarios = obtenerUsuarios();
        const usuarioActual = obtenerSesion();
        const usuarioAEliminar = usuarios.find(item => Number(item.id) === Number(id));

        if (!usuarioAEliminar) return false;
        if (usuarioActual && usuarioActual.usuario === usuarioAEliminar.usuario) {
            alert("No puedes eliminarte a ti mismo.");
            return false;
        }

        const filtrados = usuarios.filter(item => Number(item.id) !== Number(id));
        guardarUsuarios(filtrados);
        return true;
    }

    function guardarFilas() {
        localStorage.setItem(CLAVE_FILAS, JSON.stringify(filas));
    }

    function normalizarTexto(valor) {
        return String(valor || "")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[^a-zA-Z0-9]/g, "")
            .toLowerCase();
    }

    function normalizarSerie(valor) {
        return String(valor || "").trim().toUpperCase();
    }

    function limpiarCodigoEscaneado(codigo) {
        return String(codigo || "")
            .replace(/^\](?:d2|Q3|C1)/i, "")
            .replace(/[\u001d\u001e]/g, "|")
            .trim();
    }

    function obtenerEstadoSerie(fila) {
        const serie = normalizarSerie(fila["N.º de serie"]);

        if (!serie) return "Sin número de serie";
        if (!seriesExcel.size) return "Sin comparar";
        return seriesExcel.has(serie) ? "Encontrado" : "No encontrado";
    }

    function compararExcel(file) {
        const lectorArchivo = new FileReader();

        lectorArchivo.onload = (event) => {
            const libro = XLSX.read(event.target.result, { type: "array" });
            const hoja = libro.Sheets[libro.SheetNames[0]];
            const datos = XLSX.utils.sheet_to_json(hoja, { defval: "" });

            if (!datos.length) {
                alert("El Excel no contiene registros.");
                return;
            }

            const nombreSerie = Object.keys(datos[0]).find((nombre) => {
                const encabezado = normalizarTexto(nombre);
                return ["ndeserie", "nrodeserie", "numerodeserie", "serialnumber", "serial"].includes(encabezado)
                    || (encabezado.includes("serie") && (encabezado.includes("n") || encabezado.includes("numero")));
            });

            if (!nombreSerie) {
                alert("No encontré una columna de número de serie en el Excel.");
                return;
            }

            seriesExcel = new Set(
                datos
                    .map(fila => normalizarSerie(fila[nombreSerie]))
                    .filter(Boolean)
            );
            pintar();
            alert(`Comparación cargada: ${seriesExcel.size} número(s) de serie.`);
        };

        lectorArchivo.readAsArrayBuffer(file);
    }

    const cargar = (src) => new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = src;
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });

    function leerGS1(texto) {
        const datos = {
            GTIN: "",
            "N.º de serie": "",
            "N.º de lote": "",
            "Vencimiento": ""
        };

        const ais = {
            "01": "GTIN",
            "10": "N.º de lote",
            "17": "Vencimiento",
            "21": "N.º de serie"
        };

        const valor = String(texto || "").replace(/^\](?:d2|Q3|C1)/i, "").trim();
        const aiFijo = { "01": 14, "17": 6 };
        const separador = /[\u001d\u001e|]/;

        if (/\(\d{2,4}\)/.test(valor)) {
            const regex = /\((\d{2,4})\)([\s\S]*?)(?=\(\d{2,4}\)|$)/g;
            let match;

            while ((match = regex.exec(valor)) !== null) {
                const ai = match[1];
                const dato = match[2].replace(/[\u001d\u001e|]+$/, "").trim();
                if (ais[ai] && dato) datos[ais[ai]] = dato;
            }
        } else {
            let i = 0;

            while (i < valor.length) {
                if (separador.test(valor[i])) {
                    i++;
                    continue;
                }

                const ai = valor.slice(i, i + 2);
                if (!ais[ai]) {
                    i++;
                    continue;
                }

                i += 2;
                let dato = "";

                if (aiFijo[ai]) {
                    dato = valor.slice(i, i + aiFijo[ai]);
                    i += aiFijo[ai];
                } else {
                    const siguienteSeparador = valor.slice(i).search(separador);
                    const fin = siguienteSeparador < 0 ? valor.length : i + siguienteSeparador;
                    dato = valor.slice(i, fin);
                    i = fin;
                }

                if (dato) datos[ais[ai]] = dato.trim();
            }
        }

        if (datos["Vencimiento"] && /^\d{6}$/.test(datos["Vencimiento"])) {
            const fecha = datos["Vencimiento"];
            datos["Vencimiento"] = `${fecha.slice(4, 6)}/${fecha.slice(2, 4)}/20${fecha.slice(0, 2)}`;
        } else if (datos["Vencimiento"] && /^\d{8}$/.test(datos["Vencimiento"])) {
            const fecha = datos["Vencimiento"];
            datos["Vencimiento"] = `${fecha.slice(0, 2)}/${fecha.slice(2, 4)}/${fecha.slice(4, 8)}`;
        }

        if (!Object.values(datos).some(Boolean)) {
            datos.GTIN = valor;
        }

        return datos;
    }

    function serieYaRegistrada(serie, seriesNuevas = new Set()) {
        const serieNormalizada = normalizarSerie(serie);
        if (!serieNormalizada) return false;

        return filas.some(fila => normalizarSerie(fila["N.º de serie"]) === serieNormalizada)
            || seriesNuevas.has(serieNormalizada);
    }

    function agregar(codigo) {
        codigo = limpiarCodigoEscaneado(codigo);
        if (!codigo) return;

        const registro = {
            ...leerGS1(codigo),
            "Fecha y hora": new Date().toLocaleString("es-AR")
        };

        if (serieYaRegistrada(registro["N.º de serie"])) {
            alert(`El número de serie ${registro["N.º de serie"]} ya está registrado.`);
            return;
        }

        filas.push(registro);
        guardarFilas();
        pintar();
    }

    function agregarVarias(codigos) {
        const seriesNuevas = new Set();
        const registros = codigos
            .map(codigo => codigo.trim())
            .filter(Boolean)
            .map(codigo => ({
                ...leerGS1(codigo),
                "Fecha y hora": new Date().toLocaleString("es-AR")
            }))
            .filter(registro => {
                const serie = normalizarSerie(registro["N.º de serie"]);
                if (serieYaRegistrada(serie, seriesNuevas)) return false;
                if (serie) seriesNuevas.add(serie);
                return true;
            });

        if (!registros.length) return;

        filas.push(...registros);
        guardarFilas();
        pintar();
    }

    function pintar() {
        const cuerpo = document.querySelector("#insumos tbody");

        cuerpo.innerHTML = filas.map((fila, indice) => {
            const valores = [
                indice + 1,
                fila.GTIN || "",
                fila["N.º de serie"] || "",
                fila["N.º de lote"] || "",
                fila["Vencimiento"] || "",
                fila["Fecha y hora"] || ""
            ];

            return `<tr>${valores.map(v => `<td>${String(v).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]))}</td>`).join("")}</tr>`;
        }).join("");

        document.querySelector("#cantidad").textContent = `${filas.length} registro(s)`;
    }

    function leerDesdeImagen(file) {
        if (!file) return;

        const img = new Image();

        img.onload = () => {
            const canvas = document.createElement("canvas");
            canvas.width = img.width;
            canvas.height = img.height;

            const ctx = canvas.getContext("2d", { willReadFrequently: true });
            ctx.drawImage(img, 0, 0);

            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const code = jsQR(imageData.data, canvas.width, canvas.height);

            if (!code) {
                alert("No encontré un QR en la imagen.");
                return;
            }

            agregar(code.data);
            alert("QR leído desde la foto.");
        };

        img.src = URL.createObjectURL(file);
    }

    function mostrarLogin() {
        document.body.innerHTML = `
            <style>
                :root {
                    --bg-primary: #07111f;
                    --bg-secondary: #0d1b2a;
                    --panel: rgba(15, 23, 42, 0.82);
                    --border: rgba(148, 163, 184, 0.2);
                    --text: #e5eefb;
                    --muted: #9fb3c8;
                    --primary: #3b82f6;
                    --primary-strong: #2563eb;
                    --shadow: 0 18px 45px rgba(15, 23, 42, 0.45);
                }

                * { box-sizing: border-box; }

                body {
                    margin: 0;
                    min-height: 100vh;
                    display: grid;
                    place-items: center;
                    font-family: "Segoe UI", Tahoma, sans-serif;
                    background:
                        radial-gradient(circle at top left, rgba(59, 130, 246, 0.28), transparent 30%),
                        linear-gradient(135deg, var(--bg-primary), var(--bg-secondary));
                    color: var(--text);
                }

                .login-box {
                    width: min(420px, calc(100% - 24px));
                    background: rgba(15, 23, 42, 0.8);
                    border: 1px solid var(--border);
                    border-radius: 20px;
                    box-shadow: var(--shadow);
                    padding: 28px 24px;
                }

                .login-box h1 {
                    margin: 0 0 8px;
                    font-size: clamp(2rem, 5vw, 2.5rem);
                    letter-spacing: -0.04em;
                }

                .login-box p {
                    margin: 0 0 20px;
                    color: var(--muted);
                }

                form {
                    display: grid;
                    gap: 16px;
                }

                label {
                    display: grid;
                    gap: 8px;
                    font-weight: 600;
                    color: #dfeafc;
                }

                input {
                    width: 100%;
                    border: 1px solid rgba(148, 163, 184, 0.2);
                    background: rgba(15, 23, 42, 0.6);
                    color: var(--text);
                    border-radius: 12px;
                    padding: 12px 14px;
                    font: inherit;
                }

                input::placeholder {
                    color: var(--muted);
                }

                button {
                    border: none;
                    border-radius: 12px;
                    background: linear-gradient(135deg, var(--primary), var(--primary-strong));
                    color: white;
                    font-weight: 700;
                    padding: 12px 18px;
                    cursor: pointer;
                    box-shadow: 0 12px 20px rgba(37, 99, 235, 0.22);
                }

                .login-meta {
                    margin-top: 12px;
                    color: var(--muted);
                    font-size: 0.82rem;
                    line-height: 1.5;
                }
            </style>

            <div class="login-box">
                <h1>MedicTrack</h1>
                <p>Acceso al sistema</p>

                <form id="login-form">
                    <label>
                        Usuario
                        <input id="usuario" type="text" value="admin" autocomplete="username" required>
                    </label>

                    <label>
                        Contraseña
                        <input id="password" type="password" placeholder="Ingrese su contraseña" autocomplete="current-password" required>
                    </label>

                    <button type="submit">Entrar</button>
                </form>

        `;

        document.querySelector("#login-form").addEventListener("submit", (event) => {
            event.preventDefault();

            const usuario = document.querySelector("#usuario").value.trim();
            const password = document.querySelector("#password").value;
            const usuarioValido = autenticarUsuario(usuario, password);

            if (usuarioValido) {
                guardarSesion(usuarioValido);
                iniciarInterfaz();
                return;
            }

            alert("Credenciales incorrectas.");
            document.querySelector("#password").value = "";
        });
    }

    function renderUsuariosAdmin() {
        const lista = document.querySelector("#lista-usuarios");
        if (!lista) return;

        const usuarios = obtenerUsuarios();
        lista.innerHTML = usuarios.map((usuario) => `
            <tr>
                <td>${usuario.nombre}</td>
                <td>${usuario.usuario}</td>
                <td>${usuario.rol}</td>
                <td>
                    ${usuario.usuario === "admin"
                        ? "<span class='admin-tag'>Admin</span>"
                        : `<button type="button" class="btn-eliminar-usuario" data-id="${usuario.id}">Eliminar</button>`}
                </td>
            </tr>
        `).join("");

        document.querySelectorAll(".btn-eliminar-usuario").forEach((boton) => {
            boton.addEventListener("click", () => {
                const id = Number(boton.dataset.id);
                if (!id) return;
                if (confirm("¿Desea eliminar este usuario?")) {
                    const ok = borrarUsuario(id);
                    if (ok) renderUsuariosAdmin();
                }
            });
        });
    }

    function iniciarInterfaz() {
        const sesionActual = obtenerSesion();
        const nombreUsuario = sesionActual?.nombre || "Usuario";
        const esAdmin = sesionActual?.rol === "admin";

        document.body.innerHTML = `
            <style>
                :root {
                    --bg-primary: #07111f;
                    --bg-secondary: #0d1b2a;
                    --panel: rgba(15, 23, 42, 0.82);
                    --panel-strong: #111827;
                    --border: rgba(148, 163, 184, 0.2);
                    --text: #e5eefb;
                    --muted: #9fb3c8;
                    --primary: #3b82f6;
                    --primary-strong: #2563eb;
                    --success: #22c55e;
                    --danger: #ef4444;
                    --warning: #f59e0b;
                    --shadow: 0 18px 45px rgba(15, 23, 42, 0.45);
                }

                * { box-sizing: border-box; }

                body {
                    margin: 0;
                    min-height: 100vh;
                    font-family: "Segoe UI", Tahoma, sans-serif;
                    background:
                        radial-gradient(circle at top left, rgba(59, 130, 246, 0.28), transparent 30%),
                        radial-gradient(circle at bottom right, rgba(16, 185, 129, 0.20), transparent 25%),
                        linear-gradient(135deg, var(--bg-primary), var(--bg-secondary));
                    color: var(--text);
                }

                .app-shell {
                    max-width: 1200px;
                    margin: 32px auto;
                    padding: 0 18px 32px;
                }

                .topbar {
                    background: rgba(15, 23, 42, 0.72);
                    backdrop-filter: blur(10px);
                    border: 1px solid var(--border);
                    border-radius: 22px;
                    box-shadow: var(--shadow);
                    padding: 24px 26px;
                    margin-bottom: 22px;
                }

                .topbar-header {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    gap: 12px;
                    flex-wrap: wrap;
                }

                .topbar h1 {
                    margin: 0 0 8px;
                    font-size: clamp(1.9rem, 3.2vw, 2.7rem);
                    letter-spacing: -0.04em;
                }

                .topbar p {
                    margin: 0;
                    color: var(--muted);
                    font-size: 1rem;
                }

                .logout-btn {
                    background: linear-gradient(135deg, #ef4444, #b91c1c);
                    box-shadow: 0 10px 18px rgba(239, 68, 68, 0.18);
                }

                .panel {
                    background: rgba(15, 23, 42, 0.78);
                    border: 1px solid var(--border);
                    border-radius: 22px;
                    box-shadow: var(--shadow);
                    padding: 20px;
                    margin-bottom: 20px;
                }

                .camera-wrap {
                    display: flex;
                    flex-wrap: wrap;
                    justify-content: center;
                    gap: 18px;
                    align-items: center;
                }

                #lector {
                    flex: 1 1 420px;
                    min-height: 280px;
                    border-radius: 18px;
                    border: 1px solid rgba(148, 163, 184, 0.3);
                    background: rgba(2, 6, 23, 0.85);
                    overflow: hidden;
                }

                #lector video,
                #lector canvas {
                    width: 100%;
                    max-height: 420px;
                    border-radius: 18px;
                    display: block;
                }

                .toolbar {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 10px;
                    margin-top: 18px;
                }

                button {
                    border: none;
                    border-radius: 12px;
                    background: linear-gradient(135deg, var(--primary), var(--primary-strong));
                    color: white;
                    font-weight: 600;
                    padding: 11px 18px;
                    cursor: pointer;
                    transition: transform 0.18s ease, opacity 0.18s ease, box-shadow 0.18s ease;
                    box-shadow: 0 10px 18px rgba(37, 99, 235, 0.25);
                }

                button:hover:not(:disabled) {
                    transform: translateY(-1px);
                }

                button:disabled {
                    opacity: 0.6;
                    cursor: not-allowed;
                    box-shadow: none;
                }

                .btn-secondary {
                    background: linear-gradient(135deg, #334155, #1f2937);
                    box-shadow: none;
                }

                .btn-success {
                    background: linear-gradient(135deg, #16a34a, #15803d);
                    box-shadow: 0 10px 18px rgba(22, 163, 74, 0.2);
                }

                .btn-warning {
                    background: linear-gradient(135deg, #f59e0b, #d97706);
                    box-shadow: 0 10px 18px rgba(245, 158, 11, 0.18);
                }

                .manual-controls {
                    display: grid;
                    grid-template-columns: minmax(0, 1fr) auto;
                    gap: 12px;
                    align-items: start;
                    margin-top: 20px;
                }

                textarea {
                    width: 100%;
                    min-height: 110px;
                    resize: vertical;
                    border-radius: 14px;
                    border: 1px solid rgba(148, 163, 184, 0.28);
                    background: rgba(15, 23, 42, 0.7);
                    color: var(--text);
                    padding: 14px 16px;
                    font: inherit;
                }

                textarea::placeholder {
                    color: var(--muted);
                }

                .stats-row {
                    display: flex;
                    flex-wrap: wrap;
                    align-items: center;
                    gap: 10px;
                    margin-top: 18px;
                }

                .count {
                    display: inline-flex;
                    align-items: center;
                    padding: 8px 12px;
                    border-radius: 999px;
                    background: rgba(59, 130, 246, 0.12);
                    border: 1px solid rgba(59, 130, 246, 0.25);
                    color: #dbeafe;
                    font-weight: 700;
                }

                .admin-layout {
                    display: grid;
                    grid-template-columns: 1.5fr 1fr;
                    gap: 18px;
                    align-items: start;
                }

                .admin-lista {
                    width: 100%;
                }

                .admin-lista table {
                    margin-top: 10px;
                }

                .admin-form {
                    display: grid;
                    gap: 10px;
                }

                .admin-form input,
                .admin-form select {
                    width: 100%;
                    border: 1px solid rgba(148, 163, 184, 0.2);
                    background: rgba(15, 23, 42, 0.6);
                    color: var(--text);
                    border-radius: 12px;
                    padding: 10px 12px;
                    font: inherit;
                }

                .admin-tag {
                    display: inline-block;
                    padding: 6px 10px;
                    border-radius: 999px;
                    background: rgba(59, 130, 246, 0.12);
                    color: #dbeafe;
                    font-weight: 700;
                    font-size: 0.75rem;
                }

                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 18px;
                    background: rgba(15, 23, 42, 0.7);
                    border-radius: 16px;
                    overflow: hidden;
                    border: 1px solid var(--border);
                }

                th, td {
                    padding: 12px 10px;
                    text-align: left;
                    border-bottom: 1px solid rgba(148, 163, 184, 0.16);
                    vertical-align: middle;
                }

                th {
                    background: rgba(30, 41, 59, 0.9);
                    color: #dfeafc;
                    font-size: 0.8rem;
                    text-transform: uppercase;
                    letter-spacing: 0.06em;
                }

                tbody tr:hover {
                    background: rgba(59, 130, 246, 0.06);
                }

                .fila-encontrada td {
                    background: rgba(34, 197, 94, 0.18);
                    color: #ecfdf5;
                }

                .badge {
                    display: inline-block;
                    border-radius: 999px;
                    padding: 6px 10px;
                    font-weight: 700;
                    font-size: 0.8rem;
                    min-width: 100px;
                    text-align: center;
                }

                .badge-found {
                    background: rgba(34, 197, 94, 0.12);
                    color: #bbf7d0;
                }

                .badge-missing {
                    background: rgba(239, 68, 68, 0.12);
                    color: #fecaca;
                }

                .badge-warning {
                    background: rgba(245, 158, 11, 0.12);
                    color: #fcd34d;
                }

                .badge-plain {
                    background: rgba(148, 163, 184, 0.12);
                    color: #dfeafc;
                }

                .borrar-fila {
                    background: linear-gradient(135deg, #ef4444, #b91c1c);
                    padding: 8px 12px;
                    font-size: 0.8rem;
                    box-shadow: none;
                }

                @media (max-width: 700px) {
                    .topbar, .panel {
                        padding: 18px 14px;
                    }

                    .manual-controls {
                        grid-template-columns: 1fr;
                    }

                    th, td {
                        padding: 10px 8px;
                        font-size: 0.82rem;
                    }
                }
            </style>

            <main class="app-shell">
                <header class="topbar">
                    <div class="topbar-header">
                        <h1>Registro de insumos médicos</h1>
                        <button id="logout-top" class="logout-btn" type="button">Cerrar sesión</button>
                    </div>
                    <p>Bienvenido, ${nombreUsuario}. Escanee un código QR GS1/UDI para obtener GTIN, número de serie y lote.</p>
                </header>

                <section class="panel">
                    <div class="camera-wrap">
                        <div id="lector"></div>
                    </div>

                    <div class="toolbar">
                        <button id="iniciar">Iniciar cámara</button>
                        <button id="detener" class="btn-secondary" disabled>Detener</button>
                        <button id="foto" class="btn-success">Sacar foto</button>
                    </div>

                    <div class="manual-controls">
                        <textarea id="manual" rows="4" placeholder="Pegue un código por línea"></textarea>
                        <button id="agregar" class="btn-success">Agregar</button>
                    </div>

                    <div class="stats-row">
                        <button id="exportar" class="btn-secondary">Exportar a Excel</button>
                        <button id="comparar" class="btn-secondary">Comparar con Excel</button>
                        <button id="limpiar" class="btn-warning">Limpiar</button>
                        <button id="cerrar-sesion" class="btn-secondary">Cerrar sesión</button>
                        <span id="cantidad" class="count">0 registro(s)</span>
                    </div>
                </section>

                <section class="panel">
                    <table id="insumos">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>GTIN</th>
                                <th>N.º de serie</th>
                                <th>N.º de lote</th>
                                <th>Vencimiento</th>
                                <th>Fecha y hora</th>
                                <th>Estado</th>
                                <th>Acción</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>
                </section>

                ${esAdmin ? `
                    <section class="panel">
                        <h2>Administración de usuarios</h2>
                        <div class="admin-layout">
                            <div class="admin-lista">
                                <table>
                                    <thead>
                                        <tr>
                                            <th>Nombre</th>
                                            <th>Usuario</th>
                                            <th>Rol</th>
                                            <th>Acción</th>
                                        </tr>
                                    </thead>
                                    <tbody id="lista-usuarios"></tbody>
                                </table>
                            </div>

                            <form id="form-nuevo-usuario" class="admin-form">
                                <h3>Crear usuario</h3>
                                <input type="text" id="nuevo-nombre" placeholder="Nombre completo" required>
                                <input type="text" id="nuevo-usuario" placeholder="Nombre de usuario" required>
                                <input type="password" id="nuevo-password" placeholder="Contraseña" required>
                                <select id="nuevo-rol">
                                    <option value="usuario">Usuario</option>
                                    <option value="admin">Admin</option>
                                </select>
                                <button type="submit">Crear usuario</button>
                            </form>
                        </div>
                    </section>
                ` : ""}
            </main>
        `;

        const inputFoto = document.createElement("input");
        inputFoto.type = "file";
        inputFoto.accept = "image/*";
        inputFoto.capture = "environment";
        inputFoto.style.display = "none";

        inputFoto.onchange = (event) => {
            const file = event.target.files[0];
            if (file) leerDesdeImagen(file);
            inputFoto.value = "";
        };

        document.body.appendChild(inputFoto);

        const inputExcel = document.createElement("input");
        inputExcel.type = "file";
        inputExcel.accept = ".xlsx,.xls,.csv";
        inputExcel.style.display = "none";
        inputExcel.onchange = (event) => {
            const file = event.target.files[0];
            if (file) compararExcel(file);
            inputExcel.value = "";
        };
        document.body.appendChild(inputExcel);

        document.querySelector("#foto").onclick = () => inputFoto.click();

        document.querySelector("#comparar").onclick = () => inputExcel.click();

        document.querySelector("#agregar").onclick = () => {
            const campo = document.querySelector("#manual");
            if (campo.value.trim()) {
                agregarVarias(campo.value.split(/\r?\n/));
                campo.value = "";
            }
        };

        document.querySelector("#limpiar").onclick = () => {
            if (!confirm("¿Está seguro de que desea borrar todos los registros?")) return;

            filas = [];
            seriesExcel = new Set();
            guardarFilas();
            pintar();
        };

        document.querySelector("#cerrar-sesion").onclick = () => {
            cerrarSesion();
        };

        const logoutTop = document.querySelector("#logout-top");
        if (logoutTop) {
            logoutTop.addEventListener("click", () => {
                cerrarSesion();
            });
        }

        if (esAdmin) {
            renderUsuariosAdmin();
            const formNuevoUsuario = document.querySelector("#form-nuevo-usuario");
            if (formNuevoUsuario) {
                formNuevoUsuario.addEventListener("submit", (event) => {
                    event.preventDefault();

                    const nombre = document.querySelector("#nuevo-nombre").value.trim();
                    const usuario = document.querySelector("#nuevo-usuario").value.trim();
                    const password = document.querySelector("#nuevo-password").value.trim();
                    const rol = document.querySelector("#nuevo-rol").value;

                    try {
                        crearUsuario({ nombre, usuario, password, rol });
                        formNuevoUsuario.reset();
                        renderUsuariosAdmin();
                        alert("Usuario creado correctamente.");
                    } catch (error) {
                        alert(error.message);
                    }
                });
            }
        }

        document.querySelector("#exportar").onclick = () => {
            const filasExportar = obtenerFilasMostradas().map((fila, indice) => ({
                "#": indice + 1,
                GTIN: fila.GTIN || "",
                "N.º de serie": fila["N.º de serie"] || "",
                "N.º de lote": fila["N.º de lote"] || "",
                Vencimiento: fila.Vencimiento || "",
                "Fecha y hora": fila["Fecha y hora"] || "",
                Estado: fila.soloExcel ? "Solo en Excel" : obtenerEstadoSerie(fila)
            }));

            if (!filasExportar.length) return alert("No hay registros.");

            const libro = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(libro, XLSX.utils.json_to_sheet(filasExportar), "Insumos");
            XLSX.writeFile(libro, "insumos-medicos.xlsx");
        };

        document.querySelector("#iniciar").onclick = async () => {
            if (!libreriasListas || typeof Html5Qrcode === "undefined") {
                alert("El lector todavía está cargando. Intente nuevamente en unos segundos.");
                return;
            }

            try {
                lector = new Html5Qrcode("lector");
                let lecturaProcesada = false;

                await lector.start(
                    { facingMode: "environment" },
                    {
                        fps: 10,
                        qrbox: { width: 280, height: 280 },
                        formatsToSupport: [
                            Html5QrcodeSupportedFormats.QR_CODE,
                            Html5QrcodeSupportedFormats.DATA_MATRIX,
                            Html5QrcodeSupportedFormats.CODE_128
                        ]
                    },
                    async (decodedText) => {
                        if (lecturaProcesada) return;

                        lecturaProcesada = true;
                        agregar(decodedText);
                        await lector.stop();
                        lector.clear();
                        lector = null;
                        document.querySelector("#iniciar").disabled = false;
                        document.querySelector("#detener").disabled = true;
                    }
                );

                document.querySelector("#iniciar").disabled = true;
                document.querySelector("#detener").disabled = false;
            } catch (error) {
                console.error(error);
                if (lector) {
                    try {
                        await lector.stop();
                        lector.clear();
                    } catch (errorLimpieza) {
                        console.error(errorLimpieza);
                    }
                    lector = null;
                }

                document.querySelector("#iniciar").disabled = false;
                document.querySelector("#detener").disabled = true;
                alert("No se pudo abrir la cámara. Abra la aplicación desde http://localhost:8000, permita el acceso y cierre otras aplicaciones que estén usando la cámara.");
            }
        };

        document.querySelector("#detener").onclick = async () => {
            if (lector) {
                await lector.stop();
                lector.clear();
                lector = null;
            }

            document.querySelector("#iniciar").disabled = false;
            document.querySelector("#detener").disabled = true;
        };
    }


function borrarFila(indice) {
    filas.splice(indice, 1);
    guardarFilas();
    pintar();
}

function pintar() {
    const cuerpo = document.querySelector("#insumos tbody");
    const filasMostradas = obtenerFilasMostradas();

    cuerpo.innerHTML = filasMostradas.map((fila, indice) => {
        const valores = [
            indice + 1,
            fila.GTIN || "",
            fila["N.º de serie"] || "",
            fila["N.º de lote"] || "",
            fila["Vencimiento"] || "",
            fila["Fecha y hora"] || ""
        ];

        const celdas = valores
            .map(v => `<td>${String(v).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]))}</td>`)
            .join("");

        const estado = fila.soloExcel ? "Solo en Excel" : obtenerEstadoSerie(fila);
        const estadoClass = estado === "Encontrado" ? "badge-found"
            : estado === "No encontrado" ? "badge-missing"
                : estado === "Solo en Excel" ? "badge-warning" : "badge-plain";
        const filaEncontrada = estado === "Encontrado" ? "fila-encontrada" : "";
        const accion = fila.soloExcel
            ? "-"
            : `<button type="button" class="borrar-fila" data-indice="${indice}">Borrar</button>`;

        return `<tr class="${filaEncontrada}">${celdas}<td><span class="badge ${estadoClass}">${estado}</span></td><td>${accion}</td></tr>`;
    }).join("");

    document.querySelector("#cantidad").textContent = `${filas.length} registro(s)`;

    document.querySelectorAll(".borrar-fila").forEach((btn) => {
        btn.addEventListener("click", () => {
            borrarFila(Number(btn.dataset.indice));
        });
    });
}

function obtenerFilasMostradas() {
    const seriesGuardadas = new Set(
        filas
            .map(fila => normalizarSerie(fila["N.º de serie"]))
            .filter(Boolean)
    );
    const filasSoloExcel = [...seriesExcel]
        .filter(serie => !seriesGuardadas.has(serie))
        .map(serie => ({
            "N.º de serie": serie,
            soloExcel: true
        }));
    return [...filas, ...filasSoloExcel];
}


    const sesionActual = obtenerSesion();

    if (!sesionActual || !["admin", "usuario"].includes(sesionActual.rol)) {
        mostrarLogin();
    } else {
        iniciarInterfaz();
        pintar();

        Promise.all([
            cargar("https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js"),
            cargar("https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"),
            cargar("https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js")
        ]).then(() => {
            libreriasListas = true;
        }).catch(() => {
            alert("No se pudieron cargar las librerías de escaneo, Excel y lectura de foto.");
        });
    }
})();