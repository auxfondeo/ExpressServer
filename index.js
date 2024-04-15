const express = require("express");
const path = require("path");
const { spawn } = require("child_process");
const app = express();

// Establecer la carpeta para servir archivos estáticos
app.use(express.static(path.join(__dirname, 'css')));
app.use(express.static(path.join(__dirname, 'img')));
//console.log("hokfa ññ")
const PORT = 8080;
const IP = "192.168.68.106"; // Cambia esta dirección IP según tu configuración
//const IP = "192.168.100.211";

app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "/html/login.html"));
});

app.get("/html/Menu.html", (req, res) => {
    res.sendFile(path.join(__dirname, "/html/Menu.html"));
});

app.get("/login.html", (req, res) => {
    res.sendFile(path.join(__dirname + "/login.html"));
});

app.post("/login", (req, res) => {
    const username = req.body.username;
    const password = req.body.password;

    // Aquí deberías agregar tu lógica de autenticación
    // Por ejemplo, puedes verificar las credenciales en una base de datos o en una lista predefinida de usuarios
    // Por ahora, simplemente redirigimos al usuario si las credenciales son correctas o mostramos un mensaje de error si no lo son.

    if ((username === 'alberto' && password === 'valora') ||
        (username === 'Vanessa' && password === 'valora') ||
        (username === 'Marco' && password === 'valora')  ||
        (username === 'Michelle' && password === 'valora')) {
        res.redirect('/html/Menu.html'); // Redirige a la página de menú si las credenciales son correctas
    } else {
        
        res.send('Credenciales incorrectas. Por favor, inténtalo de nuevo.'); // Muestra un mensaje de error si las credenciales son incorrectas
    }
});

app.get("/EstCuenta.html", (req, res) => {
    res.sendFile(path.join(__dirname + "/html/EstCuenta.html"));
});

app.post("/concatenar", (req, res) => {
    const Paterno = req.body.ApPat;
    const Materno = req.body.ApMat;
    const Nombres = req.body.Nombres;

    const nombreCompleto = Paterno.concat(" ", Materno, " ", Nombres);

    const pythonScript = path.join(__dirname, "python","EstCuentas","Main.py");

    const pythonProcess = spawn("python", [pythonScript, nombreCompleto],{ env: { PYTHONIOENCODING: 'utf-8' } });

    let rutaDelArchivoGenerado;

    pythonProcess.stdout.on("data", (data) => {
        const message_EstC = data.toString().trim();
        if (message_EstC.startsWith("RUTA_ARCHIVO: ")) {
            rutaDelArchivoGenerado = message_EstC.substring("RUTA_ARCHIVO:".length).trim();
            console.log("Terminado")
        } else {
            console.log(message_EstC);
        }
    });

    pythonProcess.on("exit", (code) => {
        if (rutaDelArchivoGenerado) {
            res.download(rutaDelArchivoGenerado);
        } else {
            res.status(500).send("Error: " + "El cliente " + nombreCompleto + " no fue hallado");
        }
    });
});

app.get("/Cartas.html",(req,res)=> {
    res.sendFile(path.join(__dirname + "/html/Cartas.html"));
});

app.post("/Cartas", (req,res) => {
    const Paterno = req.body.ApPat;
    const Materno = req.body.ApMat;
    const Nombres = req.body.Nombres;
    const Plazo = req.body.Lista;
    const fecha = req.body.fecha;
    const Numero = req.body.Numero;


    const nombreCompleto = Paterno.concat(" ", Materno, " ", Nombres);

    console.log("Nombre del cliente: ", nombreCompleto);
    console.log("Plazo seleccionado", Plazo);
    console.log("Fecha de Descuento", fecha);
    console.log("Numero:",Numero)

    const pythonScript = path.join(__dirname, "python","CartasReestructura","cartas.py");
    //const pythonScript = "C:/Users/alber/Documents/Archivos de Python/Cartera/cartas.py";

    const pythonProcess = spawn("python", [pythonScript, nombreCompleto,Plazo,fecha,Numero],{ env: { PYTHONIOENCODING: 'utf-8' } });

    let rutaDelArchivoGenerado;

    pythonProcess.stdout.on("data", (data) => {
        const message = data.toString().trim();
        if (message.startsWith("RUTA_ARCHIVO: ")) {
            rutaDelArchivoGenerado = message.substring("RUTA_ARCHIVO:".length).trim();
        } else {
            console.log(message);
        }
    });

    pythonProcess.on("exit", (code) => {
        if (rutaDelArchivoGenerado) {
            res.download(rutaDelArchivoGenerado);
        } else {
            res.status(500).send("Error: " + nombreCompleto + " no fue hallado");
        }
    });
});



app.listen(PORT, () => {
    //console.log("El servidor se está ejecutando en http://:" + PORT);
    console.log("El servidor se está ejecutando en http://" + IP + ":" + PORT + "/");
});
