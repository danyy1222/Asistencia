const express = require("express");
const QRCode = require("qrcode");
const XLSX = require("xlsx");

const app = express();
app.set("trust proxy", true);
app.use(express.urlencoded({ extended: true }));
app.set("view engine", "ejs");

const TIMEZONE = process.env.APP_TZ || "America/Lima";
const BASE_URL = process.env.BASE_URL || "";
let asistencias = [];

function generarID() {
  return Math.random().toString(36).substring(2, 8).toUpperCase();
}

function getBaseUrl(req) {
  if (BASE_URL) return BASE_URL;
  const host = req.get("host");
  const isVercel = host && host.includes("vercel.app");
  const forceHttps = process.env.FORCE_HTTPS === "1" || isVercel;
  const proto = forceHttps ? "https" : req.protocol;
  return `${proto}://${host}`;
}

function dateKey(ts = Date.now()) {
  const fmt = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  });
  return fmt.format(new Date(ts)); // YYYY-MM-DD
}

function listBySesion(sesion) {
  return asistencias.filter(a => a.sesion === sesion);
}

function listByDia(dia) {
  return asistencias.filter(a => a.dia === dia);
}

app.get("/", (req,res)=> res.redirect("/admin"));

app.get("/admin", async (req, res) => {
  const id = generarID();

  const baseUrl = getBaseUrl(req);
  const url = `${baseUrl}/asistencia/${id}`;

  const qr = await QRCode.toDataURL(url);

  res.render("admin", { qr, url, id, fecha: dateKey() });
});

app.get("/asistencia/:id", (req, res) => {
  res.render("asistencia", { sesion: req.params.id });
});

app.post("/registrar", async (req, res) => {
  const { nombre, sesion, ip } = req.body;
  const ahora = Date.now();

  const dia = dateKey(ahora);
  const listaSesion = listBySesion(sesion);
  const listaDia = listByDia(dia);

  const existe = listaSesion.find(a => a.nombre === nombre && a.sesion === sesion);
  if (existe) return res.send("Ya registrado");

  const ipReciente = listaDia.find(a => a.ip === ip && (ahora - a.hora) < 5*60*1000);
  if (ipReciente) return res.send("Espera unos minutos");

  const registro = { nombre, sesion, ip, hora: ahora, dia };
  asistencias.push(registro);

  res.send("OK registrado");
});

// EXPORTAR EXCEL
app.get("/excel/:sesion", async (req, res) => {
  const lista = listBySesion(req.params.sesion);

  const data = lista.map(a => ({
    Nombre: a.nombre,
    Hora: new Date(a.hora).toLocaleString(),
    IP: a.ip
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Asistencia");

  const file = `asistencia_${req.params.sesion}.xlsx`;
  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Disposition", `attachment; filename="${file}"`);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
});

// EXPORTAR EXCEL DEL DIA
app.get("/excel-dia", async (req, res) => {
  const dia = req.query.fecha || dateKey();
  const lista = listByDia(dia);

  const data = lista.map(a => ({
    Nombre: a.nombre,
    Sesion: a.sesion,
    Hora: new Date(a.hora).toLocaleString(),
    IP: a.ip,
    Dia: a.dia
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Asistencia");

  const file = `asistencia_${dia}.xlsx`;
  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Disposition", `attachment; filename="${file}"`);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
});
app.get("/ver/:sesion", async (req, res) => {
  const lista = listBySesion(req.params.sesion);
  res.render("lista", { lista, sesion: req.params.sesion });
});

if (require.main === module) {
  app.listen(3000, "0.0.0.0", ()=> console.log("Servidor listo"));
}

module.exports = app;
