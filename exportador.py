"""
Exportador Tableau Server / Tableau Cloud
=========================================

Descarga a una carpeta local el contenido de uno o más sitios de Tableau:
  - Workbooks (.twbx / .twb)
  - Fuentes de datos publicadas (.tdsx / .tds)
  - Flujos de Prep (.tflx / .tfl)

Además genera:
  - inventario_<fecha>.xlsx : inventario completo del sitio (proyectos, workbooks, vistas,
                              fuentes, conexiones, flujos, usuarios, grupos, tareas,
                              suscripciones y campos por hoja como en el proyecto Copec)
  - pendientes_<fecha>.csv  : todo lo que NO se pudo descargar, con su URL y la acción sugerida
                              para hacerlo manualmente
  - logs/ejecucion_<fecha>.log

Uso:
    python exportador.py                    (usa config.ini)
    python exportador.py --config otro.ini
    python exportador.py --solo-inventario  (no descarga archivos)
"""

import argparse
import configparser
import csv
import getpass
import logging
import os
import re
import sys
import time
from datetime import date, datetime

import pandas as pd
import tableauserverclient as TSC
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# --------------------------- #
#        CONFIGURACIÓN        #
# --------------------------- #

EXTENSIONES = {
    "Workbook": [".twbx", ".twb"],
    "Fuente de datos": [".tdsx", ".tds"],
    "Flujo": [".tflx", ".tfl"],
}

CARPETAS = {
    "Workbook": "Workbooks",
    "Fuente de datos": "Fuentes de datos",
    "Flujo": "Flujos",
}

# Query de campos por hoja (misma estructura que el extractor de Copec, filtrando por LUID)
GRAPHQL_CAMPOS_WORKBOOK = """
query CamposPorWorkbook($luid: String) {
  workbooks(filter: {luid: $luid}) {
    name
    luid
    sheets {
      id
      name
      sheetFieldInstances {
        name
        __typename
        ... on ColumnField { dataType }
        ... on CalculatedField { dataType }
        ... on GroupField { dataType }
        ... on DatasourceField {
          remoteField {
            ... on ColumnField { dataType }
          }
        }
      }
    }
  }
}
"""

log = logging.getLogger("exportador")
pendientes = []      # lo que falta descargar / revisar manualmente


# ---------------------- #
#       UTILIDADES       #
# ---------------------- #

def leer_config(ruta):
    if not os.path.exists(ruta):
        sys.exit(f"No se encontró el archivo de configuración '{ruta}'")
    cfg = configparser.ConfigParser()
    cfg.read(ruta, encoding="utf-8")
    return cfg


def configurar_logs(carpeta, marca):
    os.makedirs(os.path.join(carpeta, "logs"), exist_ok=True)
    ruta_log = os.path.join(carpeta, "logs", f"ejecucion_{marca}.log")
    formato = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S")

    # La consola de Windows (cp1252) no soporta emojis/tildes en todos los casos
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    archivo = logging.FileHandler(ruta_log, encoding="utf-8")
    archivo.setFormatter(formato)
    consola = logging.StreamHandler(sys.stdout)
    consola.setFormatter(logging.Formatter("%(message)s"))

    log.setLevel(logging.INFO)
    log.addHandler(archivo)
    log.addHandler(consola)
    return ruta_log


def limpiar_nombre(nombre, largo_max=80):
    """Deja un nombre válido para carpeta/archivo en Windows."""
    nombre = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(nombre or "sin_nombre")).strip(" .")
    return (nombre or "sin_nombre")[:largo_max]


def ruta_windows(ruta):
    """Permite rutas de más de 260 caracteres en Windows."""
    ruta = os.path.abspath(ruta)
    if os.name == "nt" and not ruta.startswith("\\\\?\\"):
        return "\\\\?\\" + ruta
    return ruta


def fecha(valor):
    """Fechas sin zona horaria (Excel no las soporta)."""
    return valor.replace(tzinfo=None) if isinstance(valor, datetime) else valor


def listar(endpoint, **kwargs):
    """Lista todos los elementos de un endpoint paginado."""
    return list(TSC.Pager(endpoint, **kwargs))


def intentar(descripcion, sitio, funcion, por_defecto=None):
    """Ejecuta una consulta del inventario; si falla, la deja registrada y sigue."""
    try:
        return funcion()
    except Exception as e:
        log.warning(f"   ⚠ No se pudo obtener {descripcion}: {e}")
        registrar_pendiente(sitio, "Inventario", "", descripcion, "", "", e,
                            "Revisar permisos del token (se requiere rol Administrador de sitio/servidor).")
        return por_defecto


def accion_sugerida(error, tipo):
    texto = str(error).lower()
    if "403" in texto or "forbidden" in texto or "permission" in texto:
        return f"Sin permiso para descargar el {tipo.lower()}. Descargarlo con un administrador o dar permiso de descarga al usuario del token."
    if "404" in texto or "not found" in texto:
        return "El contenido no existe o no es visible para el usuario del token. Verificar en la web."
    if "timed out" in texto or "timeout" in texto or "connection" in texto:
        return "Error de red / tiempo de espera (archivo muy pesado). Re-ejecutar el exportador o descargar manualmente desde la URL."
    if "too long" in texto or "errno 2" in texto or "errno 22" in texto:
        return "Problema con la ruta local del archivo. Descargar manualmente desde la URL."
    return "Descargar manualmente desde la URL web."


def texto_error(error):
    """Mensaje de error en una sola línea (los de Tableau vienen con saltos de línea)."""
    return " ".join(str(error).split())


def registrar_pendiente(sitio, tipo, proyecto, nombre, item_id, url, error, accion=None):
    pendientes.append({
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sitio": sitio,
        "tipo": tipo,
        "proyecto": proyecto,
        "nombre": nombre,
        "id": item_id,
        "url_web": url,
        "error": texto_error(error)[:500],
        "accion_sugerida": accion or accion_sugerida(error, tipo),
    })


# ---------------------- #
#        CONEXIÓN        #
# ---------------------- #

class Conexion:
    """Maneja el login con PAT y la reconexión si la sesión expira durante descargas largas."""

    def __init__(self, url, token_name, token_value, verify_ssl):
        self.url = url.rstrip("/")
        self.token_name = token_name
        self.token_value = token_value
        self.server = TSC.Server(self.url, use_server_version=False)
        if not verify_ssl:
            import urllib3
            urllib3.disable_warnings()
            self.server.add_http_options({"verify": False})
        self.server.use_server_version()
        self.sitio = None

    def conectar(self, sitio):
        if self.server.is_signed_in():
            self.server.auth.sign_out()
        auth = TSC.PersonalAccessTokenAuth(self.token_name, self.token_value, site_id=sitio)
        self.server.auth.sign_in(auth)
        self.sitio = sitio

    def reconectar(self):
        log.info("   🔄 Renovando sesión...")
        self.conectar(self.sitio)

    def desconectar(self):
        if self.server.is_signed_in():
            self.server.auth.sign_out()

    def url_vista(self, content_url):
        # content_url de vista: "Workbook/sheets/Vista" -> #/views/Workbook/Vista
        ruta = (content_url or "").replace("/sheets/", "/")
        prefijo = f"site/{self.sitio}/" if self.sitio else ""
        return f"{self.url}/#/{prefijo}views/{ruta}"


# ---------------------- #
#       DESCARGAS        #
# ---------------------- #

def descargar(conexion, tipo, item, carpeta_tipo, ruta_proyecto, opciones):
    """Descarga un elemento. Retorna (estado, ruta_local)."""
    carpeta = os.path.join(carpeta_tipo, *[limpiar_nombre(p) for p in ruta_proyecto.split("/")])
    base = os.path.join(carpeta, limpiar_nombre(item.name))

    # Si ya existe (de una ejecución anterior) se salta
    for ext in EXTENSIONES[tipo]:
        if os.path.exists(ruta_windows(base + ext)) and not opciones["sobrescribir"]:
            return "Ya existía", base + ext

    os.makedirs(ruta_windows(carpeta), exist_ok=True)
    temporal = ruta_windows(base + "__descargando")
    server = conexion.server

    for intento in range(2):
        try:
            if tipo == "Workbook":
                ruta = server.workbooks.download(item.id, filepath=temporal, include_extract=opciones["incluir_extractos"])
            elif tipo == "Fuente de datos":
                ruta = server.datasources.download(item.id, filepath=temporal, include_extract=opciones["incluir_extractos"])
            else:
                ruta = server.flows.download(item.id, filepath=temporal)

            extension = os.path.splitext(ruta)[1]
            final = ruta_windows(base + extension)
            os.replace(ruta, final)
            return "Descargado", base + extension

        except TSC.ServerResponseError as e:
            borrar_temporales(temporal, tipo)
            if str(e.code).startswith("401") and intento == 0:
                conexion.reconectar()
                continue
            raise
        except Exception:
            borrar_temporales(temporal, tipo)
            if intento == 0:
                time.sleep(3)
                continue
            raise


def borrar_temporales(temporal, tipo):
    for ext in EXTENSIONES[tipo]:
        if os.path.exists(temporal + ext):
            os.remove(temporal + ext)


def descargar_lista(conexion, tipo, items, carpeta_sitio, rutas, opciones):
    """Descarga una lista de workbooks/fuentes/flujos y retorna {id: (estado, archivo)}."""
    resultados = {}
    carpeta_tipo = os.path.join(carpeta_sitio, CARPETAS[tipo])
    total = len(items)

    for i, item in enumerate(items, 1):
        proyecto = rutas.get(item.project_id, "Espacio personal")
        inicio = time.time()
        try:
            estado, archivo = descargar(conexion, tipo, item, carpeta_tipo, proyecto, opciones)
            tamano = os.path.getsize(ruta_windows(archivo)) / 1024 / 1024
            log.info(f"   [{i}/{total}] {tipo} '{proyecto}/{item.name}' -> {estado} "
                     f"({tamano:.1f} MB, {time.time() - inicio:.0f}s)")
        except Exception as e:
            estado, archivo = "Error", ""
            log.error(f"   [{i}/{total}] {tipo} '{proyecto}/{item.name}' -> ERROR: {texto_error(e)}")
            registrar_pendiente(conexion.sitio, tipo, proyecto, item.name, item.id,
                                getattr(item, "webpage_url", ""), e)
        resultados[item.id] = (estado, archivo)

    return resultados


# ---------------------- #
#       INVENTARIO       #
# ---------------------- #

def rutas_proyectos(proyectos):
    """Arma la ruta completa de cada proyecto: 'Padre/Hijo/Nieto'."""
    por_id = {p.id: p for p in proyectos}
    rutas = {}
    for p in proyectos:
        partes, actual = [], p
        while actual and len(partes) < 50:
            partes.insert(0, actual.name)
            actual = por_id.get(actual.parent_id)
        rutas[p.id] = "/".join(partes)
    return rutas


def filtrar_por_proyecto(items, rutas, filtro):
    if not filtro:
        return items
    return [i for i in items if rutas.get(i.project_id, "").split("/")[0] in filtro]


def campos_por_hoja(conexion, workbooks, rutas):
    """Inventario de campos usados en cada hoja (mismo formato que el proyecto Copec)."""
    registros = []
    errores_seguidos = 0
    total = len(workbooks)
    for i, wb in enumerate(workbooks, 1):
        log.info(f"   [{i}/{total}] Campos de '{wb.name}'")
        try:
            try:
                respuesta = conexion.server.metadata.query(GRAPHQL_CAMPOS_WORKBOOK, variables={"luid": wb.id})
            except TSC.ServerResponseError as e:
                if not str(e.code).startswith("401"):
                    raise
                conexion.reconectar()
                respuesta = conexion.server.metadata.query(GRAPHQL_CAMPOS_WORKBOOK, variables={"luid": wb.id})

            if respuesta.get("errors"):
                raise RuntimeError(respuesta["errors"])

            for wb_meta in respuesta.get("data", {}).get("workbooks", []):
                for hoja in wb_meta.get("sheets", []):
                    for campo in hoja.get("sheetFieldInstances", []):
                        tipo = campo.get("__typename")
                        if tipo == "DatasourceField":
                            data_type = (campo.get("remoteField") or {}).get("dataType")
                        else:
                            data_type = campo.get("dataType")
                        registros.append({
                            "site_name": conexion.sitio or "Default",
                            "project": rutas.get(wb.project_id, "Espacio personal"),
                            "workbook_name": wb_meta.get("name"),
                            "workbook_id": wb_meta.get("luid"),
                            "sheet_name": hoja.get("name"),
                            "sheet_id": hoja.get("id"),
                            "field_name": campo.get("name"),
                            "field_type": tipo,
                            "data_type": data_type,
                            "extraction_date": date.today(),
                        })
            errores_seguidos = 0
        except Exception as e:
            errores_seguidos += 1
            log.warning(f"   ⚠ Error leyendo campos de '{wb.name}': {e}")
            registrar_pendiente(conexion.sitio, "Inventario campos", rutas.get(wb.project_id, ""), wb.name, wb.id,
                                wb.webpage_url, e,
                                "Verificar que la Metadata API esté habilitada "
                                "(tsm maintenance metadata-services enable) y que el token vea el workbook.")
            if errores_seguidos >= 5 and not registros:
                log.warning("   ⚠ La Metadata API no responde; se omite el inventario de campos de este sitio.")
                break
    return registros


def inventario_sitio(conexion, opciones, carpeta_sitio, solo_inventario):
    """Lista todo el contenido del sitio, descarga archivos y retorna las tablas del inventario."""
    server = conexion.server
    sitio = conexion.sitio
    nombre_sitio = sitio or "Default"
    filtro = opciones["filtro_proyectos"]
    tablas = {}

    log.info("📂 Leyendo proyectos, usuarios y grupos...")
    proyectos = listar(server.projects)
    rutas = rutas_proyectos(proyectos)
    usuarios = intentar("usuarios", sitio, lambda: listar(server.users), [])
    nombre_usuario = {u.id: u.name for u in usuarios}

    tablas["Proyectos"] = [{
        "sitio": nombre_sitio, "id": p.id, "nombre": p.name, "ruta": rutas[p.id],
        "proyecto_padre": rutas.get(p.parent_id, ""), "descripcion": p.description,
        "modo_permisos": p.content_permissions, "propietario": nombre_usuario.get(getattr(p, "owner_id", None), ""),
    } for p in proyectos if not filtro or rutas[p.id].split("/")[0] in filtro]

    tablas["Usuarios"] = [{
        "sitio": nombre_sitio, "id": u.id, "usuario": u.name, "nombre_completo": u.fullname, "email": u.email,
        "rol_sitio": u.site_role, "autenticacion": u.auth_setting, "ultimo_login": fecha(u.last_login),
    } for u in usuarios]

    grupos = intentar("grupos", sitio, lambda: listar(server.groups), [])
    tablas["Grupos"], tablas["Usuarios por grupo"] = [], []
    for g in grupos:
        miembros = intentar(f"miembros del grupo {g.name}", sitio, lambda: _miembros_grupo(server, g), [])
        tablas["Grupos"].append({
            "sitio": nombre_sitio, "id": g.id, "nombre": g.name, "dominio": g.domain_name,
            "rol_minimo": g.minimum_site_role, "cantidad_usuarios": len(miembros),
        })
        tablas["Usuarios por grupo"] += [{"sitio": nombre_sitio, "grupo": g.name, "usuario": u.name} for u in miembros]

    # --- Workbooks, vistas y fuentes de datos ---
    log.info("📚 Leyendo workbooks, vistas, fuentes de datos y flujos...")
    workbooks = filtrar_por_proyecto(listar(server.workbooks), rutas, filtro)
    fuentes = filtrar_por_proyecto(listar(server.datasources), rutas, filtro)
    flujos = filtrar_por_proyecto(intentar("flujos", sitio, lambda: listar(server.flows), []), rutas, filtro)
    vistas = intentar("vistas con uso", sitio, lambda: listar(server.views, usage=True), None)
    if vistas is None:
        vistas = intentar("vistas", sitio, lambda: listar(server.views), [])
    ids_wb = {w.id: w for w in workbooks}
    vistas = [v for v in vistas if v.workbook_id in ids_wb]
    log.info(f"   {len(proyectos)} proyectos | {len(workbooks)} workbooks | {len(vistas)} vistas | "
             f"{len(fuentes)} fuentes | {len(flujos)} flujos | {len(usuarios)} usuarios")

    # --- Descargas ---
    descargas = {}
    if not solo_inventario:
        if opciones["descargar_workbooks"]:
            log.info(f"⬇  Descargando {len(workbooks)} workbooks...")
            descargas.update(descargar_lista(conexion, "Workbook", workbooks, carpeta_sitio, rutas, opciones))
        if opciones["descargar_fuentes"]:
            log.info(f"⬇  Descargando {len(fuentes)} fuentes de datos...")
            descargas.update(descargar_lista(conexion, "Fuente de datos", fuentes, carpeta_sitio, rutas, opciones))
        if opciones["descargar_flujos"] and flujos:
            log.info(f"⬇  Descargando {len(flujos)} flujos...")
            descargas.update(descargar_lista(conexion, "Flujo", flujos, carpeta_sitio, rutas, opciones))

    def estado(item_id):
        return descargas.get(item_id, ("No descargado", ""))

    # --- Tablas de contenido ---
    log.info("🔌 Leyendo conexiones de workbooks y fuentes...")
    tablas["Workbooks"], tablas["Fuentes de datos"], tablas["Conexiones"] = [], [], []
    for wb in workbooks:
        tablas["Workbooks"].append({
            "sitio": nombre_sitio, "id": wb.id, "nombre": wb.name,
            "proyecto": rutas.get(wb.project_id, "Espacio personal"),
            "propietario": nombre_usuario.get(wb.owner_id, wb.owner_id), "creado": fecha(wb.created_at),
            "actualizado": fecha(wb.updated_at), "tamano": wb.size, "mostrar_pestanas": wb.show_tabs,
            "descripcion": wb.description, "etiquetas": ", ".join(sorted(wb.tags or [])),
            "url_web": wb.webpage_url, "estado_descarga": estado(wb.id)[0], "archivo_local": estado(wb.id)[1],
        })
        conexiones = intentar(f"conexiones del workbook {wb.name}", sitio,
                              lambda: server.workbooks.populate_connections(wb) or wb.connections, [])
        tablas["Conexiones"] += filas_conexiones(nombre_sitio, "Workbook", wb, rutas, conexiones)

    for ds in fuentes:
        tablas["Fuentes de datos"].append({
            "sitio": nombre_sitio, "id": ds.id, "nombre": ds.name,
            "proyecto": rutas.get(ds.project_id, "Espacio personal"),
            "propietario": nombre_usuario.get(ds.owner_id, ds.owner_id), "tipo": ds.datasource_type,
            "tiene_extracto": ds.has_extracts, "certificada": ds.certified, "creado": fecha(ds.created_at),
            "actualizado": fecha(ds.updated_at), "descripcion": ds.description,
            "etiquetas": ", ".join(sorted(ds.tags or [])), "url_web": ds.webpage_url,
            "estado_descarga": estado(ds.id)[0], "archivo_local": estado(ds.id)[1],
        })
        conexiones = intentar(f"conexiones de la fuente {ds.name}", sitio,
                              lambda: server.datasources.populate_connections(ds) or ds.connections, [])
        tablas["Conexiones"] += filas_conexiones(nombre_sitio, "Fuente de datos", ds, rutas, conexiones)

    tablas["Vistas"] = [{
        "sitio": nombre_sitio, "id": v.id, "nombre": v.name, "workbook": ids_wb[v.workbook_id].name,
        "proyecto": rutas.get(ids_wb[v.workbook_id].project_id, "Espacio personal"),
        "propietario": nombre_usuario.get(v.owner_id, v.owner_id), "tipo_hoja": getattr(v, "sheet_type", None),
        "total_visitas": getattr(v, "_total_views", None),
        "creado": fecha(v.created_at), "actualizado": fecha(v.updated_at), "url_web": conexion.url_vista(v.content_url),
    } for v in vistas]

    tablas["Flujos"] = [{
        "sitio": nombre_sitio, "id": f.id, "nombre": f.name, "proyecto": rutas.get(f.project_id, "Espacio personal"),
        "propietario": nombre_usuario.get(f.owner_id, f.owner_id), "creado": fecha(f.created_at),
        "actualizado": fecha(f.updated_at), "descripcion": f.description, "url_web": f.webpage_url,
        "estado_descarga": estado(f.id)[0], "archivo_local": estado(f.id)[1],
    } for f in flujos]

    # --- Tareas y suscripciones ---
    nombres = {i.id: i.name for i in workbooks + fuentes + flujos}
    tareas = intentar("tareas de actualización de extractos", sitio, lambda: list(server.tasks.get()[0]), [])
    tablas["Tareas extraccion"] = [{
        "sitio": nombre_sitio, "id": t.id, "tipo_tarea": t.task_type, "tipo_contenido": t.target.type,
        "contenido": nombres.get(t.target.id, t.target.id), "prioridad": t.priority,
        "fallos_consecutivos": t.consecutive_failed_count,
        "programacion": getattr(t.schedule_item, "name", None) if t.schedule_item else None,
        "frecuencia": describir_programacion(t.schedule_item),
        "proxima_ejecucion": fecha(getattr(t.schedule_item, "next_run_at", None)) if t.schedule_item else None,
    } for t in tareas]

    suscripciones = intentar("suscripciones", sitio, lambda: listar(server.subscriptions), [])
    tablas["Suscripciones"] = [{
        "sitio": nombre_sitio, "id": s.id, "asunto": s.subject,
        "usuario": nombre_usuario.get(s.user_id, s.user_id), "tipo_contenido": s.target.type,
        "contenido": nombres.get(s.target.id, s.target.id), "programacion": getattr(s, "schedule_id", None),
    } for s in suscripciones]

    # --- Campos por hoja (inventario Copec) ---
    tablas["Campos"] = []
    if opciones["inventario_campos"] and workbooks:
        log.info("🧬 Leyendo campos por hoja desde la Metadata API...")
        tablas["Campos"] = campos_por_hoja(conexion, workbooks, rutas)

    tablas["_resumen"] = {
        "sitio": nombre_sitio, "proyectos": len(tablas["Proyectos"]), "workbooks": len(workbooks),
        "vistas": len(vistas), "fuentes_de_datos": len(fuentes), "flujos": len(flujos),
        "usuarios": len(usuarios), "grupos": len(grupos),
        "descargados": sum(1 for e, _ in descargas.values() if e == "Descargado"),
        "ya_existian": sum(1 for e, _ in descargas.values() if e == "Ya existía"),
        "con_error": sum(1 for e, _ in descargas.values() if e == "Error"),
    }
    return tablas


def _miembros_grupo(server, grupo):
    server.groups.populate_users(grupo)
    return list(grupo.users)


def filas_conexiones(sitio, tipo, item, rutas, conexiones):
    return [{
        "sitio": sitio, "tipo_contenido": tipo, "contenido": item.name,
        "proyecto": rutas.get(item.project_id, "Espacio personal"), "id_conexion": c.id,
        "tipo_conexion": c.connection_type, "servidor": c.server_address, "puerto": c.server_port,
        "usuario_bd": c.username, "password_embebida": c.embed_password,
        "fuente_datos": getattr(c, "datasource_name", None),
    } for c in (conexiones or [])]


def describir_programacion(programacion):
    if not programacion:
        return None
    intervalo = getattr(programacion, "interval_item", None)
    if intervalo is None:
        return None
    partes = [type(intervalo).__name__.replace("Interval", "")]
    for atributo in ("start_time", "interval", "_interval"):
        valor = getattr(intervalo, atributo, None)
        if valor:
            partes.append(str(valor))
    return " ".join(partes)


# ---------------------- #
#        SALIDAS         #
# ---------------------- #

ORDEN_HOJAS = ["Proyectos", "Workbooks", "Vistas", "Fuentes de datos", "Conexiones", "Flujos", "Campos",
               "Usuarios", "Grupos", "Usuarios por grupo", "Tareas extraccion", "Suscripciones"]


def guardar_inventario(inventarios, info, ruta):
    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        resumen = pd.DataFrame([inv["_resumen"] for inv in inventarios])
        encabezado = pd.DataFrame(list(info.items()), columns=["dato", "valor"])
        encabezado.to_excel(writer, sheet_name="Resumen", index=False)
        resumen.to_excel(writer, sheet_name="Resumen", index=False, startrow=len(encabezado) + 2)

        for hoja in ORDEN_HOJAS:
            filas = [f for inv in inventarios for f in inv.get(hoja, [])]
            escribir_hoja(writer, hoja, pd.DataFrame(filas))
        escribir_hoja(writer, "Pendientes", pd.DataFrame(pendientes))

        for ws in writer.book.worksheets:
            ajustar_columnas(ws)


def escribir_hoja(writer, nombre, df):
    if df.empty:
        df = pd.DataFrame({"sin datos": []})
    df = df.apply(lambda col: col.map(lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v))
    df.to_excel(writer, sheet_name=nombre, index=False)
    ws = writer.sheets[nombre]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def ajustar_columnas(ws):
    for columna in ws.columns:
        largo = max((len(str(c.value)) for c in list(columna)[:200] if c.value is not None), default=8)
        ws.column_dimensions[columna[0].column_letter].width = min(max(largo + 2, 10), 60)


def guardar_pendientes(ruta):
    columnas = ["fecha", "sitio", "tipo", "proyecto", "nombre", "id", "url_web", "error", "accion_sugerida"]
    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        escritor = csv.DictWriter(f, fieldnames=columnas, delimiter=";")
        escritor.writeheader()
        escritor.writerows(pendientes)


# ---------------------- #
#          MAIN          #
# ---------------------- #

def main():
    parser = argparse.ArgumentParser(description="Exporta el contenido de Tableau Server / Cloud a una carpeta local.")
    parser.add_argument("--config", default="config.ini", help="Archivo de configuración (por defecto config.ini)")
    parser.add_argument("--solo-inventario", action="store_true", help="Genera solo el inventario, sin descargar")
    args = parser.parse_args()

    cfg = leer_config(args.config)
    t, e = cfg["tableau"], cfg["exportacion"]
    opciones = {
        "descargar_workbooks": e.getboolean("descargar_workbooks", True),
        "descargar_fuentes": e.getboolean("descargar_fuentes", True),
        "descargar_flujos": e.getboolean("descargar_flujos", True),
        "incluir_extractos": e.getboolean("incluir_extractos", True),
        "sobrescribir": e.getboolean("sobrescribir", False),
        "inventario_campos": e.getboolean("inventario_campos", True),
        "filtro_proyectos": [p.strip() for p in e.get("filtro_proyectos", "").split(",") if p.strip()],
    }

    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = os.path.abspath(e.get("carpeta_salida", "./exportacion"))
    ruta_log = configurar_logs(carpeta, marca)
    inicio = time.time()

    token_name = t.get("token_name") or os.environ.get("TABLEAU_TOKEN_NAME") or input("Nombre del token (PAT): ").strip()
    token_value = t.get("token_value") or os.environ.get("TABLEAU_TOKEN_VALUE") or getpass.getpass("Valor del token (PAT): ").strip()
    sitios = [s.strip() for s in t.get("sites", "").split(",")] or [""]

    log.info("=" * 70)
    log.info(f"Exportador Tableau | {t['server_url']} | sitios: {', '.join(s or 'Default' for s in sitios)}")
    log.info(f"Carpeta de salida: {carpeta}")
    log.info("=" * 70)

    conexion = Conexion(t["server_url"], token_name, token_value, t.getboolean("verify_ssl", True))
    inventarios = []

    try:
        if sitios == ["*"]:
            conexion.conectar("")
            sitios = [s.content_url for s in listar(conexion.server.sites)]
            log.info(f"Sitios encontrados: {', '.join(s or 'Default' for s in sitios)}")

        for sitio in sitios:
            log.info(f"\n🔨 Procesando sitio: {sitio or 'Default'}")
            try:
                conexion.conectar(sitio)
            except Exception as ex:
                log.error(f"   🚨 No se pudo iniciar sesión en el sitio '{sitio or 'Default'}': {ex}")
                registrar_pendiente(sitio, "Sitio", "", sitio or "Default", "", t["server_url"], ex,
                                    "Revisar nombre del sitio y que el token pertenezca a ese sitio.")
                continue
            try:
                carpeta_sitio = os.path.join(carpeta, limpiar_nombre(sitio or "Default"))
                inventarios.append(inventario_sitio(conexion, opciones, carpeta_sitio, args.solo_inventario))
            except Exception as ex:
                log.exception(f"   🚨 Error procesando el sitio '{sitio or 'Default'}': {ex}")
                registrar_pendiente(sitio, "Sitio", "", sitio or "Default", "", t["server_url"], ex,
                                    "Revisar el log y re-ejecutar; los archivos ya descargados se saltan.")

    except KeyboardInterrupt:
        log.warning("\n⛔ Ejecución interrumpida por el usuario. Se guarda lo obtenido hasta ahora.")
    finally:
        conexion.desconectar()
        minutos, segundos = divmod(int(time.time() - inicio), 60)

        ruta_pendientes = os.path.join(carpeta, f"pendientes_{marca}.csv")
        guardar_pendientes(ruta_pendientes)

        ruta_inventario = os.path.join(carpeta, f"inventario_{marca}.xlsx")
        info = {
            "servidor": t["server_url"], "fecha_extraccion": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "version_api": conexion.server.version, "duracion": f"{minutos:02d}:{segundos:02d}",
            "incluye_extractos": opciones["incluir_extractos"], "pendientes": len(pendientes),
        }
        try:
            guardar_inventario(inventarios, info, ruta_inventario)
        except Exception as ex:
            log.exception(f"🚨 No se pudo guardar el inventario Excel: {ex}")

        log.info("\n" + "=" * 70)
        for inv in inventarios:
            r = inv["_resumen"]
            log.info(f"Sitio {r['sitio']}: {r['workbooks']} workbooks, {r['fuentes_de_datos']} fuentes, "
                     f"{r['flujos']} flujos | descargados {r['descargados']}, ya existían {r['ya_existian']}, "
                     f"con error {r['con_error']}")
        log.info(f"📊 Inventario : {ruta_inventario}")
        log.info(f"📝 Pendientes : {ruta_pendientes} ({len(pendientes)} elementos)")
        log.info(f"🗒  Log        : {ruta_log}")
        log.info(f"⏱  Duración   : {minutos:02d}:{segundos:02d}")
        log.info("=" * 70)


if __name__ == "__main__":
    main()
