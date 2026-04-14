"""
Buscador de Empleos — Ingeniero de Datos Senior (Colombia, Remoto, Sin Inglés)
===============================================================================
Portales:
  - LinkedIn      : requests + BeautifulSoup (paginación API guest)
  - Computrabajo  : requests + BeautifulSoup (paginación)
  - Elempleo      : Playwright (requiere JS rendering)

Filtra por inglés requerido, modalidad presencial y empresas excluidas.
Calcula compatibilidad con el perfil de Sebastian Castro y exporta a Excel.
"""

import requests
import time
import random
import json
import re
from datetime import datetime
from dataclasses import dataclass, field
from typing import Optional
from urllib.parse import urljoin, urlparse, urlunparse, quote as url_quote

from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from rich.console import Console
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn

# ---------------------------------------------------------------------------
# Configuración global
# ---------------------------------------------------------------------------

# Búsquedas a realizar en cada portal
KEYWORDS_BUSQUEDA = [
    "Ingeniero de Datos",
    "Data Engineer",
    "Analista de Datos",
    "Analista de Base de Datos",
    "Analytics Engineer",
    "Data Platform Engineer",
    "Ingeniero ETL",
]

TITULOS_RELEVANTES = [
    "ingeniero de datos", "data engineer", "analytics engineer",
    "data platform", "analista de datos", "data analyst",
    "analista de base de datos", "analista bases de datos",
    "etl", "pipeline de datos", "data pipeline",
    "bigquery", "spark", "databricks", "dbt", "airflow",
    "data warehouse", "lakehouse", "data lake",
    "machine learning", "data scientist", "científico de datos",
    "base de datos", "bases de datos",
]

# Modalidades aceptadas (incluir en descripción/título para no descartar)
MODALIDADES_ACEPTADAS = ["remoto", "remote", "híbrido", "hibrido", "teletrabajo", "work from home"]

PALABRAS_INGLES = [
    # Patrones en español (portales colombianos)
    "inglés requerido", "inglés obligatorio", "ingles requerido",
    "nivel de inglés", "nivel ingles", "manejo de inglés", "manejo ingles",
    "inglés avanzado", "ingles avanzado",
    "inglés intermedio", "ingles intermedio",
    "bilingüe", "bilingual",
    # Niveles explícitos en español
    "b1", "b2", "c1", "c2", "a2",
    "nivel de inglés: b", "nivel de inglés: c",
    # Patrón "Idiomas: Inglés" muy común en portales colombianos
    "idiomas: inglés", "idiomas: ingles",
    "idioma: inglés", "idioma: ingles",
    "idiomas inglés", "idiomas ingles",
    "idioma inglés", "idioma ingles",
    # Patrones en inglés (descripciones en inglés = pide inglés)
    "english required", "english mandatory", "fluent in english",
    "advanced english", "intermediate english",
    "english proficiency", "proficiency in english",
    "english skills", "english fluency",
    "good level of english", "great level of english",
    "advanced proficiency in english",
    "communication skills in english",
    "professional communication skills in english",
    "upper-intermediate english", "upper intermediate english",
    "spoken and written english", "written and spoken english",
    " in english",          # "communicate in english", "experience in english"
    "both english",         # "skills in both english"
    "english (written", "english (spoken",
    "language advanced",    # "Language Advanced 80-95%"
    "language: advanced",
    "language: english",
]

EXCEPCIONES_INGLES = [
    "no se requiere inglés", "no requiere inglés",
    "inglés no requerido", "no es necesario inglés",
    "inglés no es", "no exige inglés",
    "no requiere ingles", "sin inglés",
]

# Empresas siempre excluidas (requieren inglés implícitamente o no aplican al perfil)
EMPRESAS_EXCLUIDAS = {
    "bairesdev", "baires dev", "bairesdev colombia",
}

# Patrones que indican modalidad 100% presencial
PALABRAS_PRESENCIAL = [
    "100% presencial", "modalidad: presencial", "modalidad de trabajo: presencial",
    "trabajo presencial", "cargo presencial", "puesto presencial",
    "presencial obligatorio", "asistencia presencial",
    "on-site only", "onsite only", "fully on-site", "fully onsite",
    "work on-site", "in-office", "in office only",
    "no teletrabajo", "no tiene opción de teletrabajo",
]

# Palabras que SOLO existen en inglés (heurística de descripción en inglés)
_TOKENS_EN = [
    " the ", " and ", " for ", " with ", " you ", " our ",
    " are ", " will ", " have ", " this ", " that ", " from ",
    " your ", " we ", " they ", " their ", " been ",
]

HEADERS_BASE = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-CO,es;q=0.9,en;q=0.4",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
}

console = Console()

# ---------------------------------------------------------------------------
# Perfil de Sebastian Castro
# ---------------------------------------------------------------------------

PERFIL = {
    "titulo_exacto": {
        "senior data engineer": 25, "ingeniero de datos senior": 25,
        "data engineer": 20,        "ingeniero de datos": 20,
        "analytics engineer": 18,   "data platform engineer": 18,
        "ingeniero etl": 15,        "data architect": 15,
        "analista de datos": 10,    "data analyst": 10,
        "data scientist": 12,       "científico de datos": 12,
    },
    "tecnologias_core": {
        "bigquery": 10, "snowflake": 10, "gcp": 10, "google cloud": 10,
        "dbt": 10, "python": 10,
    },
    "tecnologias_importantes": {
        "spark": 8, "databricks": 8, "aws": 8, "pyspark": 8,
        "airflow": 7, "kafka": 7,
    },
    "herramientas": {
        "sql": 5, "etl": 5, "elt": 5, "pipeline": 5, "data warehouse": 5,
        "cloud run": 5, "cloud functions": 5, "docker": 4,
        "fastapi": 4, "pandas": 4, "postgresql": 4,
        "ci/cd": 3, "github actions": 3, "pytest": 3,
    },
    "senior_bonus": {
        "senior": 8, "sr.": 8, "lead": 6, "manager": 6,
        "tech lead": 6, "staff": 5, "principal": 5,
    },
    "junior_penalizacion": {
        "junior": -15, "jr.": -15, "trainee": -20,
        "aprendiz": -20, "practicante": -20, "internship": -20,
    },
    "max_puntaje": 99,
}

NIVELES_COMPAT = {
    "Alta":  {"min": 65, "color_excel": "70AD47", "color_rich": "green"},
    "Media": {"min": 35, "color_excel": "FFD966", "color_rich": "yellow"},
    "Baja":  {"min": 0,  "color_excel": "FF7676", "color_rich": "red"},
}

# ---------------------------------------------------------------------------
# Modelo de datos
# ---------------------------------------------------------------------------

@dataclass
class Oferta:
    titulo: str
    empresa: str
    ubicacion: str
    modalidad: str
    descripcion: str
    url: str
    fuente: str
    fecha_publicacion: str = ""
    salario: str = "No especificado"
    pide_ingles: bool = False
    compatibilidad_pct: int = 0
    nivel_compat: str = "Baja"
    tecnologias_encontradas: str = ""
    fecha_extraccion: str = field(
        default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M")
    )


# ---------------------------------------------------------------------------
# Motor de compatibilidad
# ---------------------------------------------------------------------------

def calcular_compatibilidad(oferta: "Oferta") -> tuple[int, str, str]:
    titulo_lower = oferta.titulo.lower()
    texto = (oferta.titulo + " " + oferta.descripcion + " " + oferta.empresa).lower()

    puntaje = 0
    techs = []

    # Bonus título exacto
    mejor_bonus = 0
    for key, pts in PERFIL["titulo_exacto"].items():
        if key in titulo_lower and pts > mejor_bonus:
            mejor_bonus = pts
    puntaje += mejor_bonus

    for tech, pts in PERFIL["tecnologias_core"].items():
        if tech in texto:
            puntaje += pts
            techs.append(tech.upper())

    for tech, pts in PERFIL["tecnologias_importantes"].items():
        if tech in texto:
            puntaje += pts
            techs.append(tech.capitalize())

    for tech, pts in PERFIL["herramientas"].items():
        if tech in texto:
            puntaje += pts
            if tech not in ("etl", "elt", "sql", "pipeline"):
                techs.append(tech)

    for kw, pts in PERFIL["senior_bonus"].items():
        if kw in texto:
            puntaje += pts
            break

    for kw, pen in PERFIL["junior_penalizacion"].items():
        if kw in texto:
            puntaje += pen
            break

    pct = min(100, max(0, round(puntaje * 100 / PERFIL["max_puntaje"])))
    nivel = next(n for n, cfg in NIVELES_COMPAT.items() if pct >= cfg["min"])
    return pct, nivel, ", ".join(dict.fromkeys(techs))


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def get_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS_BASE)
    return s


def pausa(a: float = 1.5, b: float = 3.5):
    time.sleep(random.uniform(a, b))


def detectar_ingles(texto: str) -> bool:
    t = " " + texto.lower() + " "
    for exc in EXCEPCIONES_INGLES:
        if exc in t:
            return False
    for pal in PALABRAS_INGLES:
        if pal in t:
            idx = t.find(pal)
            ctx = t[max(0, idx - 70): idx]
            if not any(neg in ctx for neg in ["no ", "sin ", "no se ", "no requiere", "no exige"]):
                return True
    return False


def descripcion_en_ingles(texto: str, umbral: int = 6) -> bool:
    """True si la descripción está escrita principalmente en inglés."""
    t = " " + texto.lower() + " "
    return sum(1 for tok in _TOKENS_EN if tok in t) >= umbral


def detectar_presencial(texto: str) -> bool:
    """True si la oferta indica modalidad 100% presencial."""
    t = texto.lower()
    return any(p in t for p in PALABRAS_PRESENCIAL)


def empresa_excluida(empresa: str) -> bool:
    """True si la empresa está en la lista negra."""
    return empresa.strip().lower() in EMPRESAS_EXCLUIDAS


def normalizar_url(url: str) -> str:
    """Elimina query params y fragmentos de la URL para evitar duplicados por tracking."""
    if not url:
        return url
    p = urlparse(url)
    return urlunparse((p.scheme, p.netloc, p.path, "", "", ""))


def limpiar(texto: str) -> str:
    return re.sub(r"\s+", " ", (texto or "").strip())


def es_relevante(titulo: str) -> bool:
    t = titulo.lower()
    return any(p in t for p in TITULOS_RELEVANTES)


def deduplicar(ofertas: list[Oferta]) -> list[Oferta]:
    vistas, resultado = set(), []
    for o in ofertas:
        clave = (o.titulo.lower()[:60], o.empresa.lower()[:40])
        if clave not in vistas:
            vistas.add(clave)
            resultado.append(o)
    return resultado


# ---------------------------------------------------------------------------
# Scraper: LinkedIn — paginación completa
# ---------------------------------------------------------------------------

class ScraperLinkedIn:
    SEARCH_URL = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
    DETAIL_URL = "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{}"

    def buscar(self, session: requests.Session, max_total: int = 80) -> list[Oferta]:
        ofertas = []
        # f_WT=2 → Remoto, f_WT=3 → Híbrido (buscamos ambos)
        for modalidad_code in ["2", "3"]:
            for keyword in KEYWORDS_BUSQUEDA:
                start = 0
                while len(ofertas) < max_total:
                    params = {
                        "keywords": keyword,
                        "location": "Colombia",
                        "f_WT": modalidad_code,
                        "f_TPR": "r2592000",  # Últimos 30 días
                        "start": str(start),
                        "count": "25",
                    }
                    try:
                        resp = session.get(self.SEARCH_URL, params=params, timeout=15)
                        if resp.status_code != 200:
                            break
                        soup = BeautifulSoup(resp.text, "lxml")
                        tarjetas = soup.select("li")
                        if not tarjetas:
                            break
                        nuevas = 0
                        for tarjeta in tarjetas:
                            o = self._parsear(tarjeta, session)
                            if o:
                                modalidad_label = "Remoto" if modalidad_code == "2" else "Híbrido"
                                o.modalidad = modalidad_label
                                ofertas.append(o)
                                nuevas += 1
                        if nuevas == 0:
                            break
                        start += 25
                        pausa(1.0, 2.5)
                    except requests.RequestException:
                        break
                pausa()
        return deduplicar(ofertas)

    def _parsear(self, tarjeta, session: requests.Session) -> Optional[Oferta]:
        try:
            job_id_tag  = tarjeta.select_one("[data-entity-urn]")
            titulo_tag  = tarjeta.select_one(".base-search-card__title")
            empresa_tag = tarjeta.select_one(".base-search-card__subtitle")
            loc_tag     = tarjeta.select_one(".job-search-card__location")
            link_tag    = tarjeta.select_one("a.base-card__full-link")
            fecha_tag   = tarjeta.select_one("time")

            titulo = limpiar(titulo_tag.get_text() if titulo_tag else "")
            if not titulo or not es_relevante(titulo):
                return None

            empresa  = limpiar(empresa_tag.get_text() if empresa_tag else "")
            ubicacion= limpiar(loc_tag.get_text() if loc_tag else "")
            url      = normalizar_url(link_tag["href"] if link_tag else "")
            fecha    = fecha_tag.get("datetime", "") if fecha_tag else ""
            desc     = self._descripcion(job_id_tag, session)

            texto_completo = titulo + " " + empresa + " " + desc
            if empresa_excluida(empresa):
                return None
            if detectar_ingles(texto_completo):
                return None
            if descripcion_en_ingles(desc):
                return None
            if detectar_presencial(texto_completo):
                return None

            return Oferta(
                titulo=titulo, empresa=empresa, ubicacion=ubicacion,
                modalidad="Remoto", descripcion=desc[:700],
                url=url, fuente="LinkedIn", fecha_publicacion=fecha,
            )
        except Exception:
            return None

    def _descripcion(self, tag, session: requests.Session) -> str:
        try:
            urn = (tag.get("data-entity-urn", "") if tag else "")
            job_id = urn.split(":")[-1]
            if not job_id:
                return ""
            resp = session.get(self.DETAIL_URL.format(job_id), timeout=10)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, "lxml")
                el = soup.select_one(".show-more-less-html__markup")
                return limpiar(el.get_text() if el else "")
        except Exception:
            pass
        return ""


# ---------------------------------------------------------------------------
# Scraper: Computrabajo — paginación completa
# ---------------------------------------------------------------------------

class ScraperComputrabajo:
    BASE_URL = "https://co.computrabajo.com/trabajo-de-{termino}"

    # Términos de búsqueda
    TERMINOS = [
        "ingeniero-de-datos",
        "data-engineer",
        "analista-de-datos",
        "analista-base-de-datos",
        "analytics-engineer",
        "etl-developer",
    ]

    def buscar(self, session: requests.Session, max_total: int = 60) -> list[Oferta]:
        ofertas = []
        # Pasada 1: solo remoto (teletrabajo=1) — todos los términos
        # Pasada 2: sin filtro de teletrabajo (captura híbridos/presencial) — términos principales
        pasadas = [
            (self.TERMINOS, {"teletrabajo": "1", "pubdate": "30"}, "Remoto"),
            (self.TERMINOS[:3], {"pubdate": "30"}, "Híbrido"),   # términos principales sin filtro
        ]
        for terminos_pasada, params_base, modalidad_default in pasadas:
            for termino in terminos_pasada:
                pagina = 1
                while len(ofertas) < max_total:
                    url = self.BASE_URL.format(termino=termino)
                    params = {**params_base, "p": str(pagina)}
                    try:
                        resp = session.get(url, params=params, timeout=15)
                        if resp.status_code != 200:
                            break
                        soup = BeautifulSoup(resp.text, "lxml")
                        tarjetas = soup.select("article.box_offer, div.box_offer")
                        if not tarjetas:
                            break
                        nuevas = 0
                        for t in tarjetas:
                            o = self._parsear(t, modalidad_default)
                            if o:
                                ofertas.append(o)
                                nuevas += 1
                        sig = soup.select_one("a[title='Siguiente página'], a.js-navigate-page[rel='next']")
                        if not sig or nuevas == 0:
                            break
                        pagina += 1
                        pausa(1.0, 2.5)
                    except requests.RequestException:
                        break
                pausa()
        return deduplicar(ofertas)

    def _parsear(self, tarjeta, modalidad_default: str = "Remoto") -> Optional[Oferta]:
        try:
            titulo_tag  = tarjeta.select_one("h2 a.js-o-link, h2 a.fc_base")
            empresa_tag = tarjeta.select_one("a[offer-grid-article-company-url], p.dFlex a.fc_base")
            loc_tag     = tarjeta.select_one("p.fs16.fc_base.mt5 span.mr10, p.fs16 span.mr10")
            sal_icon    = tarjeta.select_one("span.dIB .i_salary")
            salario_tag = sal_icon.find_parent("span", class_="dIB") if sal_icon else None
            fecha_tag   = tarjeta.select_one("p.fs13.fc_aux, p.fs13.fc_aux.mt15")

            titulo = limpiar(titulo_tag.get_text() if titulo_tag else "")
            if not titulo or not es_relevante(titulo):
                return None

            empresa  = limpiar(empresa_tag.get_text() if empresa_tag else "")
            ubicacion= limpiar(loc_tag.get_text() if loc_tag else "")
            sal_raw  = limpiar(salario_tag.get_text() if salario_tag else "")
            salario  = re.sub(r"\s*\S*salary\S*\s*", "", sal_raw).strip() or "No especificado"
            fecha    = limpiar(fecha_tag.get_text() if fecha_tag else "")
            href     = titulo_tag.get("href", "") if titulo_tag else ""
            # Limpiar fragmento (#) que Computrabajo agrega para tracking
            href_clean = href.split("#")[0] if href else ""
            url      = normalizar_url(urljoin("https://co.computrabajo.com", href_clean)) if href_clean else ""

            if empresa_excluida(empresa):
                return None
            if detectar_ingles(titulo + " " + empresa):
                return None

            return Oferta(
                titulo=titulo, empresa=empresa, ubicacion=ubicacion,
                modalidad=modalidad_default, descripcion="",
                url=url, fuente="Computrabajo",
                fecha_publicacion=fecha, salario=salario,
            )
        except Exception:
            return None


# ---------------------------------------------------------------------------
# Scraper: Elempleo — Playwright (JS rendering)
# ---------------------------------------------------------------------------

class ScraperElempleo:

    URLS_BUSQUEDA = [
        # tipoempleo=6 → teletrabajo/remoto | tipoempleo=5 → híbrido
        "https://www.elempleo.com/co/ofertas-empleo/ingeniero-de-datos?idPais=1&tipoempleo=6",
        "https://www.elempleo.com/co/ofertas-empleo/ingeniero-de-datos?idPais=1&tipoempleo=5",
        "https://www.elempleo.com/co/ofertas-empleo/data-engineer?idPais=1&tipoempleo=6",
        "https://www.elempleo.com/co/ofertas-empleo/data-engineer?idPais=1&tipoempleo=5",
        "https://www.elempleo.com/co/ofertas-empleo/analista-de-datos?idPais=1&tipoempleo=6",
        "https://www.elempleo.com/co/ofertas-empleo/analista-base-datos?idPais=1&tipoempleo=6",
    ]

    # Selectores candidatos para tarjetas de oferta
    SEL_TARJETA = (
        "article[class*='offer'], div[class*='offerCard'], "
        "li[class*='offer'], div[class*='job-item'], "
        "div[class*='jobCard'], section[class*='offer']"
    )

    def buscar(self, page, max_total: int = 60) -> list[Oferta]:
        ofertas = []
        for url_base in self.URLS_BUSQUEDA:
            pagina = 1
            while len(ofertas) < max_total:
                url = url_base if pagina == 1 else f"{url_base}&pagina={pagina}"
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=20000)
                    page.wait_for_timeout(3000)

                    # Esperar tarjetas
                    try:
                        page.wait_for_selector(self.SEL_TARJETA, timeout=8000)
                    except PWTimeout:
                        # Intentar scroll para activar lazy load
                        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                        page.wait_for_timeout(2000)

                    soup = BeautifulSoup(page.content(), "lxml")

                    # Intentar JSON-LD primero (más fiable)
                    extraidas = self._json_ld(soup)
                    if extraidas:
                        ofertas.extend(extraidas)
                        pagina += 1
                        pausa(1.5, 3.0)
                        continue

                    # Fallback: tarjetas HTML
                    tarjetas = soup.select(self.SEL_TARJETA)
                    if not tarjetas:
                        break

                    nuevas = 0
                    for t in tarjetas:
                        o = self._parsear(t)
                        if o:
                            ofertas.append(o)
                            nuevas += 1

                    # Botón siguiente
                    siguiente = soup.select_one(
                        "a[aria-label*='siguiente'], a[class*='next'], "
                        "button[aria-label*='siguiente']"
                    )
                    if not siguiente or nuevas == 0:
                        break
                    pagina += 1
                    pausa(1.5, 3.0)
                except Exception:
                    break
            pausa()
        return deduplicar(ofertas)

    def _json_ld(self, soup: BeautifulSoup) -> list[Oferta]:
        ofertas = []
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or "")
                items = data if isinstance(data, list) else [data]
                for item in items:
                    if item.get("@type") != "JobPosting":
                        continue
                    titulo = item.get("title", "")
                    if not titulo or not es_relevante(titulo):
                        continue
                    empresa = item.get("hiringOrganization", {}).get("name", "")
                    loc_obj = item.get("jobLocation", {})
                    if isinstance(loc_obj, list):
                        loc_obj = loc_obj[0] if loc_obj else {}
                    ubicacion = loc_obj.get("address", {}).get("addressLocality", "Colombia")
                    desc = re.sub(r"<[^>]+>", " ", item.get("description", ""))
                    url  = item.get("url") or item.get("sameAs", "")
                    fecha= item.get("datePosted", "")

                    if detectar_ingles(titulo + " " + desc):
                        continue

                    ofertas.append(Oferta(
                        titulo=limpiar(titulo), empresa=limpiar(empresa),
                        ubicacion=limpiar(ubicacion), modalidad="Remoto",
                        descripcion=limpiar(desc)[:700],
                        url=url, fuente="Elempleo", fecha_publicacion=fecha,
                    ))
            except Exception:
                continue
        return ofertas

    def _parsear(self, tarjeta) -> Optional[Oferta]:
        try:
            titulo_tag  = tarjeta.select_one(
                "h2 a, h3 a, [class*='title'] a, a[class*='title'], a[class*='offer']"
            )
            empresa_tag = tarjeta.select_one("[class*='company'], [class*='empresa']")
            loc_tag     = tarjeta.select_one("[class*='location'], [class*='ciudad']")
            salario_tag = tarjeta.select_one("[class*='salary'], [class*='salario']")
            desc_tag    = tarjeta.select_one("[class*='description'], p")
            fecha_tag   = tarjeta.select_one("time, [class*='date'], [class*='fecha']")

            titulo = limpiar(titulo_tag.get_text() if titulo_tag else "")
            if not titulo or not es_relevante(titulo):
                return None

            empresa  = limpiar(empresa_tag.get_text() if empresa_tag else "")
            ubicacion= limpiar(loc_tag.get_text() if loc_tag else "")
            salario  = limpiar(salario_tag.get_text() if salario_tag else "No especificado")
            desc     = limpiar(desc_tag.get_text() if desc_tag else "")
            fecha    = limpiar(fecha_tag.get_text() if fecha_tag else "")
            href     = titulo_tag.get("href", "") if titulo_tag else ""
            url      = normalizar_url(urljoin("https://www.elempleo.com", href)) if href else ""

            if detectar_ingles(titulo + " " + desc):
                return None

            return Oferta(
                titulo=titulo, empresa=empresa, ubicacion=ubicacion,
                modalidad="Remoto", descripcion=desc[:700],
                url=url, fuente="Elempleo",
                fecha_publicacion=fecha, salario=salario,
            )
        except Exception:
            return None


# ---------------------------------------------------------------------------
# Motor principal
# ---------------------------------------------------------------------------

class BuscadorEmpleos:

    def buscar_todo(self, max_por_portal: int = 80) -> list[Oferta]:
        todas: list[Oferta] = []
        session = get_session()

        with Progress(
            SpinnerColumn(), TextColumn("[progress.description]{task.description}"),
            console=console
        ) as prog:

            # ---------------------------------------------------------------
            # Portales via requests (sin navegador)
            # ---------------------------------------------------------------

            # LinkedIn
            t = prog.add_task("Buscando en [bold cyan]LinkedIn[/bold cyan]...", total=None)
            try:
                li_ofertas = ScraperLinkedIn().buscar(session, max_total=max_por_portal)
                self._enriquecer(li_ofertas)
                todas.extend(li_ofertas)
                prog.update(t, description=f"[green]✓ LinkedIn[/green] — {len(li_ofertas)} ofertas")
            except Exception as e:
                prog.update(t, description=f"[red]✗ LinkedIn[/red] — {e}")
            prog.stop_task(t)

            # Computrabajo
            t = prog.add_task("Buscando en [bold cyan]Computrabajo[/bold cyan]...", total=None)
            try:
                ct_ofertas = ScraperComputrabajo().buscar(session, max_total=max_por_portal)
                self._enriquecer(ct_ofertas)
                todas.extend(ct_ofertas)
                prog.update(t, description=f"[green]✓ Computrabajo[/green] — {len(ct_ofertas)} ofertas")
            except Exception as e:
                prog.update(t, description=f"[red]✗ Computrabajo[/red] — {e}")
            prog.stop_task(t)

            # ---------------------------------------------------------------
            # Portales que requieren Playwright (solo Elempleo)
            # ---------------------------------------------------------------
            t_el = prog.add_task("Buscando en [bold cyan]Elempleo[/bold cyan] (Playwright)...", total=None)
            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.launch(
                        headless=True,
                        args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
                    )
                    ctx = browser.new_context(
                        user_agent=HEADERS_BASE["User-Agent"],
                        locale="es-CO",
                        viewport={"width": 1280, "height": 800},
                    )
                    ctx.add_init_script(
                        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
                    )
                    page_el = ctx.new_page()
                    el_ofertas = ScraperElempleo().buscar(page_el, max_total=max_por_portal)
                    page_el.close()
                    browser.close()
                self._enriquecer(el_ofertas)
                todas.extend(el_ofertas)
                prog.update(t_el, description=f"[green]✓ Elempleo[/green] — {len(el_ofertas)} ofertas")
            except Exception as e:
                prog.update(t_el, description=f"[red]✗ Elempleo[/red] — {e}")
            prog.stop_task(t_el)

        # Deduplicar global y ordenar
        todas = deduplicar(todas)
        todas.sort(key=lambda o: o.compatibilidad_pct, reverse=True)
        return todas

    def _enriquecer(self, ofertas: list[Oferta]):
        """Calcula compatibilidad in-place para cada oferta."""
        for o in ofertas:
            pct, nivel, techs = calcular_compatibilidad(o)
            o.compatibilidad_pct = pct
            o.nivel_compat = nivel
            o.tecnologias_encontradas = techs

    # -----------------------------------------------------------------------
    # Display
    # -----------------------------------------------------------------------

    def mostrar_tabla(self, ofertas: list[Oferta]):
        if not ofertas:
            console.print("\n[bold red]No se encontraron ofertas.[/bold red]")
            return

        tabla = Table(
            title=(
                f"\nOfertas para Sebastian Castro — Ordenadas por Compatibilidad\n"
                f"Total: {len(ofertas)} | {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            ),
            show_header=True, header_style="bold magenta",
            show_lines=True, expand=True,
        )
        tabla.add_column("#", width=3)
        tabla.add_column("Compat.", width=9)
        tabla.add_column("Titulo", min_width=28)
        tabla.add_column("Empresa", min_width=18)
        tabla.add_column("Salario", min_width=15)
        tabla.add_column("Tecnologias", min_width=22)
        tabla.add_column("Fuente", width=12)

        for i, o in enumerate(ofertas, 1):
            color = NIVELES_COMPAT[o.nivel_compat]["color_rich"]
            tabla.add_row(
                str(i),
                f"[{color}]{o.compatibilidad_pct}%\n{o.nivel_compat}[/{color}]",
                o.titulo,
                o.empresa or "—",
                o.salario,
                (o.tecnologias_encontradas[:48] + "…")
                    if len(o.tecnologias_encontradas) > 48 else o.tecnologias_encontradas,
                o.fuente,
            )
        console.print(tabla)

    # -----------------------------------------------------------------------
    # Excel
    # -----------------------------------------------------------------------

    def exportar_excel(self, ofertas: list[Oferta], archivo: str = None) -> str:
        if not archivo:
            fecha = datetime.now().strftime("%Y%m%d_%H%M")
            archivo = f"ofertas_sebastian_castro_{fecha}.xlsx"

        wb = Workbook()
        borde = self._borde()

        self._hoja_principal(wb, ofertas, borde)
        altas = [o for o in ofertas if o.nivel_compat == "Alta"]
        self._hoja_filtrada(wb, altas, "Alta Compatibilidad", "375623", borde)
        self._hoja_resumen(wb, ofertas, borde)

        wb.save(archivo)
        return archivo

    def _borde(self):
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def _hoja_principal(self, wb, ofertas, borde):
        ws = wb.active
        ws.title = "Todas las Ofertas"

        # Título
        ws.merge_cells("A1:K1")
        c = ws["A1"]
        c.value = "OFERTAS LABORALES — SEBASTIAN CASTRO | Ingeniero de Datos Senior"
        c.font = Font(bold=True, size=14, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        ws.merge_cells("A2:K2")
        ws["A2"].value = (
            f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  "
            "Portales: LinkedIn • Computrabajo • Elempleo  |  "
            "Filtros: Remoto • Colombia • Sin inglés obligatorio"
        )
        ws["A2"].font = Font(italic=True, size=10, color="595959")
        ws["A2"].alignment = Alignment(horizontal="center")
        ws.row_dimensions[2].height = 16

        cabeceras = [
            "#", "Compat %", "Nivel", "Titulo", "Empresa",
            "Salario", "Ubicacion", "Tecnologias detectadas",
            "Fuente", "Publicado", "URL"
        ]
        anchos = [4, 12, 10, 42, 28, 22, 20, 46, 14, 14, 18]

        for col, (cab, ancho) in enumerate(zip(cabeceras, anchos), 1):
            c = ws.cell(row=3, column=col, value=cab)
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = PatternFill("solid", fgColor="2E75B6")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = borde
            ws.column_dimensions[get_column_letter(col)].width = ancho
        ws.row_dimensions[3].height = 22

        for row_i, o in enumerate(ofertas, 4):
            datos = [
                row_i - 3, o.compatibilidad_pct, o.nivel_compat,
                o.titulo, o.empresa, o.salario, o.ubicacion,
                o.tecnologias_encontradas, o.fuente,
                o.fecha_publicacion, o.url,
            ]
            for col, val in enumerate(datos, 1):
                c = ws.cell(row=row_i, column=col, value=val)
                c.border = borde
                c.alignment = Alignment(vertical="center", wrap_text=(col in (4, 8, 11)))

                if row_i % 2 == 0:
                    c.fill = PatternFill("solid", fgColor="EBF3FB")

                if col == 2:  # Compat %
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    color_font = {"Alta": "375623", "Media": "7F6000", "Baja": "9C0006"}.get(o.nivel_compat, "000000")
                    c.font = Font(bold=True, size=11, color=color_font)

                if col == 3:  # Nivel badge
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.fill = PatternFill("solid", fgColor=NIVELES_COMPAT[o.nivel_compat]["color_excel"])
                    c.font = Font(bold=True, size=10)

                if col == 11 and o.url:  # URL hipervínculo
                    c.hyperlink = o.url
                    c.font = Font(color="0563C1", underline="single")
                    c.value = "Ver oferta →"

            ws.row_dimensions[row_i].height = 38

        ws.freeze_panes = "A4"
        if len(ofertas) > 0:
            ws.conditional_formatting.add(
                f"B4:B{3 + len(ofertas)}",
                ColorScaleRule(
                    start_type="num", start_value=0,  start_color="FF7676",
                    mid_type="num",   mid_value=50,   mid_color="FFD966",
                    end_type="num",   end_value=100,  end_color="70AD47",
                )
            )

    def _hoja_filtrada(self, wb, ofertas, nombre, color_header, borde):
        ws = wb.create_sheet(title=nombre)

        ws.merge_cells("A1:J1")
        ws["A1"].value = f"ALTA COMPATIBILIDAD CON PERFIL SEBASTIAN CASTRO — {len(ofertas)} ofertas"
        ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor=color_header)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        cabeceras = ["#", "Compat%", "Titulo", "Empresa", "Salario", "Tecnologias", "Fuente", "Publicado", "URL"]
        anchos    = [4,   11,        44,       28,        22,        50,            14,       14,          18]

        for col, (cab, ancho) in enumerate(zip(cabeceras, anchos), 1):
            c = ws.cell(row=2, column=col, value=cab)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=color_header)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = borde
            ws.column_dimensions[get_column_letter(col)].width = ancho
        ws.row_dimensions[2].height = 20

        for i, o in enumerate(ofertas, 3):
            datos = [
                i - 2, o.compatibilidad_pct, o.titulo, o.empresa,
                o.salario, o.tecnologias_encontradas, o.fuente,
                o.fecha_publicacion, o.url,
            ]
            for col, val in enumerate(datos, 1):
                c = ws.cell(row=i, column=col, value=val)
                c.border = borde
                c.alignment = Alignment(vertical="center", wrap_text=(col in (3, 6, 9)))
                if i % 2 == 0:
                    c.fill = PatternFill("solid", fgColor="E2EFDA")
                if col == 2:
                    c.font = Font(bold=True, color="375623")
                    c.alignment = Alignment(horizontal="center", vertical="center")
                if col == 9 and o.url:
                    c.hyperlink = o.url
                    c.font = Font(color="0563C1", underline="single")
                    c.value = "Ver oferta →"
            ws.row_dimensions[i].height = 38
        ws.freeze_panes = "A3"

    def _hoja_resumen(self, wb, ofertas, borde):
        ws = wb.create_sheet(title="Resumen")
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 16

        ws.merge_cells("A1:B1")
        ws["A1"].value = "RESUMEN DE BÚSQUEDA"
        ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor="1F4E79")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        def fila(row, label, valor, fondo="F2F2F2"):
            for col, val in enumerate([label, valor], 1):
                c = ws.cell(row=row, column=col, value=val)
                c.fill = PatternFill("solid", fgColor=fondo)
                c.border = borde
                if col == 1:
                    c.font = Font(bold=True)
                else:
                    c.alignment = Alignment(horizontal="center")

        alta  = sum(1 for o in ofertas if o.nivel_compat == "Alta")
        media = sum(1 for o in ofertas if o.nivel_compat == "Media")
        baja  = sum(1 for o in ofertas if o.nivel_compat == "Baja")
        prom  = round(sum(o.compatibilidad_pct for o in ofertas) / len(ofertas)) if ofertas else 0

        fila(3,  "Total ofertas encontradas",      len(ofertas),      "DEEAF1")
        fila(4,  "Alta compatibilidad (≥65%)",     alta,              "E2EFDA")
        fila(5,  "Media compatibilidad (35-64%)",  media,             "FFF2CC")
        fila(6,  "Baja compatibilidad (<35%)",     baja,              "FCE4D6")
        fila(7,  "Compatibilidad promedio",         f"{prom}%",        "DEEAF1")
        fila(8,  "Fecha de búsqueda",              datetime.now().strftime("%Y-%m-%d %H:%M"))

        ws.cell(row=10, column=1, value="Por portal:").font = Font(bold=True, size=11)
        portales = ["LinkedIn", "Computrabajo", "Elempleo"]
        for i, fuente in enumerate(portales, 11):
            fila(i, fuente, sum(1 for o in ofertas if o.fuente == fuente))

        perfil_row = 11 + len(portales) + 1
        ws.cell(row=perfil_row, column=1, value="Perfil evaluado:").font = Font(bold=True, size=11)
        fila(perfil_row + 1, "Nombre",         "Sebastian Castro")
        fila(perfil_row + 2, "Cargo objetivo", "Ingeniero de Datos Senior")
        fila(perfil_row + 3, "Stack principal","GCP, BigQuery, Snowflake, dbt, Python, Spark")
        fila(perfil_row + 4, "Experiencia",    "7+ años")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main():
    console.print(
        "\n[bold green]============================================================[/bold green]"
        "\n[bold green]  Buscador para: Sebastian Castro[/bold green]"
        "\n[bold green]  Ingeniero de Datos Senior[/bold green]"
        "\n[bold green]  GCP • BigQuery • Snowflake • dbt • Python • Spark[/bold green]"
        "\n[bold green]  Colombia | Remoto | Sin inglés obligatorio[/bold green]"
        "\n[bold green]  Portales: LinkedIn • Computrabajo • Elempleo[/bold green]"
        "\n[bold green]============================================================[/bold green]\n"
    )

    buscador = BuscadorEmpleos()
    ofertas  = buscador.buscar_todo(max_por_portal=80)

    if not ofertas:
        console.print("[bold red]No se encontraron ofertas.[/bold red]")
        return

    buscador.mostrar_tabla(ofertas)

    archivo = buscador.exportar_excel(ofertas)
    console.print(f"\n[bold green]Excel exportado:[/bold green] [cyan]{archivo}[/cyan]")

    alta  = [o for o in ofertas if o.nivel_compat == "Alta"]
    media = [o for o in ofertas if o.nivel_compat == "Media"]
    baja  = [o for o in ofertas if o.nivel_compat == "Baja"]
    prom  = round(sum(o.compatibilidad_pct for o in ofertas) / len(ofertas))

    console.print(f"\n[bold]Resumen:[/bold]")
    console.print(f"  [green]Alta  (≥65%)[/green]   → {len(alta)} ofertas")
    console.print(f"  [yellow]Media (35-64%)[/yellow] → {len(media)} ofertas")
    console.print(f"  [red]Baja  (<35%)[/red]    → {len(baja)} ofertas")
    console.print(f"  Compatibilidad promedio → {prom}%")

    if ofertas:
        console.print(f"\n[bold]Top 5 mejores matches:[/bold]")
        for i, o in enumerate(ofertas[:5], 1):
            color = NIVELES_COMPAT[o.nivel_compat]["color_rich"]
            console.print(
                f"  {i}. [{color}]{o.compatibilidad_pct}%[/{color}] "
                f"[bold]{o.titulo}[/bold] @ {o.empresa or '—'} "
                f"([dim]{o.fuente}[/dim])"
            )
            if o.tecnologias_encontradas:
                console.print(f"     Techs: {o.tecnologias_encontradas}")


if __name__ == "__main__":
    main()
