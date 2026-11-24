
import os
import re
import json
import time
import html
import base64
import requests
import pandas as pd

from datetime import datetime, timezone
from collections import defaultdict
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urlunparse, parse_qsl, urlencode

# ============================================================
# GERMANY TECH JOB COLLECTOR - EXPANDED 2026
#
# Sources:
# 1. Bundesagentur für Arbeit
# 2. Arbeitnow
# 3. Adzuna (optional API credentials)
# 4. Jobicy
# 5. RemoteOK
# 6. Greenhouse
# 7. Lever
# 8. Ashby
# 9. SmartRecruiters
# 10. Recruitee
# 11. Personio
# 12. Workable
# 13. Breezy HR
#
# Collects currently available jobs. It cannot recover every
# closed job from earlier in 2026.
# ============================================================

OUTPUT = "Germany_Tech_Jobs_2026.xlsx"
CSV_OUTPUT = "Germany_Tech_Jobs_2026.csv"

START_DATE = datetime(2026, 1, 1, tzinfo=timezone.utc)
END_DATE = datetime(2027, 1, 1, tzinfo=timezone.utc)

REQUEST_DELAY = 0.8
TIMEOUT = 35
MAX_RETRIES = 3

HEADERS = {
    "User-Agent": "GermanyTechJobsResearch/3.0 (public job data research)",
    "Accept": "application/json, text/plain, */*"
}

session = requests.Session()
session.headers.update(HEADERS)

jobs = []
logs = []
stats = defaultdict(lambda: defaultdict(int))
tested_boards = []
seen = set()


# ============================================================
# COMPANY BOARD CANDIDATES
#
# These are candidate board identifiers, not guaranteed active
# boards. The script tests each and logs success/failure.
# Add more verified career-board slugs as you discover them.
# ============================================================

GREENHOUSE_BOARDS = [
    # German technology companies / startups
    "adahealth",
    "adjust",
    "appsflyer",
    "auto1",
    "babbel",
    "blacklane",
    "bonial",
    "clue",
    "cognigy",
    "commercetools",
    "dash0",
    "deepl",
    "enpal",
    "eyeo",
    "factorial",
    "grover",
    "improbable",
    "infarm",
    "infinit",
    "innoscripta",
    "instaffo",
    "kamino",
    "klara",
    "komoot",
    "lilium",
    "mambu",
    "navan",
    "omio",
    "personio",
    "planradar",
    "quantilope",
    "quantum-systems",
    "relayr",
    "robinhood",
    "sennder",
    "teamviewer",
    "trade-republic",
    "trivago",
    "verimi",
    "volocopter",
    "wefox",
    "xentral",

    # International software / SaaS
    "1password",
    "airtable",
    "algolia",
    "appsmith",
    "benchling",
    "bolt",
    "box",
    "brex",
    "calendly",
    "circleci",
    "clickhouse",
    "cockroachlabs",
    "crunchydata",
    "dbt-labs",
    "docusign",
    "drata",
    "fivetran",
    "flyio",
    "front",
    "gusto",
    "launchdarkly",
    "linear",
    "loom",
    "monday",
    "mixpanel",
    "netlify",
    "okta",
    "onepassword",
    "plaid",
    "posthog",
    "retool",
    "sentry",
    "snowflake",
    "sourcegraph",
    "supabase",
    "temporal",
    "thoughtspot",
    "tines",
    "vercel",
    "zapier",
    "zendesk",

    # AI / machine learning
    "adept",
    "ai21labs",
    "anyscale",
    "cohere",
    "elevenlabs",
    "huggingface",
    "mistral",
    "modal",
    "perplexity",
    "replicate",
    "scaleai",
    "togetherai",
    "weightsandbiases",

    # Cybersecurity / infrastructure
    "1password",
    "abnormal",
    "aqua",
    "arcticwolf",
    "crowdstrike",
    "fastly",
    "fortinet",
    "lacework",
    "proton",
    "wiz"
]

LEVER_BOARDS = [
    "spotify", "revolut", "klarna", "doctolib",
    "backmarket", "mistral", "personio", "pleo",
    "typeform", "contentsquare", "qonto", "aircall",
    "n26", "getaround", "ovhcloud", "checkout",
    "factorial", "papernest", "payfit", "algolia",
    "snyk", "remote", "deel", "safetywing",
    "gocardless", "monzo", "wise", "bolt"
]

ASHBY_BOARDS = [
    "linear", "notion", "ramp", "deel", "airbyte",
    "supabase", "webflow", "mercury", "rippling",
    "cursor", "perplexity", "retool", "vercel",
    "sourcegraph", "scale", "gusto", "remote",
    "sentry", "posthog", "cal", "langchain",
    "modal", "replicate", " elevenlabs".strip(),
    "replit", "loom", "incidentio", "clerk",
    "railway", "resend", "turso", " neon".strip()
]

SMARTRECRUITERS_BOARDS = [
    "SAP", "Bosch", "Siemens", "BMW", "Mercedes-Benz",
    "DeutscheBank", "DeutscheTelekom", "DeliveryHero",
    "HelloFresh", "Zalando", "Infineon", "Continental",
    "Adidas", "Allianz", "DHL", "Roche", "Bayer",
    "EON", "Henkel", "Fresenius", "Lufthansa",
    "Capgemini", "KPMG", "EY", "PwC", "Deloitte",
    "Vodafone", "Ericsson", "Nokia", "Airbus",
    "TSystems", "Freenet", "SUSE", "TeamViewer"
]

RECRUITEE_BOARDS = [
    "deepL", "personio", "taxfix", "sennder",
    "forto", "moss", "ada", "contentful",
    "pitch", "quantilope", "frontastic", "orderbird",
    "staffbase", "kenjo", "heycar", "urban-sports-club",
    "everdrop", "billie", "finoa", "plan-a"
]

PERSONIO_BOARDS = [
    "personio", "taxfix", "forto", "sennder",
    "staffbase", "moss", "orderbird", "kenjo",
    "finoa", "billie", "everdrop", "heycar",
    "urban-sports-club", "quantilope", "ada"
]

WORKABLE_BOARDS = [
    "n26", "contentful", "mollie", "pleo",
    "qonto", "aircall", "factorial", "payfit",
    "algolia", "typeform", "backmarket", "doctolib",
    "safetywing", "gocardless", "monzo"
]

BREEZY_BOARDS = [
    "techstars", "testcompany"
]


# ============================================================
# GERMAN STATES AND TECHNOLOGY TERMS
# ============================================================

GERMAN_STATES = [
    "Baden-Württemberg", "Bavaria", "Bayern", "Berlin",
    "Brandenburg", "Bremen", "Hamburg", "Hesse", "Hessen",
    "Mecklenburg-Vorpommern", "Lower Saxony",
    "Niedersachsen", "North Rhine-Westphalia",
    "Nordrhein-Westfalen", "Rhineland-Palatinate",
    "Rheinland-Pfalz", "Saarland", "Saxony", "Sachsen",
    "Saxony-Anhalt", "Sachsen-Anhalt", "Schleswig-Holstein",
    "Thuringia", "Thüringen"
]

# All 16 states, used for broad BA searches.
BA_STATES = [
    "Baden-Württemberg", "Bayern", "Berlin", "Brandenburg",
    "Bremen", "Hamburg", "Hessen",
    "Mecklenburg-Vorpommern", "Niedersachsen",
    "Nordrhein-Westfalen", "Rheinland-Pfalz", "Saarland",
    "Sachsen", "Sachsen-Anhalt", "Schleswig-Holstein",
    "Thüringen"
]

TECH_KEYWORDS = [
    "software", "developer", "entwickler", "entwicklerin",
    "softwareentwickler", "softwareentwicklerin",
    "engineer", "ingenieur", "informatik", "informatiker",
    "programmierer", "programmierung",
    "frontend", "front-end", "backend", "back-end",
    "full stack", "fullstack", "web development",
    "react", "angular", "vue", "javascript", "typescript",
    "node.js", "nodejs", "python", "java", "golang",
    "c++", "c#", ".net", "php", "ruby", "swift", "kotlin",
    "rust", "scala", "spring boot", "django", "flask",
    "laravel", "next.js", "mobile developer",
    "app developer", "ios developer", "android developer",
    "data analyst", "datenanalyst", "data scientist",
    "data engineer", "datenbank", "data science",
    "business intelligence", "power bi", "tableau",
    "data warehouse", "sql developer", "etl",
    "artificial intelligence", "künstliche intelligenz",
    "machine learning", "ki engineer", "ai engineer",
    "ml engineer", "deep learning", "computer vision",
    "nlp", "mlops", "generative ai", "llm",
    "devops", "cloud engineer", "cloud architect",
    "aws", "azure", "google cloud", "kubernetes",
    "docker", "terraform", "site reliability",
    "platform engineer", "infrastructure engineer",
    "cybersecurity", "cyber security",
    "it security", "security engineer",
    "information security", "penetration tester",
    "soc analyst", "system administrator",
    "systemadministrator", "netzwerkadministrator",
    "network engineer", "it administrator",
    "it support engineer", "it specialist",
    "qa engineer", "quality assurance",
    "test automation", "automation engineer",
    "software tester", "sdet", "test engineer",
    "sap", "abap", "salesforce", "servicenow",
    "erp developer", "dynamics 365",
    "embedded software", "embedded systems",
    "firmware engineer", "robotics engineer",
    "blockchain developer", "solutions architect",
    "technical architect", "hardware engineer",
    "electronics engineer", "ui developer", "ux engineer",
    "werkstudent informatik", "praktikum software",
    "it praktikum", "software internship"
]


# ============================================================
# HELPERS
# ============================================================

def log(message):
    print(message)
    logs.append({
        "Time": datetime.now(timezone.utc).isoformat(),
        "Message": str(message)
    })


def clean(value):
    if value is None:
        return ""

    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False)

    text = html.unescape(str(value))
    soup = BeautifulSoup(text, "html.parser")
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True)).strip()


def request_json(url, params=None, headers=None):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(
                url,
                params=params,
                headers=headers,
                timeout=TIMEOUT
            )

            if response.status_code in (403, 404, 410):
                return None

            if response.status_code == 429:
                log(f"Rate limited: {url}")
                time.sleep(10 * attempt)
                continue

            response.raise_for_status()
            return response.json()

        except Exception as error:
            log(f"Request error {attempt}/{MAX_RETRIES}: "
                f"{url} | {str(error)[:120]}")
            if attempt < MAX_RETRIES:
                time.sleep(attempt * 2)

    return None


def request_response(url, params=None):
    try:
        response = session.get(
            url,
            params=params,
            timeout=TIMEOUT
        )
        response.raise_for_status()
        return response
    except Exception as error:
        log(f"Request error: {url} | {str(error)[:120]}")
        return None


def parse_date(value):
    if not value:
        return None

    try:
        if isinstance(value, (int, float)):
            if value > 100000000000:
                value = value / 1000
            return datetime.fromtimestamp(value, timezone.utc)

        result = pd.to_datetime(value, errors="coerce", utc=True)

        if pd.isna(result):
            return None

        return result.to_pydatetime()

    except Exception:
        return None


def is_tech(title, description=""):
    text = f"{title} {description}".lower()
    return any(word in text for word in TECH_KEYWORDS)


def germany_match(location, description=""):
    text = f"{clean(location)} {clean(description)}".lower()

    country_words = [
        "germany", "deutschland", "bundesrepublik deutschland"
    ]

    if any(word in text for word in country_words):
        return True

    if any(state.lower() in text for state in GERMAN_STATES):
        return True

    # Common German city names, including smaller cities.
    cities = [
        "berlin", "potsdam", "munich", "münchen", "hamburg",
        "frankfurt", "köln", "cologne", "düsseldorf",
        "stuttgart", "leipzig", "dresden", "hannover",
        "nürnberg", "bremen", "bonn", "dortmund", "essen",
        "mannheim", "karlsruhe", "freiburg", "mainz",
        "wiesbaden", "münster", "aachen", "heidelberg",
        "darmstadt", "augsburg", "regensburg", "ulm",
        "würzburg", "jena", "erfurt", "rostock", "kiel",
        "lübeck", "magdeburg", "saarbrücken", "chemnitz",
        "oldenburg", "braunschweig", "ingolstadt",
        "göttingen", "tübingen", "heilbronn", "paderborn",
        "osnabrück", "bochum", "wuppertal", "bielefeld",
        "duisburg", "kassel", "trier", "passau", "flensburg",
        "leverkusen", "ludwigshafen", "offenbach", "eschborn"
    ]

    if any(city in text for city in cities):
        return True

    # German postal code plus a location field.
    if re.search(r"\b\d{5}\b", clean(location)):
        return True

    return False


def get_city(location):
    text = clean(location).lower()

    cities = [
        "Berlin", "Potsdam", "München", "Munich", "Hamburg",
        "Frankfurt", "Köln", "Düsseldorf", "Stuttgart",
        "Leipzig", "Dresden", "Hannover", "Nürnberg",
        "Bremen", "Bonn", "Dortmund", "Essen", "Mannheim",
        "Karlsruhe", "Freiburg", "Mainz", "Wiesbaden",
        "Münster", "Aachen", "Heidelberg", "Darmstadt",
        "Augsburg", "Regensburg", "Ulm", "Würzburg",
        "Jena", "Erfurt", "Rostock", "Kiel", "Lübeck",
        "Magdeburg", "Saarbrücken", "Chemnitz", "Oldenburg",
        "Braunschweig", "Ingolstadt", "Göttingen", "Tübingen",
        "Heilbronn", "Paderborn", "Osnabrück", "Bochum",
        "Wuppertal", "Bielefeld", "Duisburg", "Kassel",
        "Trier", "Passau", "Flensburg", "Leverkusen",
        "Ludwigshafen", "Offenbach", "Eschborn"
    ]

    for city in sorted(cities, key=len, reverse=True):
        if city.lower() in text:
            return city

    if "remote" in text.lower() or "home office" in text.lower():
        return "Remote"

    return "Other / Not Specified"


def category(title, description=""):
    text = f"{title} {description}".lower()

    groups = {
        "Software Development": [
            "software", "developer", "entwickler", "frontend",
            "backend", "full stack", "programmer", "react"
        ],
        "Data & Analytics": [
            "data analyst", "data scientist", "data engineer",
            "analytics", "power bi", "tableau", "sql"
        ],
        "AI & Machine Learning": [
            "artificial intelligence", "machine learning",
            "ai engineer", "ml engineer", "deep learning",
            "generative ai", "llm"
        ],
        "Cloud & DevOps": [
            "cloud", "devops", "kubernetes", "terraform",
            "site reliability", "platform engineer"
        ],
        "Cybersecurity": [
            "cybersecurity", "cyber security", "security engineer",
            "information security", "penetration tester"
        ],
        "IT & Networking": [
            "system administrator", "systemadministrator",
            "network engineer", "it administrator", "it support"
        ],
        "QA & Testing": [
            "qa engineer", "quality assurance",
            "test automation", "software tester"
        ],
        "Enterprise Technology": [
            "sap", "salesforce", "erp", "servicenow"
        ],
        "Embedded & Hardware": [
            "embedded", "firmware", "robotics",
            "hardware engineer", "electronics engineer"
        ]
    }

    for name, words in groups.items():
        if any(word in text for word in words):
            return name

    return "Other Technology"


def seniority(title):
    text = title.lower()

    if any(x in text for x in [
        "intern", "internship", "working student",
        "werkstudent", "praktikum", "praktikant"
    ]):
        return "Intern / Working Student"

    if any(x in text for x in [
        "graduate", "entry level", "entry-level",
        "junior", "associate", "trainee"
    ]):
        return "Junior / Graduate"

    if any(x in text for x in [
        "principal", "staff engineer", "distinguished",
        "fellow", "architect"
    ]):
        return "Principal / Architect"

    if any(x in text for x in [
        "director", "head of", "vice president",
        "chief", "cto", "cio"
    ]):
        return "Director / Executive"

    if any(x in text for x in [
        "lead", "team lead", "tech lead",
        "engineering manager"
    ]):
        return "Lead / Manager"

    if any(x in text for x in ["senior", "sr. ", "sr "]):
        return "Senior"

    if any(x in text for x in [
        "mid-level", "mid level", "intermediate"
    ]):
        return "Mid-level"

    return "Unspecified / Other"


def normalize_url(url):
    url = clean(url)

    if not url:
        return ""

    try:
        parsed = urlparse(url)
        ignored = {
            "utm_source", "utm_medium", "utm_campaign",
            "utm_term", "utm_content", "ref", "source"
        }

        query = [
            (k, v) for k, v in parse_qsl(parsed.query)
            if k.lower() not in ignored
        ]

        return urlunparse((
            parsed.scheme.lower(),
            parsed.netloc.lower(),
            parsed.path.rstrip("/"),
            "",
            urlencode(query),
            ""
        ))

    except Exception:
        return url.lower().split("?")[0].rstrip("/")


def add_job(
    title,
    company,
    location,
    description,
    url,
    source,
    date_value=None,
    date_basis="",
    salary="",
    employment="",
    work_model=""
):
    source = str(source)
    stats[source]["Records_Received"] += 1

    title = clean(title)
    company = clean(company) or "Company not specified"
    location = clean(location)
    description = clean(description)
    url = clean(url)

    if not title:
        stats[source]["Missing_Title"] += 1
        return

    if not is_tech(title, description):
        stats[source]["Not_Technology"] += 1
        return

    if not germany_match(location, description):
        stats[source]["Not_Germany"] += 1
        return

    job_date = parse_date(date_value)

    if job_date is None:
        date_status = "Unknown"
        stats[source]["Date_Unknown"] += 1
    elif START_DATE <= job_date < END_DATE:
        date_status = "2026"
    else:
        stats[source]["Outside_2026"] += 1
        return

    # Fallback key avoids treating every job without a URL as identical.
    normalized = normalize_url(url)

    if normalized:
        unique_key = normalized
    else:
        unique_key = "|".join([
            title.lower(), company.lower(), location.lower(),
            job_date.date().isoformat() if job_date else ""
        ])

    if unique_key in seen:
        stats[source]["Duplicate"] += 1
        return

    seen.add(unique_key)

    jobs.append({
        "Job_Title": title,
        "Company": company,
        "City": get_city(location),
        "Original_Location": location,
        "Category": category(title, description),
        "Seniority": seniority(title),
        "Employment_Type": clean(employment),
        "Work_Model": clean(work_model),
        "Salary": clean(salary),
        "Job_Date": job_date.isoformat() if job_date else "",
        "Date_Basis": date_basis,
        "Date_Status": date_status,
        "Description": description,
        "Job_URL": url,
        "Source": source,
        "Collected_At": datetime.now(timezone.utc).isoformat()
    })

    stats[source]["Accepted"] += 1


# ============================================================
# 1. BUNDESAGENTUR FÜR ARBEIT
#
# Main German national job database.
# The API is community-documented, not officially supported.
# Searches each German state using the Informatik occupation
# field, then additional broad job-title searches.
# ============================================================

def bundesagentur():
    source = "Bundesagentur für Arbeit"
    log("\n[1] Bundesagentur für Arbeit")

    endpoint = (
        "https://rest.arbeitsagentur.de/"
        "jobboerse/jobsuche-service/pc/v6/jobs"
    )

    headers = {"X-API-Key": "jobboerse-jobsuche"}

    # State searches are broad; title searches catch English roles.
    searches = [
        {"wo": state, "berufsfeld": "Informatik"}
        for state in BA_STATES
    ]

    searches += [
        {"wo": "Deutschland", "was": term}
        for term in [
            "Softwareentwickler",
            "Software Engineer",
            "Data Engineer",
            "Data Scientist",
            "IT Administrator",
            "DevOps Engineer",
            "Cloud Engineer",
            "IT Security",
            "Frontend Developer",
            "Backend Developer",
            "Systemadministrator",
            "Werkstudent Informatik",
            "Praktikum Softwareentwicklung",
            "KI Entwickler",
            "Netzwerkadministrator"
        ]
    ]

    for search_number, base_params in enumerate(searches, 1):
        page = 1
        total_pages = 1

        while page <= total_pages and page <= 30:
            params = {
                **base_params,
                "page": page,
                "size": 100,
                "angebotsart": 1,
                "zeitarbeit": "true",
                "pav": "false"
            }

            data = request_json(
                endpoint,
                params=params,
                headers=headers
            )

            if not data:
                break

            records = data.get("stellenangebote", [])

            if not records:
                break

            total_pages = int(
                data.get("maxErgebnis", 0) or
                data.get("maxSeiten", 0) or
                total_pages
            )

            if total_pages == 1 and len(records) == 100:
                total_pages = page + 1

            for job in records:
                place = job.get("arbeitsort", {}) or {}

                if isinstance(place, dict):
                    location = ", ".join(
                        str(place.get(k, ""))
                        for k in ["ort", "region", "land"]
                        if place.get(k)
                    )
                else:
                    location = str(place)

                ref = (
                    job.get("refnr")
                    or job.get("referenznummer")
                    or ""
                )

                url = (
                    "https://www.arbeitsagentur.de/jobsuche/jobdetail/"
                    + str(ref)
                    if ref else ""
                )

                add_job(
                    title=job.get("beruf", ""),
                    company=job.get("arbeitgeber", ""),
                    location=location,
                    description="",
                    url=url,
                    source=source,
                    date_value=job.get(
                        "aktuelleVeroeffentlichungsdatum"
                    ),
                    date_basis="aktuelleVeroeffentlichungsdatum"
                )

            log(
                f"BA search {search_number}/{len(searches)}, "
                f"page {page}: {len(records)} records"
            )

            page += 1
            time.sleep(REQUEST_DELAY)

    log("Bundesagentur search completed")


# ============================================================
# 2. ARBEITNOW
# ============================================================

def arbeitnow():
    log("\n[2] Arbeitnow")

    url = "https://www.arbeitnow.com/api/job-board-api"
    page = 0
    next_url = url

    while next_url:
        data = request_json(next_url)

        if not data:
            break

        records = data.get("data", [])

        if not records:
            break

        page += 1

        for job in records:
            add_job(
                title=job.get("title"),
                company=job.get("company_name"),
                location=job.get("location"),
                description=job.get("description"),
                url=job.get("url"),
                source="Arbeitnow",
                date_value=job.get("created_at"),
                date_basis="created_at",
                employment=", ".join(job.get("job_types", [])),
                work_model="Remote" if job.get("remote") else ""
            )

        log(f"Arbeitnow page {page}: {len(records)}")
        next_url = data.get("links", {}).get("next")
        time.sleep(REQUEST_DELAY)


# ============================================================
# 3. ADZUNA - OPTIONAL, REQUIRES FREE DEVELOPER CREDENTIALS
#
# Set environment variables before running:
# export ADZUNA_APP_ID="your_app_id"
# export ADZUNA_APP_KEY="your_app_key"
# ============================================================

def adzuna():
    app_id = os.getenv("ADZUNA_APP_ID")
    app_key = os.getenv("ADZUNA_APP_KEY")

    if not app_id or not app_key:
        log("\n[3] Adzuna skipped: API credentials not configured")
        return

    log("\n[3] Adzuna")

    terms = [
        "software developer", "software engineer",
        "data scientist", "data engineer", "IT administrator",
        "DevOps", "cloud engineer", "cyber security",
        "frontend developer", "backend developer",
        "machine learning", "SAP", "Werkstudent IT",
        "IT Praktikum"
    ]

    for term in terms:
        for page in range(1, 6):
            url = (
                f"https://api.adzuna.com/v1/api/jobs/de/search/{page}"
            )

            params = {
                "app_id": app_id,
                "app_key": app_key,
                "what": term,
                "results_per_page": 50,
                "content-type": "application/json"
            }

            data = request_json(url, params=params)

            if not data:
                break

            records = data.get("results", [])

            if not records:
                break

            for job in records:
                location = job.get("location", {})
                location = location.get("display_name", "") if isinstance(
                    location, dict
                ) else location

                add_job(
                    title=job.get("title"),
                    company=job.get("company", {}).get("display_name", ""),
                    location=location,
                    description=job.get("description", ""),
                    url=job.get("redirect_url", ""),
                    source="Adzuna",
                    date_value=job.get("created"),
                    date_basis="created",
                    salary=(
                        f'{job.get("salary_min", "")} - '
                        f'{job.get("salary_max", "")}'
                    )
                )

            log(f"Adzuna '{term}' page {page}: {len(records)}")
            time.sleep(REQUEST_DELAY)


# ============================================================
# 4. JOBICY - REMOTE JOBS
# ============================================================

def jobicy():
    log("\n[4] Jobicy")

    url = "https://jobicy.com/api/jobs"

    for industry in ["engineering", "data-science", "devops"]:
        params = {
            "count": 200,
            "geo": "germany",
            "industry": industry
        }

        data = request_json(url, params=params)

        if not data:
            continue

        records = data.get("jobs", [])

        for job in records:
            location = (
                job.get("jobGeo")
                or job.get("jobGeoCountry")
                or ""
            )

            add_job(
                title=job.get("jobTitle"),
                company=job.get("companyName"),
                location=location,
                description=job.get("jobDescription", ""),
                url=job.get("url"),
                source="Jobicy",
                date_value=job.get("pubDate"),
                date_basis="pubDate",
                employment=job.get("jobType", ""),
                work_model="Remote"
            )

        log(f"Jobicy {industry}: {len(records)}")
        time.sleep(2)


# ============================================================
# 5. REMOTEOK - INCLUDE ONLY GERMANY-ELIGIBLE LISTINGS
# ============================================================

def remoteok():
    log("\n[5] RemoteOK")

    data = request_json("https://remoteok.com/api")

    if not isinstance(data, list):
        return

    for job in data:
        if not isinstance(job, dict) or not job.get("position"):
            continue

        location = job.get("location", "")
        description = job.get("description", "")
        tags = " ".join(job.get("tags", []))

        # Do not treat worldwide remote jobs as Germany jobs
        # unless the listing explicitly mentions Germany.
        combined = f"{location} {description} {tags}".lower()

        if not any(x in combined for x in [
            "germany", "deutschland", "berlin", "munich",
            "münchen", "hamburg", "europe", "eu timezone"
        ]):
            continue

        add_job(
            title=job.get("position"),
            company=job.get("company"),
            location=location,
            description=description,
            url=job.get("url"),
            source="RemoteOK",
            date_value=job.get("date"),
            date_basis="date",
            salary=job.get("salary", ""),
            work_model="Remote"
        )

    log(f"RemoteOK received: {len(data)}")


# ============================================================
# 6. GREENHOUSE
# ============================================================

def greenhouse():
    log("\n[6] Greenhouse")

    for board in GREENHOUSE_BOARDS:
        url = f"https://boards-api.greenhouse.io/v1/boards/{board}/jobs"

        data = request_json(url, params={"content": "true"})

        if not data:
            tested_boards.append({
                "ATS": "Greenhouse",
                "Board": board,
                "Status": "Not found / unavailable",
                "Jobs": 0
            })
            continue

        records = data.get("jobs", [])

        company_data = request_json(
            f"https://boards-api.greenhouse.io/v1/boards/{board}"
        )

        company = (
            company_data.get("name", board)
            if company_data else board
        )

        for job in records:
            loc = job.get("location", {})
            loc = loc.get("name", "") if isinstance(loc, dict) else loc

            add_job(
                title=job.get("title"),
                company=company,
                location=loc,
                description=job.get("content", ""),
                url=job.get("absolute_url"),
                source="Greenhouse",
                date_value=job.get("updated_at"),
                date_basis="updated_at (not original posting date)"
            )

        tested_boards.append({
            "ATS": "Greenhouse",
            "Board": board,
            "Status": "Found",
            "Jobs": len(records)
        })

        log(f"Greenhouse {board}: {len(records)}")
        time.sleep(REQUEST_DELAY)


# ============================================================
# 7. LEVER
# ============================================================

def lever():
    log("\n[7] Lever")

    for board in LEVER_BOARDS:
        found = False

        for host in ["https://api.lever.co", "https://api.eu.lever.co"]:
            records = request_json(
                f"{host}/v0/postings/{board}",
                params={"mode": "json", "limit": 100}
            )

            if not isinstance(records, list):
                continue

            found = True

            for job in records:
                categories = job.get("categories", {}) or {}

                add_job(
                    title=job.get("text"),
                    company=board,
                    location=categories.get("location", ""),
                    description=job.get("descriptionPlain", ""),
                    url=job.get("hostedUrl"),
                    source="Lever",
                    date_value=job.get("createdAt"),
                    date_basis="createdAt",
                    employment=categories.get("commitment", ""),
                    work_model=job.get("workplaceType", "")
                )

            log(f"Lever {board}: {len(records)}")
            break

        tested_boards.append({
            "ATS": "Lever",
            "Board": board,
            "Status": "Found" if found else "Not found / unavailable",
            "Jobs": len(records) if found else 0
        })

        time.sleep(REQUEST_DELAY)


# ============================================================
# 8. ASHBY
# ============================================================

def ashby():
    log("\n[8] Ashby")

    for board in ASHBY_BOARDS:
        data = request_json(
            f"https://api.ashbyhq.com/posting-api/job-board/{board}",
            params={"includeCompensation": "true"}
        )

        if not data:
            tested_boards.append({
                "ATS": "Ashby",
                "Board": board,
                "Status": "Not found / unavailable",
                "Jobs": 0
            })
            continue

        records = data.get("jobs", [])

        for job in records:
            compensation = job.get("compensation", {}) or {}
            salary = (
                compensation.get(
                    "scrapeableCompensationSalarySummary", ""
                )
                if isinstance(compensation, dict) else ""
            )

            add_job(
                title=job.get("title"),
                company=board,
                location=job.get("location", ""),
                description=job.get("descriptionPlain", ""),
                url=job.get("jobUrl") or job.get("applyUrl"),
                source="Ashby",
                date_value=job.get("publishedAt"),
                date_basis="publishedAt",
                salary=salary,
                employment=job.get("employmentType", ""),
                work_model=job.get("workplaceType", "")
            )

        tested_boards.append({
            "ATS": "Ashby",
            "Board": board,
            "Status": "Found",
            "Jobs": len(records)
        })

        log(f"Ashby {board}: {len(records)}")
        time.sleep(REQUEST_DELAY)


# ============================================================
# 9. SMARTRECRUITERS
# ============================================================

def smartrecruiters():
    log("\n[9] SmartRecruiters")

    for board in SMARTRECRUITERS_BOARDS:
        offset = 0
        total = 0

        while True:
            data = request_json(
                f"https://api.smartrecruiters.com/v1/companies/{board}/postings",
                params={
                    "limit": 100,
                    "offset": offset,
                    "destination": "PUBLIC"
                }
            )

            if not data:
                break

            records = data.get("content", [])

            if not records:
                break

            for job in records:
                loc = job.get("location", {}) or {}

                location = ", ".join(
                    str(loc.get(k, ""))
                    for k in ["city", "region", "country"]
                    if loc.get(k)
                )

                job_id = job.get("id", "")
                job_url = (
                    job.get("ref")
                    or (
                        f"https://careers.smartrecruiters.com/"
                        f"{board}/{job_id}"
                    )
                )

                employment = job.get("typeOfEmployment", {})
                employment = (
                    employment.get("label", "")
                    if isinstance(employment, dict)
                    else employment
                )

                add_job(
                    title=job.get("name"),
                    company=board,
                    location=location,
                    description="",
                    url=job_url,
                    source="SmartRecruiters",
                    date_value=job.get("releasedDate"),
                    date_basis="releasedDate",
                    employment=employment
                )

            total += len(records)
            offset += len(records)

            if offset >= data.get("totalFound", 0):
                break

            time.sleep(REQUEST_DELAY)

        tested_boards.append({
            "ATS": "SmartRecruiters",
            "Board": board,
            "Status": "Found" if total else "Not found / empty",
            "Jobs": total
        })

        log(f"SmartRecruiters {board}: {total}")


# ============================================================
# 10. RECRUITEE
# ============================================================

def recruitee():
    log("\n[10] Recruitee")

    for board in RECRUITEE_BOARDS:
        data = request_json(
            f"https://{board}.recruitee.com/api/offers/"
        )

        if not data:
            tested_boards.append({
                "ATS": "Recruitee",
                "Board": board,
                "Status": "Not found / unavailable",
                "Jobs": 0
            })
            continue

        records = data.get("offers", [])

        for job in records:
            add_job(
                title=job.get("title"),
                company=job.get("company_name", board),
                location=job.get("location", ""),
                description=job.get("description", ""),
                url=job.get("careers_url") or job.get("url"),
                source="Recruitee",
                date_value=job.get("published_at"),
                date_basis="published_at",
                employment=job.get("employment_type_code", "")
            )

        tested_boards.append({
            "ATS": "Recruitee",
            "Board": board,
            "Status": "Found",
            "Jobs": len(records)
        })

        log(f"Recruitee {board}: {len(records)}")
        time.sleep(REQUEST_DELAY)


# ============================================================
# 11. PERSONIO
# ============================================================

def personio():
    log("\n[11] Personio")

    for board in PERSONIO_BOARDS:
        response = request_response(
            f"https://{board}.jobs.personio.de/xml",
            params={"language": "en"}
        )

        if not response:
            tested_boards.append({
                "ATS": "Personio",
                "Board": board,
                "Status": "Not found / unavailable",
                "Jobs": 0
            })
            continue

        soup = BeautifulSoup(response.content, "xml")
        records = soup.find_all("position")

        for job in records:
            def field(name):
                item = job.find(name)
                return item.get_text(" ", strip=True) if item else ""

            title = field("name")
            location = field("office")
            description = field("jobDescriptions")

            job_id = field("id")
            url = (
                f"https://{board}.jobs.personio.de/job/{job_id}"
                if job_id else f"https://{board}.jobs.personio.de/"
            )

            add_job(
                title=title,
                company=board,
                location=location,
                description=description,
                url=url,
                source="Personio",
                date_value=None,
                date_basis="No published date in XML feed"
            )

        tested_boards.append({
            "ATS": "Personio",
            "Board": board,
            "Status": "Found",
            "Jobs": len(records)
        })

        log(f"Personio {board}: {len(records)}")
        time.sleep(REQUEST_DELAY)


# ============================================================
# 12. WORKABLE
# ============================================================

def workable():
    log("\n[12] Workable")

    for board in WORKABLE_BOARDS:
        data = request_json(
            f"https://apply.workable.com/api/v1/widget/accounts/{board}",
            params={"details": "true"}
        )

        if not data:
            tested_boards.append({
                "ATS": "Workable",
                "Board": board,
                "Status": "Not found / unavailable",
                "Jobs": 0
            })
            continue

        records = (
            data if isinstance(data, list)
            else data.get("jobs", data.get("results", []))
        )

        for job in records:
            location = job.get("location", {})
            if isinstance(location, dict):
                location = location.get("location_str", str(location))

            add_job(
                title=job.get("title"),
                company=job.get("company", board),
                location=location,
                description=job.get("description", ""),
                url=job.get("url", ""),
                source="Workable",
                date_value=job.get("published"),
                date_basis="published"
            )

        tested_boards.append({
            "ATS": "Workable",
            "Board": board,
            "Status": "Found",
            "Jobs": len(records)
        })

        log(f"Workable {board}: {len(records)}")
        time.sleep(REQUEST_DELAY)


# ============================================================
# 13. BREEZY HR
# ============================================================

def breezy():
    log("\n[13] Breezy HR")

    for board in BREEZY_BOARDS:
        data = request_json(f"https://{board}.breezy.hr/json")

        if not isinstance(data, list):
            tested_boards.append({
                "ATS": "Breezy HR",
                "Board": board,
                "Status": "Not found / unavailable",
                "Jobs": 0
            })
            continue

        for job in data:
            location = job.get("location", {})
            if isinstance(location, dict):
                location = ", ".join(
                    str(v) for v in location.values() if v
                )

            add_job(
                title=job.get("name"),
                company=board,
                location=location,
                description=job.get("description", ""),
                url=job.get("url"),
                source="Breezy HR",
                date_value=job.get("published_date"),
                date_basis="published_date"
            )

        tested_boards.append({
            "ATS": "Breezy HR",
            "Board": board,
            "Status": "Found",
            "Jobs": len(data)
        })

        log(f"Breezy {board}: {len(data)}")
        time.sleep(REQUEST_DELAY)


# ============================================================
# EXCEL EXPORT
# ============================================================

def export_files():
    if not jobs:
        log("No matching jobs found. Check Collection_Log.")
        return

    df = pd.DataFrame(jobs)

    df_2026 = df[df["Date_Status"] == "2026"].copy()
    df_unknown = df[df["Date_Status"] == "Unknown"].copy()

    if not df_2026.empty:
        df_2026["_sort"] = pd.to_datetime(
            df_2026["Job_Date"], errors="coerce", utc=True
        )
        df_2026 = df_2026.sort_values(
            "_sort", ascending=False
        ).drop(columns=["_sort"])

    def summary(data, column):
        if data.empty:
            return pd.DataFrame(columns=[column, "Total_Jobs"])

        return (
            data.groupby(column)
            .size()
            .reset_index(name="Total_Jobs")
            .sort_values("Total_Jobs", ascending=False)
        )

    stats_df = pd.DataFrame([
        {"Source": source, **dict(values)}
        for source, values in stats.items()
    ]).fillna(0)

    boards_df = pd.DataFrame(tested_boards)

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        df_2026.to_excel(
            writer, sheet_name="2026_Tech_Jobs", index=False
        )
        df_unknown.to_excel(
            writer, sheet_name="Date_Unknown", index=False
        )
        summary(df, "Company").to_excel(
            writer, sheet_name="Company_Summary", index=False
        )
        summary(df, "City").to_excel(
            writer, sheet_name="City_Summary", index=False
        )
        summary(df, "Category").to_excel(
            writer, sheet_name="Category_Summary", index=False
        )
        summary(df, "Seniority").to_excel(
            writer, sheet_name="Seniority_Summary", index=False
        )
        summary(df, "Source").to_excel(
            writer, sheet_name="Source_Summary", index=False
        )
        stats_df.to_excel(
            writer, sheet_name="Collection_Stats", index=False
        )
        boards_df.to_excel(
            writer, sheet_name="Company_Board_Test", index=False
        )
        pd.DataFrame(logs).to_excel(
            writer, sheet_name="Collection_Log", index=False
        )

        for sheet in writer.sheets.values():
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

            for column in sheet.columns:
                letter = column[0].column_letter
                width = max(
                    len(str(cell.value or ""))
                    for cell in column
                )
                sheet.column_dimensions[letter].width = min(
                    width + 2, 55
                )

            if sheet.title in ["2026_Tech_Jobs", "Date_Unknown"]:
                headers = {
                    cell.value: cell.column
                    for cell in sheet[1]
                }
                url_col = headers.get("Job_URL")

                if url_col:
                    for row in range(2, sheet.max_row + 1):
                        cell = sheet.cell(row, url_col)
                        if cell.value:
                            cell.hyperlink = cell.value
                            cell.style = "Hyperlink"

    df.to_csv(CSV_OUTPUT, index=False, encoding="utf-8-sig")

    print("\n" + "=" * 60)
    print("GERMANY TECH JOB COLLECTION FINISHED")
    print("=" * 60)
    print("All accepted jobs:", len(df))
    print("Confirmed 2026 jobs:", len(df_2026))
    print("Unknown-date jobs:", len(df_unknown))
    print("Unique companies:", df["Company"].nunique())
    print("Location groups:", df["City"].nunique())
    print("Sources with accepted jobs:", df["Source"].nunique())

    print("\nJobs by source:")
    print(summary(df, "Source").to_string(index=False))

    print("\nJobs by seniority:")
    print(summary(df, "Seniority").to_string(index=False))

    print("\nJobs by city:")
    print(summary(df, "City").to_string(index=False))

    print("\nCompany boards found:")
    if not boards_df.empty:
        print(
            boards_df[boards_df["Status"] == "Found"]
            .to_string(index=False)
        )

    print("\nExcel:", OUTPUT)
    print("CSV:", CSV_OUTPUT)


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    collectors = [
        bundesagentur,
        arbeitnow,
        adzuna,
        jobicy,
        remoteok,
        greenhouse,
        lever,
        ashby,
        smartrecruiters,
        recruitee,
        personio,
        workable,
        breezy
    ]

    for collector in collectors:
        try:
            collector()
        except KeyboardInterrupt:
            log("Stopped by user.")
            break
        except Exception as error:
            log(
                f"COLLECTOR FAILED: {collector.__name__}: "
                f"{str(error)[:300]}"
            )

    export_files()
