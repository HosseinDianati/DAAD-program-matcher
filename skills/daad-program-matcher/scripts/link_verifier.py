"""
Link and Deadline Verifier for German University Programs.
Executes Phase B verification:
1. Validates program links and detects generic homepage redirects.
2. Formats targeted web search queries to locate faculty/department admission regulations.
3. Distinguishes program-specific non-EU deadlines from general nationwide windows.
"""

import requests
import re
from urllib.parse import urlparse

GENERIC_DOMAINS = [
    "daad.de",
    "uni-assist.de",
    "hochschulstart.de",
    "google.com"
]

def check_url_health(url: str, timeout: int = 5) -> dict:
    """
    Checks if a URL is reachable, resolves redirects, and flags generic homepages.
    """
    if not url or not str(url).startswith("http"):
        return {"status": "invalid", "url": url, "is_specific": False}
    
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        res = requests.head(url, headers=headers, allow_redirects=True, timeout=timeout)
        final_url = res.url
        parsed = urlparse(final_url)
        path = parsed.path.strip("/").split("/")
        
        # If the path is empty or just language code (e.g. /en or /de), it's a generic homepage redirect
        is_generic = len(path) == 0 or (len(path) == 1 and path[0] in ["en", "de", "home", "index.html"])
        
        return {
            "status": "valid" if res.status_code < 400 else f"http_{res.status_code}",
            "final_url": final_url,
            "is_specific": not is_generic
        }
    except Exception as e:
        return {"status": "error", "error": str(e), "url": url, "is_specific": False}


def build_verification_search_queries(university: str, program: str) -> list[str]:
    """
    Constructs high-precision search queries for Google / Search API to find
    departmental admission regulations (Zugangssatzung/Bewerbung/Deadlines).
    """
    clean_uni = re.sub(r'\(.*?\)', '', university).strip()
    clean_prog = re.sub(r'\(.*?\)', '', program).strip()
    
    return [
        f'"{clean_uni}" "{clean_prog}" admission requirements deadline international Non-EU',
        f'"{clean_uni}" "{clean_prog}" Master application deadline faculty department',
        f'"{clean_uni}" "{clean_prog}" Zugangssatzung Bewerbungsfrist Master'
    ]


def parse_deadline_nuance(deadline_text: str) -> str:
    """
    Highlights non-EU vs EU differences and warns if only general German window is present.
    """
    text = deadline_text.strip()
    if "15 July" in text and "Non-EU" not in text and "International" not in text:
        return f"{text} (Note: General deadline; verify faculty early cutoff)"
    return text


if __name__ == "__main__":
    test_uni = "Bauhaus-Universität Weimar"
    test_prog = "Digital Technologies in Architecture and Design"
    queries = build_verification_search_queries(test_uni, test_prog)
    print("Verification queries:")
    for q in queries:
        print(" ->", q)
