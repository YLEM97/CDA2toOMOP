import json
import os
import re

import requests
from omophub import OMOPHub


CACHE_FILE = "vocab_cache.json"
OMOPHUB_API_KEY = "YOUR_OMOPHUB_KEY"  # Change with your OMOPHub key
MYMEMORY_URL = "https://api.mymemory.translated.net/get"

MIC_SUFFIX_STANDARD = "[Susceptibility] by Minimum inhibitory concentration (MIC)"
MIC_SUFFIX_HIGH_POT = ".high potency [Susceptibility] by Minimum inhibitory concentration (MIC)"
MIC_SUFFIX_INDUCED = ".induced [Susceptibility] by Minimum inhibitory concentration (MIC)"

# ESBL management.
ESBL_LOINC_CODE = "6984-9"
ESBL_LOINC_NAME = "Beta lactamase.extended spectrum [Susceptibility]"
_PATTERN_ESBL = re.compile(
    r"\besbl\b|extended[\s-]*spectrum\s+beta[\s-]*lactamase",
    re.IGNORECASE,
)

_PATTERN_HIGH_POTENCY = re.compile(
    r"alto\s*(livello|dosaggio)|high\s*(level|potency)|screening\s*resistenza\s*alto",
    re.IGNORECASE,
)
_PATTERN_INDUCED = re.compile(
    r"induc[ib]|inducibil|induced",
    re.IGNORECASE,
)

TRANSLATION_OVERRIDE = {}

try:
    from googletrans import Translator as _GTranslator

    _gtranslator = _GTranslator()
    _USE_GOOGLETRANS = True
except Exception:
    _gtranslator = None
    _USE_GOOGLETRANS = False


def is_esbl_test(display_name: str) -> bool:
    return bool(_PATTERN_ESBL.search(display_name or ""))


def get_mic_suffix(display_name: str) -> str:
    if _PATTERN_HIGH_POTENCY.search(display_name or ""):
        return MIC_SUFFIX_HIGH_POT
    if _PATTERN_INDUCED.search(display_name or ""):
        return MIC_SUFFIX_INDUCED
    return MIC_SUFFIX_STANDARD


def translate_to_english(text: str) -> str:

    text = text or ""

    override = TRANSLATION_OVERRIDE.get(text.strip().lower())
    if override:
        return override

    if _USE_GOOGLETRANS:
        try:
            result = _gtranslator.translate(text, src="it", dest="en")
            translated = result.text.strip()
            if translated and translated.lower() != text.lower():
                return translated
        except Exception as exc:
            print(f"  [WARN googletrans] Error for '{text}': {exc}")

    try:
        resp = requests.get(
            MYMEMORY_URL,
            params={"q": text, "langpair": "it|en"},
            timeout=5,
        )
        data = resp.json()
        translated = data.get("responseData", {}).get("translatedText", "")
        if translated and "MYMEMORY WARNING" not in translated.upper():
            return translated.strip()
    except Exception as exc:
        print(f"  [WARN MyMemory] Error for '{text}': {exc}")

    return text


class VocabLookup:

    def __init__(self, api_key: str = None, cache_file: str = CACHE_FILE):
        key = api_key or os.environ.get("OMOPHUB_API_KEY") or OMOPHUB_API_KEY
        self._client = OMOPHub(api_key=key)
        self._cache_file = cache_file
        self._cache = self._load_cache()

    # Cache persistente
    def _load_cache(self) -> dict:
        if os.path.exists(self._cache_file):
            try:
                with open(self._cache_file, "r", encoding="utf-8") as file_obj:
                    return json.load(file_obj)
            except (OSError, json.JSONDecodeError) as exc:
                print(f"  [WARN CACHE] Cannot read '{self._cache_file}': {exc}")
        return {}

    def _save_cache(self) -> None:
        try:
            with open(self._cache_file, "w", encoding="utf-8") as file_obj:
                json.dump(self._cache, file_obj, indent=2, ensure_ascii=False)
        except OSError as exc:
            print(f"  [WARN CACHE] Cannot write '{self._cache_file}': {exc}")

    @staticmethod
    def _cache_key(category: str, name: str) -> str:
        return f"{category}|{(name or '').strip().lower()}"

    def _search_semantic(
        self,
        name: str,
        vocabulary_ids: list,
        domain_id: str,
        concept_class_id: str = None,
    ) -> int:
        try:
            results = self._client.search.semantic(
                name,
                vocabulary_ids=vocabulary_ids,
                domain_ids=[domain_id],
                concept_class_id=concept_class_id if concept_class_id else None,
                page_size=5,
            )

            if isinstance(results, dict):
                items = results.get("results", []) or results.get("data", []) or []
            elif isinstance(results, list):
                items = results
            else:
                items = getattr(results, "data", []) or []

            if not items:
                return 0

            first = items[0]
            if isinstance(first, dict):
                return int(first.get("concept_id", 0) or 0)
            return int(getattr(first, "concept_id", 0) or 0)

        except Exception as exc:
            print(f"  [WARN OMOPHub semantic] Lookup failed for '{name}': {exc}")
            return 0

    def _search_loinc_exact(self, concept_code: str, concept_name: str) -> int:
        for query in (concept_code, concept_name):
            try:
                results = self._client.search.basic(
                    query,
                    vocabulary_ids=["LOINC"],
                    standard_concept="S",
                    page_size=50,
                )

                if isinstance(results, dict):
                    items = results.get("results", []) or results.get("data", []) or []
                elif isinstance(results, list):
                    items = results
                else:
                    items = getattr(results, "data", []) or []

                for item in items:
                    if not isinstance(item, dict):
                        continue
                    if item.get("domain_id", "").lower() != "measurement":
                        continue

                    result_code = str(item.get("concept_code", "")).strip()
                    result_name = str(item.get("concept_name", "")).strip()
                    if (
                        result_code == concept_code
                        or result_name.lower() == concept_name.lower()
                    ):
                        return int(item.get("concept_id", 0) or 0)

            except Exception as exc:
                print(f"  [WARN OMOPHub exact LOINC] Lookup failed for '{query}': {exc}")

        return 0

    def _lookup_semantic(
        self,
        category: str,
        name: str,
        vocabulary_ids: list,
        domain_id: str,
        concept_class_id: str = None,
        translate: bool = False,
    ) -> int:
        key = self._cache_key(category, name)
        if key in self._cache:
            return self._cache[key]

        search_name = name
        if translate:
            translated = translate_to_english(name)
            if translated.lower() != (name or "").lower():
                print(f"  [TRAD] '{name}' -> '{translated}'")
            search_name = translated

        concept_id = self._search_semantic(
            search_name,
            vocabulary_ids,
            domain_id,
            concept_class_id,
        )

        if concept_id == 0 and translate and search_name != name:
            print(f"  [RETRY]: '{name}'")
            concept_id = self._search_semantic(
                name,
                vocabulary_ids,
                domain_id,
                concept_class_id,
            )

        self._cache[key] = concept_id
        self._save_cache()

        if concept_id == 0:
            print(f"  [WARN UNMAPPED] {category}: '{name}' -> 0")
        else:
            print(f"  [VOCAB] {category}: '{name}' -> {concept_id}")

        return concept_id

    def organism(self, display_name: str) -> int:
        return self._lookup_semantic(
            category="organism",
            name=display_name,
            vocabulary_ids=["SNOMED"],
            domain_id="Observation",
            concept_class_id="Organism",
            translate=False,
        )

    def specimen(self, display_name: str) -> int:
        return self._lookup_semantic(
            category="specimen",
            name=display_name,
            vocabulary_ids=["SNOMED"],
            domain_id="Specimen",
            concept_class_id="Specimen",
            translate=True,
        )

    def antibiotic(self, display_name: str) -> int:

        if is_esbl_test(display_name):
            # Rimuove l'eventuale vecchio mapping errato antibiotic|esbl.
            old_key = self._cache_key("antibiotic", display_name)
            if old_key in self._cache:
                del self._cache[old_key]

            key = self._cache_key("resistance_marker", "ESBL")
            if key in self._cache:
                return self._cache[key]

            print(
                f"  [SPECIAL LOINC] ESBL -> {ESBL_LOINC_CODE} "
                f"'{ESBL_LOINC_NAME}'"
            )
            concept_id = self._search_loinc_exact(
                concept_code=ESBL_LOINC_CODE,
                concept_name=ESBL_LOINC_NAME,
            )

            self._cache[key] = concept_id
            self._save_cache()

            if concept_id == 0:
                print(
                    f"  [WARN UNMAPPED] ESBL: no exact match for LOINC "
                    f"{ESBL_LOINC_CODE}; 0"
                )
            else:
                print(
                    f"  [VOCAB] ESBL: LOINC {ESBL_LOINC_CODE} -> "
                    f"OMOP concept_id {concept_id}"
                )
            return concept_id

        key = self._cache_key("antibiotic", display_name)
        if key in self._cache:
            return self._cache[key]

        suffix = get_mic_suffix(display_name)
        translated = translate_to_english(display_name)
        if translated.lower() != (display_name or "").lower():
            print(f"  [TRAD] '{display_name}' -> '{translated}'")

        query = f"{translated}{suffix}"
        print(f"  [QUERY LOINC] '{query}'")

        concept_id = self._search_semantic(
            name=query,
            vocabulary_ids=["LOINC"],
            domain_id="Measurement",
            concept_class_id="Lab Test",
        )

        self._cache[key] = concept_id
        self._save_cache()

        if concept_id == 0:
            print(f"  [WARN UNMAPPED] antibiotic: '{display_name}' -> 0")
        else:
            print(f"  [VOCAB] antibiotic: '{display_name}' -> {concept_id}")

        return concept_id
