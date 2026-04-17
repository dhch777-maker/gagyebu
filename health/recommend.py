"""건강검진 항목 우선순위 추천기

병원 검진 항목 목록(텍스트 파일)을 입력받아,
과거 건강검진 데이터 + 의학 가이드라인 + 유전체 정보를 종합하여
올해 우선적으로 받아야 할 검사를 추천한다.

가중치 배분 (개인이력 우선):
  Layer 2 개인이력  최대 100점
  Layer 1 가이드라인 최대  70점
  Layer 3 유전체     최대  20점

사용법:
  python recommend.py items.txt
  python recommend.py items.txt --year 2026

근거:
  - USPSTF A/B 등급 권고 (uspreventiveservicestaskforce.org)
  - 국가건강검진 항목/주기 (nhis.or.kr, 2025 기준)
  - 한국지질·동맥경화학회, 대한간학회 등 전문학회 가이드라인
  - ACS 암 선별검사 권고 (cancer.org)
"""
import sys
import os
import json
import re
from datetime import datetime


# ─── 데이터 로드 ────────────────────────────────────────────

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_yearly_data(structured_dir):
    data = []
    for fname in sorted(os.listdir(structured_dir)):
        if fname.endswith(".json") and fname.replace(".json", "").isdigit():
            data.append(load_json(os.path.join(structured_dir, fname)))
    return sorted(data, key=lambda d: d["year"])


def load_items(path):
    """텍스트 파일에서 검진 항목 목록 로드 (한 줄에 하나, 빈줄/주석 무시)"""
    items = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#"):
                items.append(line)
    return items


# ─── 항목 인식 (매칭) ───────────────────────────────────────

def normalize(text):
    """비교용 정규화: 소문자, 공백/특수문자 제거"""
    return re.sub(r"[\s\-_·/()（）]", "", text.lower())


def match_item(item_text, term_aliases):
    """입력 항목 텍스트를 정규화된 카테고리로 매핑.
    반환: 매칭된 카테고리 키 리스트 (예: ["위내시경", "혈액검사"])
    """
    norm_input = normalize(item_text)
    matched = []
    for category, aliases in term_aliases.items():
        for alias in aliases:
            if normalize(alias) in norm_input or norm_input in normalize(alias):
                matched.append(category)
                break
    # 매칭 안 되면 부분 매칭 시도
    if not matched:
        for category, aliases in term_aliases.items():
            for alias in aliases:
                norm_alias = normalize(alias)
                # 3글자 이상 공통 부분 있으면 매칭
                if len(norm_input) >= 2 and len(norm_alias) >= 2:
                    if norm_input[:2] in norm_alias or norm_alias[:2] in norm_input:
                        matched.append(category)
                        break
    return list(set(matched))


# ─── 점수 계산 ──────────────────────────────────────────────

def safe_get(data, key_path):
    """점 구분 키 경로로 중첩 딕셔너리 접근 (예: 'blood.lipid.ldl')"""
    obj = data
    for k in key_path.split("."):
        if not isinstance(obj, dict):
            return None
        obj = obj.get(k)
    return obj


def get_personal_score(category, yearly_data, guidelines):
    """Layer 2: 개인 이력 기반 점수 (최대 100점)"""
    score = 0
    reasons = []

    if not yearly_data:
        return score, reasons

    latest = yearly_data[-1]

    # --- flags 이상 소견 (+35) ---
    all_related_keys = set()
    all_related_terms = set()
    # 가이드라인에서 이 카테고리와 관련된 키/용어 수집
    for section in [guidelines.get("uspstf", {}), guidelines.get("acs", {})]:
        for item in section.get("items", []):
            if any(normalize(t) in normalize(category) or normalize(category) in normalize(t)
                   for t in item.get("related_terms", [])):
                all_related_keys.update(item.get("related_keys", []))
    for item in guidelines.get("korean_society", {}).get("items", []):
        if any(normalize(t) in normalize(category) or normalize(category) in normalize(t)
               for t in item.get("related_terms", [])):
            all_related_keys.update(item.get("related_keys", []))
    for section in guidelines.get("national_checkup", {}).get("general", []):
        if any(normalize(t) in normalize(category) or normalize(category) in normalize(t)
               for t in section.get("related_terms", [])):
            all_related_keys.update(section.get("related_keys", []))
    for section in guidelines.get("national_checkup", {}).get("cancer", []):
        if any(normalize(t) in normalize(category) or normalize(category) in normalize(t)
               for t in section.get("related_terms", [])):
            all_related_keys.update(section.get("related_keys", []))

    # flags 체크
    for yd in yearly_data:
        for flag in yd.get("flags", []):
            flag_item = flag.get("item", "")
            for key in all_related_keys:
                if flag_item in key or key.endswith(flag_item):
                    score += 35
                    status = flag.get("status", "")
                    val = flag.get("value", "")
                    reasons.append(f"이상소견 ({yd['year']}): {flag_item} {val} [{status}]")
                    break
            # 카테고리명 직접 매칭
            cat_terms = guidelines.get("term_aliases", {}).get("mappings", {}).get(category, [])
            for t in cat_terms:
                if normalize(flag_item) in normalize(t) or normalize(t) in normalize(flag_item):
                    if not any(flag_item in r for r in reasons):
                        score += 35
                        reasons.append(f"이상소견 ({yd['year']}): {flag.get('item')} [{flag.get('status')}]")
                    break

    # --- 연속 이상 수치 (+25) ---
    if len(yearly_data) >= 2:
        for key in all_related_keys:
            statuses = []
            for yd in yearly_data:
                obj = safe_get(yd, key)
                if isinstance(obj, dict):
                    statuses.append(obj.get("status"))
            consecutive = all(s in ("높음", "주의", "낮음") for s in statuses[-2:] if s)
            if consecutive and len([s for s in statuses[-2:] if s]) >= 2:
                score += 25
                reasons.append(f"연속 이상: {key.split('.')[-1]} {len(yearly_data[-2:])}년 연속")
                break

    # --- 추적/재검 키워드 (+30) ---
    recheck_keywords = ["재검", "추적", "추적검사", "정기검진", "경과관찰", "f/u", "follow"]
    for yd in yearly_data:
        for flag in yd.get("flags", []):
            action = flag.get("action", "")
            flag_item = flag.get("item", "")
            # 이 카테고리와 관련된 flag인지 확인
            is_related = False
            for key in all_related_keys:
                if flag_item in key or key.endswith(flag_item):
                    is_related = True
                    break
            if not is_related:
                cat_terms = guidelines.get("term_aliases", {}).get("mappings", {}).get(category, [])
                for t in cat_terms:
                    if normalize(flag_item) in normalize(t) or normalize(t) in normalize(flag_item):
                        is_related = True
                        break
            if is_related and any(kw in action for kw in recheck_keywords):
                score += 30
                reasons.append(f"추적권고 ({yd['year']}): \"{action}\"")
                break

    # --- 경계 수치 모니터링 (+15) ---
    # 정상이지만 참고치 상한/하한에 근접한 수치 감지
    BOUNDARY_THRESHOLDS = {
        # key: (상한, 경계 비율) — 값이 상한 * 비율 이상이면 경계
        "blood.diabetes.fasting_glucose": {"upper": 100, "ratio": 0.95},
        "blood.diabetes.hba1c": {"upper": 5.7, "ratio": 0.95},
        "blood.kidney.creatinine": {"upper": 1.2, "ratio": 0.9},
        "blood.kidney.egfr": {"lower": 90, "ratio": 1.05},
        "basic.bp_systolic": {"upper": 120, "ratio": 0.95},
        "basic.bp_diastolic": {"upper": 80, "ratio": 0.95},
        "basic.bmi": {"upper": 24.9, "ratio": 0.97},
    }
    for key in all_related_keys:
        if key not in BOUNDARY_THRESHOLDS:
            continue
        thresh = BOUNDARY_THRESHOLDS[key]
        val_obj = safe_get(latest, key)
        if not isinstance(val_obj, dict) or val_obj.get("value") is None:
            continue
        v = val_obj["value"]
        status = val_obj.get("status", "")
        if status in ("높음", "주의", "낮음"):
            continue  # 이미 이상 — flags에서 처리됨
        if "upper" in thresh and v >= thresh["upper"] * thresh["ratio"]:
            score += 15
            reasons.append(f"경계수치: {key.split('.')[-1]} {v} (상한 {thresh['upper']}의 {thresh['ratio']*100:.0f}% 이상)")
        elif "lower" in thresh and v <= thresh["lower"] * thresh["ratio"]:
            score += 15
            reasons.append(f"경계수치: {key.split('.')[-1]} {v} (하한 {thresh['lower']}의 {thresh['ratio']*100:.0f}% 이하)")

    # --- 미검사 사각지대 (+10) ---
    has_any_data = False
    for yd in yearly_data:
        for key in all_related_keys:
            val = safe_get(yd, key)
            if val is not None:
                has_any_data = True
                break
        if has_any_data:
            break
    if not has_any_data and all_related_keys:
        score += 10
        reasons.append("미검사: 과거 데이터 없음 (사각지대)")

    return min(score, 100), reasons


def get_guideline_score(category, user_age, user_sex, yearly_data, guidelines, target_year):
    """Layer 1: 가이드라인 기반 점수 (최대 70점)"""
    score = 0
    reasons = []
    sex_map = {"남": "male", "여": "female"}
    user_sex_en = sex_map.get(user_sex, "all")

    def age_in_range(item):
        min_a = item.get("min_age", 0)
        max_a = item.get("max_age", 999)
        target_a = item.get("target_age")
        if target_a:
            return user_age == target_a
        return min_a <= user_age <= max_a

    def sex_matches(item):
        s = item.get("sex", "all")
        return s == "all" or s == user_sex_en

    def terms_match(item):
        return any(
            normalize(t) in normalize(category) or normalize(category) in normalize(t)
            for t in item.get("related_terms", [])
        )

    # USPSTF
    for item in guidelines.get("uspstf", {}).get("items", []):
        if terms_match(item) and age_in_range(item) and sex_matches(item):
            grade = item.get("grade", "")
            pts = 30 if grade == "A" else 20
            score += pts
            reasons.append(f"USPSTF {grade}등급: {item['name']}")
            if item.get("note"):
                reasons.append(f"  └ {item['note']}")

    # 국가건강검진
    for section_key in ["general", "cancer"]:
        for item in guidelines.get("national_checkup", {}).get(section_key, []):
            if terms_match(item) and age_in_range(item) and sex_matches(item):
                interval = item.get("interval_years")
                pts = 20
                # 주기 도래 여부 확인
                if interval and yearly_data:
                    last_year = yearly_data[-1]["year"]
                    if (target_year - last_year) >= interval:
                        pts = 25
                        reasons.append(f"국가검진: {item['name']} (주기 {interval}년, 도래)")
                    else:
                        reasons.append(f"국가검진: {item['name']} (주기 {interval}년)")
                else:
                    reasons.append(f"국가검진: {item['name']}")
                score += pts

    # 한국 학회 가이드라인
    for item in guidelines.get("korean_society", {}).get("items", []):
        if terms_match(item):
            # 트리거 조건 확인
            triggered = False
            trigger = item.get("trigger", "")
            for yd in yearly_data:
                for key in item.get("related_keys", []):
                    val_obj = safe_get(yd, key)
                    if isinstance(val_obj, dict):
                        status = val_obj.get("status", "")
                        if status in ("높음", "주의", "낮음"):
                            triggered = True
                            break
                # imaging/endoscopy 소견도 체크
                for key in item.get("related_keys", []):
                    parts = key.split(".")
                    if parts[0] in ("imaging", "endoscopy") and len(parts) >= 2:
                        section = yd.get(parts[0], {})
                        sub = section.get(parts[1], {})
                        if isinstance(sub, dict) and sub.get("status") in ("주의", "높음"):
                            triggered = True
                            break
                # flags의 finding도 체크
                for flag in yd.get("flags", []):
                    for term in item.get("related_terms", []):
                        if normalize(term) in normalize(flag.get("item", "")):
                            triggered = True
                            break
                if triggered:
                    break

            if triggered:
                score += 25
                reasons.append(f"학회권고: {item['name']}")
                if item.get("note"):
                    reasons.append(f"  └ {item['note']}")

    # ACS
    for item in guidelines.get("acs", {}).get("items", []):
        if terms_match(item) and age_in_range(item) and sex_matches(item):
            score += 20
            reasons.append(f"ACS: {item['name']}")
        elif terms_match(item) and not age_in_range(item):
            # 연령 미해당이지만 유전적 위험이 있으면 참고로 언급
            reasons.append(f"ACS 참고: {item['name']} (권고 연령 {item.get('min_age', '?')}세+, 현재 미해당)")

    return min(score, 70), reasons


def get_genetics_score(category, yearly_data, guidelines):
    """Layer 3: 유전체 기반 점수 (최대 20점)"""
    score = 0
    reasons = []

    # 유전체 데이터가 있는 가장 최근 연도 찾기
    genetics = None
    for yd in reversed(yearly_data):
        if yd.get("genetics"):
            genetics = yd["genetics"]
            break
    if not genetics:
        return score, reasons

    # 카테고리 → 유전체 항목 매핑
    gene_item_map = {
        "위내시경": [],
        "대장내시경": ["colorectal"],
        "흉부CT": ["lung"],
        "간초음파": ["liver"],
        "복부초음파": ["liver"],
        "심장초음파": ["myocardial_infarction", "atrial_fibrillation"],
        "심전도": ["myocardial_infarction", "atrial_fibrillation"],
        "동맥경화도": ["myocardial_infarction", "stroke"],
        "혈압": ["myocardial_infarction", "stroke"],
        "지질검사": ["myocardial_infarction"],
        "전립선초음파": ["prostate"],
        "정신건강": ["alzheimers", "parkinsons"],
    }

    related_gene_items = gene_item_map.get(category, [])

    for gen_section in genetics.values():
        if not isinstance(gen_section, dict) or "items" not in gen_section:
            continue
        for item_key, item_val in gen_section.get("items", {}).items():
            if item_key not in related_gene_items:
                continue
            risk = item_val.get("risk_ratio", 0)
            level = item_val.get("level", "")
            if risk > 2.0:
                score += 20
                reasons.append(f"유전체 고위험: {item_key} {risk}x [{level}]")
            elif risk > 1.3:
                score += 10
                reasons.append(f"유전체 주의: {item_key} {risk}x [{level}]")

    return min(score, 20), reasons


# ─── 메인 추천 엔진 ─────────────────────────────────────────

def recommend(items_path, data_dir, target_year=None):
    """검진 항목 우선순위 추천 실행"""
    structured_dir = os.path.join(data_dir, "structured")
    guidelines_path = os.path.join(data_dir, "guidelines.json")

    yearly_data = load_yearly_data(structured_dir)
    guidelines = load_json(guidelines_path)
    input_items = load_items(items_path)
    term_aliases = guidelines.get("term_aliases", {}).get("mappings", {})

    if not yearly_data:
        print("오류: 과거 검진 데이터가 없습니다.")
        return

    latest = yearly_data[-1]
    user_age = latest.get("age", 33)
    user_sex = latest.get("sex", "남")
    if target_year is None:
        target_year = datetime.now().year

    # 각 항목에 대해 3-layer 점수 계산
    results = []
    unmatched = []

    for raw_item in input_items:
        matched_categories = match_item(raw_item, term_aliases)

        if not matched_categories:
            unmatched.append(raw_item)
            continue

        # 매칭된 카테고리 중 최고 점수 채택
        best = None
        for cat in matched_categories:
            personal_score, personal_reasons = get_personal_score(cat, yearly_data, guidelines)
            guideline_score, guideline_reasons = get_guideline_score(
                cat, user_age, user_sex, yearly_data, guidelines, target_year
            )
            genetics_score, genetics_reasons = get_genetics_score(cat, yearly_data, guidelines)

            total = personal_score + guideline_score + genetics_score
            entry = {
                "input": raw_item,
                "category": cat,
                "total": total,
                "personal": {"score": personal_score, "reasons": personal_reasons},
                "guideline": {"score": guideline_score, "reasons": guideline_reasons},
                "genetics": {"score": genetics_score, "reasons": genetics_reasons},
            }
            if best is None or total > best["total"]:
                best = entry

        results.append(best)

    # 점수순 정렬
    results.sort(key=lambda r: r["total"], reverse=True)

    # 출력
    print_results(results, unmatched, user_age, user_sex, target_year, yearly_data)
    return results


def get_grade(total):
    if total >= 60:
        return "[필수]", "★★★"
    elif total >= 40:
        return "[강력권고]", "★★☆"
    elif total >= 20:
        return "[권고]", "★☆☆"
    else:
        return "[선택]", "☆☆☆"


def print_results(results, unmatched, age, sex, year, yearly_data):
    years_str = ", ".join(str(d["year"]) for d in yearly_data)

    print()
    print("=" * 64)
    print(f"  {year}년 건강검진 우선순위 추천 ({age}세 {sex}성)")
    print(f"  과거 데이터: {years_str}")
    print("=" * 64)
    print()

    if not results:
        print("  매칭된 검진 항목이 없습니다.")
        return

    for i, r in enumerate(results, 1):
        grade_label, stars = get_grade(r["total"])
        print(f"  {i}. {grade_label} {stars}  {r['input']}  ─ {r['total']}점")

        # 개인이력 이유
        if r["personal"]["reasons"]:
            print(f"     {'─' * 40}")
            print(f"     개인이력 ({r['personal']['score']}점):")
            for reason in r["personal"]["reasons"]:
                print(f"       · {reason}")

        # 가이드라인 이유
        if r["guideline"]["reasons"]:
            print(f"     가이드라인 ({r['guideline']['score']}점):")
            for reason in r["guideline"]["reasons"]:
                print(f"       · {reason}")

        # 유전체 이유
        if r["genetics"]["reasons"]:
            print(f"     유전체 ({r['genetics']['score']}점):")
            for reason in r["genetics"]["reasons"]:
                print(f"       · {reason}")

        print()

    if unmatched:
        print("─" * 64)
        print("  인식하지 못한 항목:")
        for item in unmatched:
            print(f"    ? {item}")
        print()

    print("─" * 64)
    print("  등급 기준: [필수] 60+점 | [강력권고] 40-59 | [권고] 20-39 | [선택] <20")
    print()
    print("  근거: USPSTF A/B등급, 국가건강검진(2025), 한국학회 가이드라인, ACS")
    print("  ※ 본 추천은 의학적 진단을 대체하지 않습니다.")
    print("    최종 판단은 담당 의사와 상의하세요.")
    print("=" * 64)


# ─── CLI ────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("사용법: python recommend.py <items.txt> [--year 2026]")
        print()
        print("  items.txt: 병원 검진 항목 목록 (한 줄에 하나)")
        print("  --year:    추천 대상 연도 (기본: 올해)")
        return

    items_path = sys.argv[1]
    target_year = None

    if "--year" in sys.argv:
        idx = sys.argv.index("--year")
        if idx + 1 < len(sys.argv):
            target_year = int(sys.argv[idx + 1])

    if not os.path.exists(items_path):
        print(f"파일을 찾을 수 없습니다: {items_path}")
        return

    health_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(health_dir, "data")

    recommend(items_path, data_dir, target_year)


if __name__ == "__main__":
    main()
