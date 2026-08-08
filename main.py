import re
from pathlib import Path
import pandas as pd

MAJOR_WEIGHTS = {

    "공학": {
        "국어": 0.20,
        "수학": 0.35,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "자연과학": {
        "국어": 0.20,
        "수학": 0.35,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "의약보건": {
        "국어": 0.20,
        "수학": 0.35,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "인문사회": {
        "국어": 0.35,
        "수학": 0.20,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "상경": {
        "국어": 0.30,
        "수학": 0.25,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "교육": {
        "국어": 0.30,
        "수학": 0.25,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "기타": {
        "국어": 0.25,
        "수학": 0.25,
        "탐구1": 0.25,
        "탐구2": 0.25
    }
}

MAJOR_SEARCH_GROUPS = {

    "컴퓨터": [
        "컴퓨터",
        "소프트웨어",
        "인공지능",
        "ai",
        "데이터",
        "정보"
    ],

    "소프트웨어": [
        "컴퓨터",
        "소프트웨어",
        "인공지능",
        "ai",
        "데이터",
        "정보"
    ],

    "인공지능": [
        "컴퓨터",
        "소프트웨어",
        "인공지능",
        "ai",
        "데이터"
    ],

    "기계": [
        "기계",
        "자동차",
        "로봇",
        "메카트로닉스"
    ],

    "전자": [
        "전자",
        "전기",
        "반도체",
        "정보통신"
    ],

    "경영": [
        "경영",
        "회계",
        "금융",
        "무역"
    ],

    "경제": [
        "경제",
        "금융"
    ],

    "화학": [
        "화학",
        "화학공학",
        "신소재"
    ],

    "생명": [
        "생명",
        "생물",
        "바이오"
    ],

    "교육": [
        "교육"
    ]
}






EXCEL_FILE = Path(__file__).resolve().with_name(
   "전국대학_정시입결_공식통합_v1.xlsx"
)

SCIENCE_INQUIRY_KEYWORDS = [
    "물리",
    "화학",
    "생명",
    "지구"
]

SOCIAL_INQUIRY_KEYWORDS = [
    "생활과윤리",
    "윤리와사상",
    "한국지리",
    "세계지리",
    "동아시아사",
    "세계사",
    "경제",
    "정치와법",
    "사회문화"
]
def normalize_text(text):
    if text is None:
        return ""

    text = str(text).lower()

    text = re.sub(
        r"[가-힣a-z0-9}",
        "",
        text
    )

    return text

def classify_inquiry_subject(subject):
    subject = normalize_text(subject)

    for keyword in SCIENCE_INQUIRY_KEYWORDS:
        if normalize_text(keyword) in subject:
            return "과탐"

    for keyword in SOCIAL_INQUIRY_KEYWORDS:
        if normalize_text(keyword) in subject:
            return "사탐"

    return "기타"

def infer_major_group(major_name):
    name = normalize_text(major_name)

    medical_keywords = [
        "의예",
        "의학",
        "치의",
        "치과",
        "한의",
        "약학",
        "수의",
        "간호",
        "물리치료",
        "작업치료"
    ]

    for keyword in medical_keywords:
        if normalize_text(keyword) in name:
            return "의학보건"

    engineering_keywords = [
        "컴퓨터",
        "소프트웨어",
        "인공지능",
        "ai",
        "데이터",
        "전자",
        "전기",
        "기계",
        "로봇",
        "자동차",
        "메카트로닉스",
        "반도체",
        "정보통신",
        "통신",
        "토록",
        "건축공학",
        "산업공학",
        "신소재",
        "에너지",
        "환경공학",
        "화학공학",
        "조선",
        "항공"
    ]

    for keyword in engineering_keywords:
        if normalize_text(keyword) in name:
            return "공학"

    natural_keywords = [
        "수학",
        "통계",
        "물리",
        "화학",
        "생명",
        "생물",
        "지구",
        "천문",
        "환경과학",
        "식품",
        "농학"
    ]

    for keyword in natural_keywords:
        if normalize_text(keyword) in name:
            return "자연과학"

    business_keywords = [
        "경영",
        "경제",
        "회계",
        "세무",
        "금융",
        "무역",
        "국제통상"
    ]

    for keyword in business_keywords:
        if normalize_text(keyword) in name:
            return "상경"

    education_keywords = [
        "교육",
        "초등교육",
        "유아교육"
    ]

    for keyword in education_keywords:
        if normalize_text(keyword) in name:
            return "교육"

    humanities_keywords = [
        "국어",
        "영어",
        "문학",
        "사학",
        "역사",
        "철학",
        "사회",
        "심리",
        "행정",
        "정치",
        "외교",
        "법",
        "미디어",
        "언론",
        "광고",
        "문헌정보",
        "문학",
        "관광"
    ]

    for keyword in humanities_keywords:
        if normalize_text(keyword) in name:
            return "인문사회"

    return "기타"


def get_major_keywords(
        desired_major
):
    if not desired_major:
        return []

    desired = normalize_text(
        desired_major
    )

    for key, keywords in (
            MAJOR_SEARCH_GROUPS.items()
    ):

        normalized_key = (
            normalize_text(key)
        )

        if (
                normalized_key
                in desired
                or desired
                in normalized_key

        ):
            return [
                normalize_text(keyword)
                for keyword
                in keywords
            ]

    # 등록되지 않은 학과
    return [desired]


def input_percentile(subject):
   while True:
       try:
           score = float(
               input(f"{subject} 백분위 입력 (0 ~ 100): ")
           )

           if 0 <= score <= 100:
               return score

           print("0 ~ 100사이의 숫자를 입력하세요.")

       except ValueError:
           print("숫자를 입력하세요.")


def input_grade(subject):
   while True:
       try:
           grade = int(
               input(f"{subject} 등급 입력 (1~9): ")
           )

           if 1 <= grade <= 9:
               return grade

           print("1~9 사이의 등급을 입력하세요")
       except ValueError:
           print("정수를 입력하세요")


def input_student_scores():
    print()
    print("===== 모의고사 성적 입력 =====")

    inquiry1_name = input(
        "탐구1 과목명: "
    ).strip()

    inquiry2_name = input(
        "탐구2 과목명: "
    ).strip()

    Korean = input_percentile("국어")
    math = input_percentile("수학")

    inquiry1 = input_percentile(f"{inquiry1_name}")
    inquiry2 = input_percentile(f"{inquiry2_name}")

    english = input_grade("영어")
    history = input_grade("한국사")

    print()
    print("[희망 학과]")
    print("희망 학과과 없다면 Enter키를 누르세요")

    desired_major = input(
        "희망 학과: "
    ).strip()

    if desired_major == "":
        desired_major = None

    return {
        "국어": Korean,
        "수학": math,
        "탐구1": inquiry1,
        "탐구2": inquiry2,
        "영어": english,
        "한국사": history,
        "희망학과": desired_major
    }

def calculate_weighted_score(
        values,
        weight
):
    total_score = 0
    total_weight = 0

    for subject, weight in weight.items():
        value = values.get(subject)

        if value is None:
            continue

        if pd.isna(value):
            continue

        total_score += float(value) * weight

        if total_weight == 0:
            return None

        return total_score / total_weight

    def calculate_university_cutoff(
            row,
            weights
    ):
        university_score = {
            "국어": row.get(
                "국어70"
            ),
            "수학": row.get(
                "수학70"
            ),
            "탐구1": row.get(
                "탐구1_70"
            ),
            "탐구2": row.get(
                "탐구2_70"
            )
        }

        return calculate_weighted_score(university_score, weights)

def check_inquiry_fit(scores, major_group):

    inquiry1_type = (classify_inquiry_subject(scores["탐구1과목"]))

    inquiry2_type = (classify_inquiry_subject(scores["탐구1과목"]))

    types = [inquiry1_type, inquiry2_type]

    science_count = (
        types.count("과탐")
    )

    social_count = (
        types.count("사탐")
    )

    if major_group in ["공학", "자연과학", "의학보건"]:
        if science_count == 2:
            return "높음"
        elif science_count == 1:
            return "보통"
        else:
            return "대학교 혀용조건 확인"

    elif major_group in ["인문사회", "상경", "교육"]:
        if science_count == 2:
            return "높음"
        elif science_count == 1:
            return "보통"
        else:
            return "대학교 혀용조건 확인"

    return "별도 확인"

def filter_by__desired_major(
        df,
        desired_major
):

    if not desired_major:
        return df.copy()

    keywords = get_major_keywords(
        desired_major
    )

    def is_related(university_major):

        major=normalize_text(university_major)

        for key in keywords:
            if key in keywords:
                return True

        return False

    filtered = df[
        df["모집단위"].apply(is_related)].copy()

    return filtered

def percentile_to_grade(percentile):
    if percentile >= 95:
        return 1
    if percentile >= 86:
        return 2
    if percentile >= 77:
        return 3
    if percentile >= 62:
        return 4
    if percentile >= 45:
        return 5
    if percentile >= 30:
        return 6
    if percentile >= 19:
        return 7
    if percentile >= 10:
        return 8
    else:
        return 9

def evaluate_university_row(row, scores):
    major_group = infer_major_group(row["모집단위"])
    weights = MAJOR_WEIGHTS[major_group]

    student_values = {
        "국어": scores["국어"],
        "수학": scores["수학"],
        "탐구1": scores["탐구1"],
        "탐구2": scores["탐구2"]
    }

    student_score = calculat_weighted_score(student_values, weights)

    university_score = calculate_university_cutoff(row, weights)

    if university_score is None:
        return pd.Series({
            "학과게열": major_group,

            "학생가증점수": student_score,

            "대학가증입결": None,

            "학생환산등급": None,

            "대학환산등급": None,

            "추천유형": "자료부족",

            "가증점수치": None,

            "탐구적합도": check_inquiry_fit(student_score, major_group)
        })

    student_grade = (
        percentile_to_grade(
            student_score,
        )
    )

    university_grade = (
        percentile_to_grade(
            university_score
        )
    )


    grade_difference = student_grade - university_grade

    if 0.8 < grade_difference < 1.3:
        recommendation = "상향"

    elif 0 < grade_difference < 0.8:
        recommendation = "적정"

    elif grade_difference < 0:
        recommendation = "하향"

    else:
        recommendation = "밤위밖"

    return pd.Series({
        "학과 게열":
            major_group,

        "학생가증점수":
        round(student_score, 2),

        "대학가증입결":
        round(university_score, 2),

        "학생환산등급":
            student_grade,

        "추천유형":
            recommendation,

        "가증점수차":
        round(student_score - university_score, 2),

        "탐구적합도":
        check_inquiry_fit(score, major_group)
    })
