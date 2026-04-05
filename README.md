# 확률통계 채점 시스템

> 확률통계 TA 업무용 채점 관리 도구  
> Excel 가로 입력 → SQLite 세로 저장 → LMS/피드백 출력

## 프로젝트 구조

```
grading_system/
├── main.py                        # CLI 도구 (단일 파일)
├── README.md
├── input/
│   └── grades.xlsx                # 입력용 Excel (수강생/문제정보/제출현황)
├── db/
│   └── grades.db                  # SQLite DB (자동 생성)
└── rounds/
    ├── 과제_1/
    │   ├── template_과제_1.xlsx   # 채점 템플릿 (자동 생성)
    │   ├── lms_과제_1.xlsx        # LMS 점수 시트 (자동 생성)
    │   └── feedback_과제_1.xlsx   # 피드백 시트 (자동 생성)
    ├── 과제_2/
    ├── 퀴즈_1/
    ├── 중간_1/
    └── 기말_1/
```

## 의존성

```
python >= 3.10
pandas
openpyxl
```

## DB 스키마

4개 테이블, 모든 PK는 DDL로 명시.

### students
| 컬럼 | 타입 | 설명 |
|---|---|---|
| 학번 | INTEGER **PK** | 학생 고유 식별자 |
| 이름 | TEXT | 학생 이름 |

### problems
| 컬럼 | 타입 | 설명 |
|---|---|---|
| 평가유형 | TEXT **PK** | 과제 / 퀴즈 / 시험 |
| 회차 | INTEGER **PK** | 1, 2, 3… |
| 문제번호 | INTEGER **PK** | 1, 2, 3… |
| 만점 | INTEGER | 해당 문제 배점 |
| 정답_소수 | TEXT | 소수 형태 정답 (nullable) |
| 정답_분수 | TEXT | 분수 형태 정답 (nullable) |
| 채점기준 | TEXT | 부분점수 기준 서술 (nullable) |

### submissions
| 컬럼 | 타입 | 설명 |
|---|---|---|
| 학번 | INTEGER **PK, FK** | → students |
| 평가유형 | TEXT **PK** | 과제 / 퀴즈 / 시험 |
| 회차 | INTEGER **PK** | 1, 2, 3… |
| 제출여부 | TEXT | O / X |
| 제출시각 | TEXT | datetime 문자열 (과제만, 나머지 null) |
| 지각유형 | TEXT | 정상 / 1형 / 2형 |

### grades (정규화된 세로 구조)
| 컬럼 | 타입 | 설명 |
|---|---|---|
| 학번 | INTEGER **PK, FK** | → students |
| 평가유형 | TEXT **PK** | 과제 / 퀴즈 / 시험 |
| 회차 | INTEGER **PK** | 1, 2, 3… |
| 문제번호 | INTEGER **PK, FK** | → problems |
| 점수 | REAL | 해당 문제 득점 (nullable) |
| 감점이유 | TEXT | 감점 사유 (nullable) |

## 지각 제출 정책

| 유형 | 조건 | 처리 |
|---|---|---|
| 정상 | 마감 이전 제출 | 감점 없음 |
| 1형 | 마감 후 72시간 이내 | 총점 × 0.9 |
| 2형 | 마감 후 72시간 초과 | 0점 |
| 미제출 | 제출 없음 (제출여부=X) | 0점 |

## CLI 명령어

### 평가별 워크플로우

```bash
# 0. 학기 초 1회 — DB 초기화 및 수강생 등록
uv run main.py init
uv run main.py import-students

# 1. 문제정보 등록 (채점 전)
uv run main.py import-problems

# 2. 제출현황 등록 (마감일 직후)
uv run main.py import-submissions

# 3. 채점 템플릿 생성 → Excel에서 채점
uv run main.py gen-template --type 과제 --round 1
#   출력: rounds/과제_1/template_과제_1.xlsx
#   터미널에 정답/채점기준 표시
#   Excel 내 "채점기준(참고)" 시트에도 포함

# 4. 채점 결과 DB 등록
uv run main.py import-grades --file rounds/과제_1/template_과제_1.xlsx
#   파일명에서 평가유형/회차 자동 추출
#   수동 지정: --type 과제 --round 1

# 5. LMS 점수 시트 생성
uv run main.py export-lms --type 과제 --round 1
#   출력: rounds/과제_1/lms_과제_1.xlsx
#   지각 감점 자동 반영

# 6. 학생 피드백 시트 생성
uv run main.py export-feedback --type 과제 --round 1
#   출력: rounds/과제_1/feedback_과제_1.xlsx
#   감점이유 있는 셀 주황 하이라이트

# 7. 통계 확인 (터미널 출력)
uv run main.py stats --type 과제 --round 1
```

## Excel 입력 파일 구성 (grades.xlsx)

### Sheet "수강생"
| 학번 | 이름 |
|---|---|
| 202312345 | 이름A |

### Sheet "문제정보"
| 평가유형 | 회차 | 문제번호 | 만점 | 정답_소수 | 정답_분수 | 채점기준 |
|---|---|---|---|---|---|---|
| 과제 | 1 | 1 | 5 | 0.333 | 1/3 | 풀이 2점 + 답 3점 |

### Sheet "제출현황"
| 학번 | 평가유형 | 회차 | 제출여부 | 제출시각 | 지각유형 |
|---|---|---|---|---|---|
| 202312345 | 과제 | 1 | O | 2024-03-15 14:32 | 정상 |

## 채점 템플릿 형식 (자동 생성)

가로 구조: 학생당 1행, 문제별 `N_점수` + `N_감점` 교차 배치.

| 학번 | 이름 | 1_점수 | 1_감점 | 2_점수 | 2_감점 | … | 20_점수 | 20_감점 |
|---|---|---|---|---|---|---|---|---|
| 202312345 | 이름A | 5 | | 3 | 풀이 누락 | … | 5 | |

* Python이 읽어서 세로 구조(grades 테이블)로 unpivot 후 저장
* `INSERT OR REPLACE` 사용 — 동일 파일 재등록 시 덮어쓰기 (중복 안전)

## 출력 파일

### lms_*.xlsx
| 학번 | 이름 | 총점 | 지각유형 | 최종점수 |
|---|---|---|---|---|
| 202312345 | 이름A | 95 | 정상 | 95 |
| 202312346 | 이름B | 80 | 1형 | 72.0 |

* LMS에 수동 입력 시 참고용

### feedback_*.xlsx
| 학번 | 이름 | 문제번호 | 만점 | 점수 | 감점이유 | 제출여부 | 지각유형 |
|---|---|---|---|---|---|---|---|
| 202312345 | 이름A | 2 | 5 | 3 | 풀이 누락 | O | 정상 |

* 감점이유가 있는 셀은 주황 하이라이트
* 미제출 학생도 포함 (문제번호/점수 빈 칸, 제출여부=X)

## 주의사항

* `input/grades.xlsx`의 Sheet 이름은 정확히 `수강생`, `문제정보`, `제출현황`이어야 함
* 학번은 정수 타입으로 입력 (문자열 불가)
* `import-*` 명령은 모두 `INSERT OR REPLACE` → 같은 데이터 재실행해도 안전
* 모든 `export-*`와 `gen-template`은 students 테이블 기준 LEFT JOIN → 미등록 학생 누락 방지
* `gen-template` 실행 전 반드시 `import-students` + `import-problems` 완료 필요
* `export-lms` 실행 전 반드시 `import-submissions` + `import-grades` 완료 필요

## 향후 확장 가능 항목

* [ ] grades.xlsx 내 Sheet 이름 유효성 검증
* [ ] 학번 FK 위반 시 친화적 에러 메시지
* [ ] 학기 전체 성적 집계 (과제 4회 + 퀴즈 + 시험 합산)
* [ ] 피드백 시트를 학생별 개별 PDF로 변환
* [ ] GitHub CI/CD 연동 (테스트 자동화)
