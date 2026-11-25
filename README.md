# Google Photos 업로드 전 옛날 영상 회전 & 병합 & MP4 정리 도구 (Windows)

2013년 이전 HDD에 보관된 구형 영상(`.avi`, `.3gp`, `.wmv`, 구형 `.mov` 등)을 Google Photos에 올리기 전에 브라우저 화면에서 순서대로 편하게 정리할 수 있는 로컬 웹 도구입니다.

1. **사전 준비 & 폴더 선택:** Git / Python / FFmpeg 설치 여부 자동 확인 및 영상 폴더 스캔
2. **영상 편집 (회전/병합):** 썸네일을 보며 돌아간 영상 회전 + 쪼개진 클립 병합 그룹 지정
3. **변환 전 작업 내용 확인:** 변환/병합 대상 리뷰 + **세이프 모드** vs **원본 백업 후 교체 모드** 선택
4. **실시간 변환 진행:** 진행률 바 & 실시간 로그를 보면서 변환 실행 + 촬영 날짜 자동 보존

---

## 1. Windows 필수 프로그램 설치 (Python + FFmpeg, Git은 선택)

> **필수:** Python 3, FFmpeg / **선택:** Git (`git clone` / `git pull` 방식으로 받고 싶을 때만 필요하며, ZIP 파일로 내려받으면 Git 없이도 바로 사용할 수 있습니다.)

### 방법 1: 가장 쉬운 설치 (PowerShell `winget` 사용, Windows 10/11 권장)

1. 시작 메뉴에서 **PowerShell** (또는 터미널)을 엽니다.
2. 아래 명령어를 순서대로 복사해서 실행합니다. (`Git.Git`은 ZIP으로 내려받을 경우 생략 가능합니다.)
   ```powershell
   winget install -e --id Python.Python.3.12
   winget install -e --id Gyan.FFmpeg
   winget install -e --id Git.Git
   ```
3. 설치가 끝나면 **반드시 PowerShell 창을 닫았다가 새로 연 뒤** 정상 설치를 확인합니다.
   ```powershell
   python --version
   ffmpeg -version
   git --version
   ```
   *(만약 `python`이 안 잡히고 `py`만 된다면 `python` 대신 `py`로 실행하시면 됩니다.)*

### 방법 2: 직접 다운로드해서 설치하기 (`winget`이 안 될 때)

- **Python (필수):** [python.org/downloads](https://www.python.org/downloads/)에서 설치 파일을 받아 실행하되, 첫 화면 하단의 **"Add python.exe to PATH"** 체크박스를 **반드시 체크**하고 설치하세요.
- **FFmpeg (필수):** [gyan.dev/ffmpeg/builds](https://www.gyan.dev/ffmpeg/builds/)에서 `ffmpeg-release-essentials.zip`을 받아 `C:\ffmpeg` 등에 압축을 푼 뒤, Windows 검색 → **"시스템 환경 변수 편집"** → `Path`에 `C:\ffmpeg\bin`을 추가하세요.
- **Git (선택):** [git-scm.com/download/win](https://git-scm.com/download/win)에서 설치 파일을 받아 **Next**만 눌러 기본 설정으로 설치하세요.

---

## 2. GitHub에서 코드 내려받기

### 방법 A: ZIP 파일로 바로 받기 (Git 없이 가장 쉬움)

1. 저장소 페이지([https://github.com/chaeci/gphotos-video-prep](https://github.com/chaeci/gphotos-video-prep))에 접속해 **`<> Code`** → **`Download ZIP`**을 클릭합니다.
   *(또는 바로 다운로드: [https://github.com/chaeci/gphotos-video-prep/archive/refs/heads/main.zip](https://github.com/chaeci/gphotos-video-prep/archive/refs/heads/main.zip))*
2. 내려받은 압축 파일을 원하는 폴더(예: `C:\Tools\gphotos-video-prep`)에 압축 해제합니다.

### 방법 B: Git으로 내려받기 (`git clone`, 추후 업데이트 편함)

PowerShell을 열고 원하는 폴더(예: `C:\Tools`)로 이동해서 저장소를 내려받습니다. 공개(Public) 저장소이므로 별도의 GitHub 로그인 없이 바로 받아집니다.

```powershell
mkdir C:\Tools
cd C:\Tools
git clone https://github.com/chaeci/gphotos-video-prep.git
```

- 나중에 최신 버전으로 업데이트할 때는 아래 명령어를 실행하면 됩니다.
  ```powershell
  cd C:\Tools\gphotos-video-prep
  git pull
  ```

---

## 3. 웹 마법사로 실행하기 (가장 쉬운 방법 ⭐)

`C:\Tools\gphotos-video-prep` 폴더에 있는 **`start.bat`** 파일을 더블클릭하거나, 터미널에서 아래처럼 실행하세요.

```powershell
cd C:\Tools\gphotos-video-prep
python video_prep.py
```

잠시 후 브라우저(`http://127.0.0.1:8765`)가 자동으로 열리며 4단계 화면이 순서대로 진행됩니다.

1. **1단계 — 사전 준비 & 폴더 선택**
   - Python / FFmpeg / Git 설치 상태를 자동으로 확인해 줍니다. (미설치 항목이 있으면 설치 명령어를 바로 복사할 수 있습니다.)
   - 영상이 들어 있는 폴더(예: `D:\OldPhotos`)를 입력하거나 `[📂 폴더 찾기]`로 선택한 뒤 **`🔍 영상 스캔 시작`**을 누릅니다.
   - *(이전에 스캔해 둔 기록이 있으면 `[📂 기존 검수 이어하기]` 버튼으로 바로 2단계로 넘어갈 수 있습니다.)*
2. **2단계 — 영상 편집 (회전 / 병합)**
   - **회전 맞추기:** 누워 있거나 거꾸로 된 영상만 `↺ 90°` / `↻ 90°` / `↕ 180°` 버튼을 눌러 똑바로 맞춥니다.
   - **영상 병합(분할 녹화 이어 붙이기):**
     - **수동 묶기:** 각 카드 왼쪽 위의 `[병합]` 체크박스로 선택(첫 클릭 후 마지막 카드 `Shift + 클릭`으로 범위 선택 가능) → **`🔗 선택한 영상 병합 그룹 만들기`**
     - **자동 묶기:** 상단 간격(초, 기본 `15`초) 설정 후 **`⚡ 연속 영상 자동 묶기`** 클릭 시 같은 폴더 내 연속 촬영 클립을 자동 그룹화
     - 상단 **병합 그룹 목록**에서 클립 순서(`◀` / `▶`), 출력 파일명, 그룹 일괄 회전(`전체 ↻ 90°` 등), `그룹 해제`를 자유롭게 편집할 수 있습니다.
   - 편집이 끝나면 오른쪽 위의 **`다음: 변환 전 작업 내용 확인 →`**을 누릅니다.
3. **3단계 — 변환 전 작업 내용 확인 & 모드 선택**
   - 단독 변환 / 회전 / 병합 그룹 / 날짜 보완 건수와 전체 작업 목록을 마지막으로 리뷰합니다.
   - 저장 모드를 선택합니다:
     - **🛡️ 세이프 모드 (`output`, 기본값):** 원본 폴더는 전혀 건드리지 않고 `_gphotos_work\converted\`에 결과물을 저장합니다.
     - **📦 원본 백업 후 교체 모드 (`replace-backup`):** 원본(분할 클립 포함)을 `_gphotos_work\original_backup\`으로 안전하게 옮긴 뒤, 변환·병합된 `.mp4`를 원래 폴더 위치에 배치해 중복 없이 Google Photos에 폴더째 올릴 수 있게 합니다.
   - **`🚀 변환 시작하기`**를 누르면 4단계로 이동합니다.
4. **4단계 — 실시간 변환 진행**
   - 진행률 바, 성공/건너뜀/실패 카운트, 실시간 변환 로그를 보며 기다립니다.
   - 필요하면 `[⏹ 현재 작업 후 안전 중단]`으로 안전하게 멈출 수 있고, 완료 후 `[📂 결과 폴더 열기]` / `[📂 백업 폴더 열기]` 버튼으로 바로 확인할 수 있습니다.

---

## 4. Google Photos 업로드 전 확인 팁

1. 변환된 `.mp4` 몇 개를 Windows 기본 플레이어로 재생해 **화면 방향과 소리**가 정상인지 확인하세요.
2. 테스트로 **10~20개만 먼저 Google Photos에 업로드**해 타임라인 날짜가 원래 연도로 정확히 들어가는지 확인한 뒤 전체 업로드를 진행하세요.
3. `--mode replace-backup`을 사용하셨다면, Google Photos에 폴더째 올리기 전에 `_gphotos_work` 폴더는 업로드 대상에서 제외하거나 다른 드라이브로 옮겨두세요. 업로드가 완전히 끝난 뒤 `_gphotos_work\original_backup`을 정리하시면 됩니다.

---

## 5. (참고) 터미널 CLI 명령어로 직접 실행하기

웹 마법사 대신 터미널 명령어만으로도 각 단계를 개별 실행할 수 있습니다.

```powershell
# 1단계: 스캔
python video_prep.py scan "D:\OldPhotos"

# 2단계: 브라우저 검수 UI만 열기
python video_prep.py review --work-dir "D:\OldPhotos\_gphotos_work"

# 3단계: 미리보기 및 변환
python video_prep.py convert --work-dir "D:\OldPhotos\_gphotos_work" --dry-run
python video_prep.py convert --work-dir "D:\OldPhotos\_gphotos_work" --mode output
python video_prep.py convert --work-dir "D:\OldPhotos\_gphotos_work" --mode replace-backup --fix-missing-dates
```
