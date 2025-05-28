# PPT_Auto_Saver (PowerPoint Auto Saver / 파워포인트 자동 저장 프로그램)

An application that automatically saves your open PowerPoint presentations at regular intervals.

This is a GUI-based application developed in Python using Tkinter and `pywin32`, designed to automatically save your active Microsoft PowerPoint presentations at user-defined intervals.

## Functionality and Purpose

This program provides a user-friendly interface to manage and automate the saving process of PowerPoint files:

* **List Active Presentations:** Displays a list of all currently open PowerPoint presentations, showing their titles and file paths.
* **Manage Auto-Save List:**
    * Users can add specific presentations from the active list to a dedicated "Auto-Save List."
    * Presentations can be removed from this auto-save list.
* **Timed Auto-Save:**
    * Presentations in the "Auto-Save List" are automatically saved at a time interval (in seconds) set by the user.
    * The default interval is 60 seconds but can be customized.
* **Configuration Persistence:** The auto-save list and the chosen save interval are saved locally in a `autosave_ppt_config.json` file, so your settings are remembered for future sessions.
* **Real-time Updates:** The list of active presentations refreshes periodically to reflect any newly opened or closed PowerPoint files. Presentations already in the auto-save list are visually indicated.
* **User Interface:** Built with Tkinter for a simple and intuitive graphical user interface.
* **PowerPoint Interaction:** Utilizes the `pywin32` library to interact with the Microsoft PowerPoint application.

This app is useful for preventing data loss due to unexpected shutdowns or if you frequently forget to save your work in PowerPoint.

## Prerequisites

Before using this application, ensure you have the following:

1.  **Microsoft PowerPoint:** This program **requires** Microsoft PowerPoint to be installed on your Windows system.
2.  **Python (for running the script directly):** If you intend to run the Python script (`.py`), you'll need Python 3.x installed.
    * **Required Python Libraries:** `pywin32` and `tkinter` (Tkinter is usually included with standard Python installations). You can install `pywin32` using pip:
        ```bash
        pip install pywin32
        ```

## How to Use

There are two ways to use this application:

**A. Using the Executable (`.exe` file - Recommended for most users):**

1.  Download the latest `PPT_Auto_Saver.exe` file from the **[Releases](https://github.com/nopigom119/PPT_Auto_Saver/releases)** page of this repository.
2.  Ensure Microsoft PowerPoint is installed on your system.
3.  Run the `PPT_Auto_Saver.exe` file.
4.  **Interface Overview:**
    * **Currently Open PPTs List (Left Panel):** Shows all PowerPoint files currently open.
        * Select a presentation and click "Add to Auto-Save List >>" to include it for automatic saving.
        * Presentations already in the auto-save list or not yet saved (e.g., "Presentation1") might have different styling or disabled "Add" button.
        * Click "Refresh List" to manually update this list.
    * **Auto-Save Target PPTs (Right Panel):** Shows presentations scheduled for automatic saving.
        * Select a presentation and click "<< Remove from List" to stop auto-saving it.
    * **Auto-Save Interval:** Enter the desired save interval in seconds (e.g., 30, 60, 120).
    * **Set Interval & Start/Restart:** Click this button to apply the new interval and (re)start the auto-save process.
5.  The application will automatically save the presentations in the "Auto-Save Target PPTs" list at the specified interval.
6.  Your auto-save list and interval settings are saved when you close the application or change the interval.

**B. Running the Python Script (`.py` file):**

1.  Clone this repository or download the `ppt_auto_saver.py` script.
2.  Ensure Python 3.x and Microsoft PowerPoint are installed (see Prerequisites).
3.  Install the required Python library:
    ```bash
    pip install pywin32
    ```
4.  Open a terminal or command prompt.
5.  Navigate to the directory where you saved the script.
6.  Run the script using the command: `python ppt_auto_saver.py`
7.  Follow steps 4-6 from the "Using the Executable" section above.

## License

This program is licensed under the **Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License (CC BY-NC-SA 4.0)**.

* **Attribution:** You must give appropriate credit, provide a link to the license, and indicate if changes were made.
* **Non-Commercial Use:** You may not use this program for commercial purposes.
* **Modification Allowed:** You can modify this program or create derivative works.
* **Same Conditions for Change Permission:** If you modify or create derivative works of this program, you must distribute your contributions under the same license as the original.

You can check the license details on the Creative Commons website: [https://creativecommons.org/licenses/by-nc-sa/4.0/deed.en](https://creativecommons.org/licenses/by-nc-sa/4.0/deed.en)

## Contact

For inquiries about this program, please contact [rycbabd@gmail.com].

---

# PPT_Auto_Saver (파워포인트 자동 저장 프로그램)

사용자가 지정한 간격으로 열려 있는 파워포인트 프레젠테이션을 자동으로 저장하는 애플리케이션입니다.

이 프로그램은 Tkinter와 `pywin32`를 사용하여 Python으로 개발된 GUI 기반 애플리케이션으로, 사용자가 정의한 간격으로 현재 활성화된 Microsoft PowerPoint 프레젠테이션을 자동으로 저장하도록 설계되었습니다.

## 기능 및 목적

본 프로그램은 파워포인트 파일의 저장 프로세스를 관리하고 자동화하는 사용자 친화적인 인터페이스를 제공합니다:

* **활성 프레젠테이션 목록 표시:** 현재 열려 있는 모든 파워포인트 프레젠테이션의 제목과 파일 경로를 목록으로 보여줍니다.
* **자동 저장 목록 관리:**
    * 사용자는 활성 목록에서 특정 프레젠테이션을 전용 "자동 저장 목록"에 추가할 수 있습니다.
    * 이 자동 저장 목록에서 프레젠테이션을 제거할 수 있습니다.
* **시간 지정 자동 저장:**
    * "자동 저장 목록"에 있는 프레젠테이션은 사용자가 설정한 시간 간격(초 단위)마다 자동으로 저장됩니다.
    * 기본 간격은 60초이며 사용자 정의할 수 있습니다.
* **설정 유지:** 자동 저장 목록과 선택한 저장 간격은 로컬 `autosave_ppt_config.json` 파일에 저장되어 다음 세션에서도 설정을 기억합니다.
* **실시간 업데이트:** 활성 프레젠테이션 목록은 새로 열리거나 닫힌 파워포인트 파일을 반영하기 위해 주기적으로 새로고침됩니다. 이미 자동 저장 목록에 있는 프레젠테이션은 시각적으로 구분됩니다.
* **사용자 인터페이스:** 간단하고 직관적인 그래픽 사용자 인터페이스를 위해 Tkinter로 구축되었습니다.
* **파워포인트 연동:** `pywin32` 라이브러리를 활용하여 Microsoft PowerPoint 애플리케이션과 상호 작용합니다.

이 앱은 예기치 않은 시스템 종료로 인한 데이터 손실을 방지하거나 파워포인트 작업 내용을 자주 저장하는 것을 잊는 경우에 유용합니다.

## 사전 준비 사항

이 애플리케이션을 사용하기 전에 다음 사항을 확인하십시오.

1.  **Microsoft PowerPoint:** 이 프로그램은 Windows 시스템에 Microsoft PowerPoint가 설치되어 있어야 **필수적으로** 작동합니다.
2.  **Python (스크립트 직접 실행 시):** Python 스크립트(`.py`)를 직접 실행하려면 Python 3.x 버전이 설치되어 있어야 합니다.
    * **필수 Python 라이브러리:** `pywin32` 및 `tkinter` (Tkinter는 일반적으로 표준 Python 설치에 포함되어 있습니다). `pywin32`는 pip를 사용하여 설치할 수 있습니다:
        ```bash
        pip install pywin32
        ```

## 사용 방법

이 애플리케이션을 사용하는 방법에는 두 가지가 있습니다.

**A. 실행 파일 (`.exe` 파일 사용 - 대부분 사용자에게 권장):**

1.  이 저장소의 **[Releases](https://github.com/nopigom119/PPT_Auto_Saver/releases)** 페이지에서 최신 `PPT_Auto_Saver.exe` 파일을 다운로드합니다.
2.  시스템에 Microsoft PowerPoint가 설치되어 있는지 확인합니다.
3.  `PPT_Auto_Saver.exe` 파일을 실행합니다.
4.  **인터페이스 개요:**
    * **현재 열려있는 PPT 목록 (왼쪽 패널):** 현재 열려 있는 모든 파워포인트 파일을 보여줍니다.
        * 프레젠테이션을 선택하고 "자동 저장 목록에 추가 >>" 버튼을 클릭하여 자동 저장 대상에 포함시킵니다.
        * 이미 자동 저장 목록에 있거나 아직 저장되지 않은 파일(예: "Presentation1")은 다르게 표시되거나 "추가" 버튼이 비활성화될 수 있습니다.
        * "목록 새로고침" 버튼을 클릭하여 이 목록을 수동으로 업데이트할 수 있습니다.
    * **자동 저장 대상 PPT (오른쪽 패널):** 자동 저장이 예약된 프레젠테이션을 보여줍니다.
        * 프레젠테이션을 선택하고 "<< 목록에서 제거" 버튼을 클릭하여 자동 저장을 중지합니다.
    * **자동 저장 간격:** 원하는 저장 간격을 초 단위로 입력합니다 (예: 30, 60, 120).
    * **간격 설정 및 저장 시작/재시작:** 이 버튼을 클릭하여 새 간격을 적용하고 자동 저장 프로세스를 (재)시작합니다.
5.  애플리케이션은 "자동 저장 대상 PPT" 목록에 있는 프레젠테이션을 지정된 간격으로 자동 저장합니다.
6.  자동 저장 목록 및 간격 설정은 애플리케이션을 닫거나 간격을 변경할 때 저장됩니다.

**B. Python 스크립트 (`.py` 파일 실행):**

1.  이 저장소를 복제하거나 `ppt_auto_saver.py` 스크립트를 다운로드합니다.
2.  Python 3.x 및 Microsoft PowerPoint가 설치되어 있는지 확인합니다 (사전 준비 사항 참조).
3.  필수 Python 라이브러리를 설치합니다:
    ```bash
    pip install pywin32
    ```
4.  터미널 또는 명령 프롬프트를 엽니다.
5.  스크립트를 저장한 디렉토리로 이동합니다.
6.  명령을 사용하여 스크립트를 실행합니다: `python ppt_auto_saver.py`
7.  위의 "실행 파일 사용" 섹션의 4-6단계를 따릅니다.

## 라이선스

본 프로그램은 **크리에이티브 커먼즈 저작자표시-비영리-동일조건변경허락 4.0 국제 라이선스 (CC BY-NC-SA 4.0)** 에 따라 이용할 수 있습니다.

* **출처 표시:** 본 프로그램의 출처 (작성자 또는 개발팀)를 명시해야 합니다.
* **비상업적 이용:** 본 프로그램을 상업적인 목적으로 이용할 수 없습니다.
* **변경 가능:** 본 프로그램을 수정하거나 2차 저작물을 만들 수 있습니다.
* **동일 조건 변경 허락:** 2차 저작물에 대해서도 동일한 조건으로 이용 허락해야 합니다.

자세한 내용은 크리에이티브 커먼즈 홈페이지에서 확인하실 수 있습니다: [https://creativecommons.org/licenses/by-nc-sa/4.0/deed.ko](https://creativecommons.org/licenses/by-nc-sa/4.0/deed.ko)

## 문의

본 프로그램에 대한 문의사항은 [rycbabd@gmail.com] 로 연락주시기 바랍니다.
