# Contributing to PPT_Auto_Saver

Thank you for your interest in contributing to PPT_Auto_Saver (PowerPoint Auto Saver / 파워포인트 자동 저장 프로그램)! All contributions are welcome.
We appreciate any form of contribution, be it bug fixes, feature additions, documentation improvements, or helping with testing.

## How to Contribute

### 1. Reporting Issues

* Please register new feature suggestions or bug reports on the [PPT_Auto_Saver Issues page](https://github.com/nopigom119/PPT_Auto_Saver/issues).
* Make sure the issue title is concise and clear.
* Provide detailed descriptions of the problem or suggestion in the issue content. Please include:
    * Steps to reproduce the bug.
    * Expected behavior and actual behavior.
    * Your operating system and version.
    * The version of Microsoft PowerPoint you are using.
    * The version of the PPT_Auto_Saver application you are using.
    * Screenshots or error messages if applicable.

### 2. Contributing Code

1.  Fork the [PPT_Auto_Saver repository](https://github.com/nopigom119/PPT_Auto_Saver).
2.  Create a new branch in your forked repository for your contribution.
    * Branch names should be descriptive. (e.g., `feature/add-multi-select-save`, `bugfix/fix-config-load-error-45`)
3.  Write your code. Please adhere to the [PEP 8](https://www.python.org/dev/peps/pep-0008/) coding style guidelines.
4.  Test your changes thoroughly, especially with various PowerPoint versions, file states (saved, unsaved, new), and configurations. If adding new functionality that interacts with PowerPoint COM objects, ensure proper error handling and resource management.
5.  Write concise and clear commit messages. (e.g., `feat: Allow selecting multiple PPTs for auto-save`, `fix: Handle COMException when PowerPoint is not running (issue #45)`)
6.  Submit a [pull request](https://github.com/nopigom119/PPT_Auto_Saver/pulls) to the `main` branch of the original PPT_Auto_Saver repository.
    * Clearly describe the changes you have made in your pull request.
    * Link to any relevant issues.

### 3. Contributing to Documentation

* You can contribute to improving documentation, such as the README file, comments within the code, code explanations, usage examples, or tips for troubleshooting.
* Make sure the documentation is clear, concise, and easy to understand.

## Code Writing Rules

* Follow the [PEP 8](https://www.python.org/dev/peps/pep-0008/) coding style guidelines.
* Write clear and descriptive docstrings for all public modules, classes, functions, and methods.
* Add comments to your code where necessary to explain complex logic.
* Pay special attention to how `win32com` objects are handled, ensuring COM objects are properly initialized and released to prevent memory leaks or application instability. User inputs or data retrieved from PowerPoint should be handled carefully.

## Commit Message Rules

* Write concise and clear commit messages.
* A good commit message should briefly describe the change.
* Consider using [Conventional Commits](https://www.conventionalcommits.org/) for a structured approach (e.g., `feat:`, `fix:`, `docs:`, `style:`, `refactor:`, `test:`, `chore:`).

## Pull Request Procedure

1.  Ensure your code adheres to the project's coding standards and all tests pass (if applicable).
2.  Create a pull request from your feature branch to the `main` branch of the `nopigom119/PPT_Auto_Saver` repository.
3.  Provide a clear description of the changes in the pull request.
4.  Be prepared to discuss your changes and make further modifications if requested by the maintainers.

## Inquiries

* For inquiries or discussions about contributing to PPT_Auto_Saver, please use the [PPT_Auto_Saver Issues page](https://github.com/nopigom119/PPT_Auto_Saver/issues).

## License

By contributing to PPT_Auto_Saver, you agree that your contributions will be licensed under its CC BY-NC-SA 4.0 License.

---

# PPT_Auto_Saver에 기여하기

PPT_Auto_Saver (PowerPoint Auto Saver / 파워포인트 자동 저장 프로그램)에 관심을 가져주셔서 감사합니다! 모든 기여를 환영합니다.
버그 수정, 기능 추가, 문서 개선, 테스트 지원 등 어떤 형태의 기여든 감사하게 생각합니다.

## 기여 방법

### 1. 이슈 등록

* 새로운 기능 제안이나 버그 보고는 [PPT_Auto_Saver 이슈 페이지](https://github.com/nopigom119/PPT_Auto_Saver/issues)에 등록해주세요.
* 이슈 제목은 간결하고 명확하게 작성해주세요.
* 이슈 내용에는 문제 상황이나 제안 내용을 상세하게 설명해주세요. 다음 정보를 포함해주세요:
    * 버그 재현 단계.
    * 예상되는 동작과 실제 동작.
    * 사용 중인 운영체제 및 버전.
    * 사용 중인 Microsoft PowerPoint 버전.
    * 사용 중인 PPT_Auto_Saver 애플리케이션 버전.
    * 해당되는 경우 스크린샷 또는 오류 메시지.

### 2. 코드 기여

1.  [PPT_Auto_Saver 저장소](https://github.com/nopigom119/PPT_Auto_Saver)를 포크해주세요.
2.  포크한 저장소에서 기여를 위한 새로운 브랜치를 만들어주세요.
    * 브랜치 이름은 설명적으로 작성해주세요. (예: `feature/다중선택-저장-기능추가`, `bugfix/설정파일-로드오류-45-수정`)
3.  코드를 작성해주세요. [PEP 8](https://www.python.org/dev/peps/pep-0008/) 코딩 스타일 가이드라인을 준수해주세요.
4.  특히 다양한 파워포인트 버전, 파일 상태(저장됨, 저장안됨, 새 파일), 설정으로 변경 사항을 철저히 테스트해주세요. 파워포인트 COM 객체와 상호작용하는 새로운 기능을 추가하는 경우, 적절한 오류 처리 및 리소스 관리를 확인해주세요.
5.  간결하고 명확한 커밋 메시지를 작성해주세요. (예: `feat: 자동 저장을 위한 다중 PPT 선택 기능 추가`, `fix: 파워포인트 미실행 시 COMException 처리 (이슈 #45)`)
6.  원본 PPT_Auto_Saver 저장소의 `main` 브랜치로 [풀 리퀘스트](https://github.com/nopigom119/PPT_Auto_Saver/pulls)를 보내주세요.
    * 풀 리퀘스트에 변경한 내용을 명확하게 설명해주세요.
    * 관련된 이슈가 있다면 링크해주세요.

### 3. 문서 기여

* README 파일, 코드 내 주석, 코드 설명, 사용 예시, 문제 해결 팁 등 문서 개선에 기여해주실 수 있습니다.
* 문서 내용은 명확하고 간결하며 이해하기 쉽게 작성해주세요.

## 코드 작성 규칙

* [PEP 8](https://www.python.org/dev/peps/pep-0008/) 코딩 스타일 가이드라인을 준수해주세요.
* 모든 공개 모듈, 클래스, 함수, 메소드에 대해 명확하고 설명적인 독스트링(docstring)을 작성해주세요.
* 복잡한 로직을 설명하기 위해 필요한 경우 코드에 주석을 추가해주세요.
* `win32com` 객체를 처리하는 방식에 특히 주의하여 메모리 누수나 애플리케이션 불안정을 방지하기 위해 COM 객체가 올바르게 초기화되고 해제되는지 확인해주세요. 파워포인트에서 가져온 사용자 입력이나 데이터는 신중하게 처리해야 합니다.

## 커밋 메시지 규칙

* 간결하고 명확한 커밋 메시지를 작성해주세요.
* 좋은 커밋 메시지는 변경 사항을 간략하게 설명해야 합니다.
* 구조화된 접근 방식을 위해 [Conventional Commits](https://www.conventionalcommits.org/ko/v1.0.0/) 사용을 고려해보세요 (예: `feat:`, `fix:`, `docs:`, `style:`, `refactor:`, `test:`, `chore:`).

## 풀 리퀘스트 절차

1.  코드가 프로젝트의 코딩 표준을 준수하고 모든 테스트를 통과하는지 확인해주세요 (해당되는 경우).
2.  기능 브랜치에서 `nopigom119/PPT_Auto_Saver` 저장소의 `main` 브랜치로 풀 리퀘스트를 생성해주세요.
3.  풀 리퀘스트에 변경 사항에 대한 명확한 설명을 제공해주세요.
4.  관리자의 요청이 있을 경우 변경 사항에 대해 논의하고 추가 수정을 할 준비가 되어 있어야 합니다.

## 문의

* PPT_Auto_Saver 기여에 대한 문의나 논의는 [PPT_Auto_Saver 이슈 페이지](https://github.com/nopigom119/PPT_Auto_Saver/issues)를 이용해주세요.

## 라이선스

PPT_Auto_Saver에 기여함으로써 귀하의 기여물은 해당 프로젝트의 CC BY-NC-SA 4.0 라이선스에 따라 사용이 허가됨에 동의하는 것으로 간주됩니다.
