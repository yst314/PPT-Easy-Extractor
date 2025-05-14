# PowerPoint自動出力システム

このツールは、PowerPointファイル(.pptx)からテキストと画像を抽出するPython製のデスクトップアプリケーションです。

## 必要なもの

- Python 3.8以上

## インストール方法

Pythonをまだインストールしていない場合は、以下の手順でインストールしてください。

### Windows

1.  **Pythonのダウンロード**:
    *   [Python公式サイトのダウンロードページ](https://www.python.org/downloads/windows/)にアクセスします。
    *   「Latest Python 3 Release - Python x.x.x」のような最新版のリンクをクリックします。
    *   ページ下部にある「Files」セクションから、お使いのWindowsが64ビット版であれば「Windows installer (64-bit)」、32ビット版であれば「Windows installer (32-bit)」をダウンロードします。(最近のPCのほとんどは64ビットです)

2.  **Pythonのインストール**:
    *   ダウンロードしたインストーラー（`.exe`ファイル）を実行します。
    *   **重要**: インストーラーの最初の画面で、必ず「Add Python x.x to PATH」または「Add python.exe to Path」のチェックボックスにチェックを入れてください。これにチェックを入れないと、コマンドプロンプトでPythonを簡単に実行できません。
    *   「Install Now」をクリックしてインストールを開始します。
    *   インストールが完了するまで待ちます。

3.  **インストールの確認**:
    *   コマンドプロンプト（`Win`キー + `R`を押し、「cmd」と入力してEnter）を開きます。
    *   以下のコマンドを入力してEnterキーを押します。
        ```bash
        python --version
        ```
    *   「Python 3.x.x」のようにバージョンが表示されれば、正しくインストールされています。

### macOS

1.  **Homebrewのインストール (推奨)**:
    *   Homebrewがインストールされていない場合は、[Homebrew公式サイト](https://brew.sh/index_ja)の手順に従ってインストールします。ターミナルで1行のコマンドを実行するだけです。

2.  **Pythonのインストール (Homebrew経由)**:
    *   ターミナルを開きます。
    *   以下のコマンドを入力してEnterキーを押します。
        ```bash
        brew install python
        ```
    *   インストールが完了するまで待ちます。

3.  **(代替)Pythonのインストール (公式サイトから)**:
    *   [Python公式サイトのダウンロードページ](https://www.python.org/downloads/macos/) にアクセスし、macOS用のインストーラーをダウンロードして実行します。

4.  **インストールの確認**:
    *   ターミナルで以下のコマンドを入力してEnterキーを押します。
        ```bash
        python3 --version
        ```
    *   「Python 3.x.x」のようにバージョンが表示されれば、正しくインストールされています。(macOSでは `python` の代わりに `python3` を使うことが一般的です)

### Linux (Ubuntu/Debian系の場合)

多くのLinuxディストリビューションにはPythonがプリインストールされていることが多いですが、最新版にしたい場合やインストールされていない場合は以下の手順です。

1.  **Pythonのインストール**:
    *   ターミナルを開きます。
    *   以下のコマンドを入力してEnterキーを押します。
        ```bash
        sudo apt update
        sudo apt install python3 python3-pip python3-venv
        ```

2.  **インストールの確認**:
    *   ターミナルで以下のコマンドを入力してEnterキーを押します。
        ```bash
        python3 --version
        ```
    *   「Python 3.x.x」のようにバージョンが表示されれば、正しくインストールされています。

## プロジェクトのセットアップと実行

ここでは、2通りのセットアップ方法を説明します。初心者の方は「簡単な方法」をおすすめします。

### 簡単な方法 (Gitやuvを使わない)

1.  **ソースコードのダウンロード**:
    *   このプロジェクトのページ (GitHubなど) からソースコードをZIPファイルとしてダウンロードします。
    *   ダウンロードしたZIPファイルを任意のフォルダに展開（解凍）します。

2.  **プロジェクトフォルダへの移動**:
    *   コマンドプロンプト (Windows) またはターミナル (macOS/Linux) を開きます。
    *   `cd` コマンドを使って、展開したプロジェクトのフォルダに移動します。
        ```bash
        cd path/to/your/PPT-Easy-Extractor
        ```
        (`path/to/your/` の部分は実際のフォルダの場所に合わせてください)

3.  **仮想環境の作成と有効化**:
    プロジェクトごとにPythonの環境を分離することで、他のプロジェクトとの影響を防ぎます。
    *   **仮想環境の作成** (初回のみ):
        ```bash
        python -m venv .venv
        ```
        または (Pythonのバージョンによっては `python3` を使います):
        ```bash
        python3 -m venv .venv
        ```
        これにより、プロジェクトフォルダ内に `.venv` という名前の仮想環境フォルダが作成されます。
    *   **仮想環境の有効化** (ターミナルを起動するたびに行う):
        *   Windows (コマンドプロンプト):
            ```bash
            .venv\Scripts\activate
            ```
        *   Windows (PowerShell):
            ```powershell
            .venv\Scripts\Activate.ps1
            ```
            (PowerShellでスクリプト実行が禁止されている場合は、`Set-ExecutionPolicy RemoteSigned -Scope Process` を実行してから再度試してください)
        *   macOS/Linux (Bash/Zshなど):
            ```bash
            source .venv/bin/activate
            ```
        有効化されると、プロンプトの先頭に `(.venv)` のように表示されます。

4.  **必要なライブラリのインストール**:
    仮想環境が有効化された状態で、以下のコマンドを実行して必要なライブラリをインストールします。
    ```bash
    pip install -r requirements.txt
    ```

5.  **アプリケーションの実行**:
    ```bash
    python main.py
    ```
    または (Pythonのバージョンによっては `python3` を使います):
    ```bash
    python3 main.py
    ```

### 開発者向けの方法 (Gitとuvを使用)

1.  **リポジトリのクローン**:
    ```bash
    git clone <リポジトリのURL>
    cd <リポジトリ名>
    ```

2.  **仮想環境の作成と有効化**:
    お好みの方法で仮想環境を作成し、有効化してください。
    *   **`uv` を使用する場合** (`uv` がインストールされている必要があります):
        ```bash
        uv venv
        ```
    *   **Python標準の `venv` を使用する場合**:
        ```bash
        python -m venv .venv # または python3 -m venv .venv
        ```
    その後、シェルに応じて有効化コマンドを実行します:
    *   Windows (コマンドプロンプト): `.venv\Scripts\activate`
    *   Windows (PowerShell): `.venv\Scripts\Activate.ps1`
    *   macOS/Linux (Bash/Zshなど): `source .venv/bin/activate`

3.  **必要なライブラリのインストール**:
    仮想環境を有効化した後、以下のコマンドで必要なライブラリをインストールします。
    ```bash
    pip install -r requirements.txt
    ```
    (`pyproject.toml` に依存関係を正しく記述し、`uv` を使用している場合は `uv pip install .` も可能です)

4.  **アプリケーションの実行**:
    ```bash
    python main.py # または python3 main.py
    ```
    (`uv` を使用している場合は `uv run python main.py` も可能です)

## 使用方法

1.  アプリケーションを起動すると、GUIが表示されます。
2.  「PowerPointファイル:」の右にある「選択...」ボタンをクリックし、処理したい `.pptx` ファイルを選択します。
3.  「出力先フォルダ:」の右にある「選択...」ボタンをクリックし、抽出したテキストや画像を保存するフォルダを選択します。
4.  「実行」ボタンをクリックすると、処理が開始されます。
5.  処理の進捗状況は、下部の「ステータス:」エリアに表示されます。
6.  処理が完了すると、メッセージが表示されます。

## 注意点

-   ファイルパスに日本語などのマルチバイト文字が含まれている場合、環境によっては問題が発生する可能性があります。なるべく英数字のパスを使用してください。
-   巨大なPowerPointファイルや多数の画像が含まれる場合、処理に時間がかかることがあります。
