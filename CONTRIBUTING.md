# Contributing to TeX2Word

Thank you for your interest in contributing to **TeX2Word**!

## Development Environment

TeX2Word targets **64-bit Windows 10 / 11**.

- **Core tests (offline, no Office required)**: Python 3.12 (x64)
- **Full Word + MathType integration tests**: Microsoft Word desktop + MathType 7 desktop
- **Desktop GUI & CLI binaries (`desktop/`)**: Built-in `.NET Framework 4.x` (`csc.exe` + WPF, no extra SDK required)
- **Offline Windows installer (`packaging/windows/build.py`)**: `uv` + Visual Studio Build Tools (x64 C++ Redistributable)

### Setup

```powershell
git clone https://github.com/<your-username>/TeX2Word.git
cd TeX2Word
.\Setup.cmd
```

Or install in editable mode inside a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\pip.exe install -e .
```

## Running Tests

### Offline Unit Tests (fast, no Word/MathType needed)

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

### Full Integration Tests (requires Word + MathType 7)

```powershell
$env:TEX2WORD_OFFICE_TEST = '1'
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

### Installed App Offline End-to-End Test

```powershell
$env:TEX2WORD_INSTALL = Join-Path $env:LOCALAPPDATA 'Programs\TeX2Word'
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## Compiling the Desktop App (`desktop/`)

Both `TeX2Word.exe` (Fluent WPF GUI + `--cli` router) and `TeX2Word.com` (console CLI) must compile cleanly with `/warnaserror+` using the system `.NET Framework 4.x` C# compiler (`csc.exe`). See `.github/workflows/ci.yml` for the exact compiler flags.

## Submitting Pull Requests

1. Fork the repository and create a feature branch.
2. Add or update unit tests in `tests/` covering your changes.
3. Verify offline unit tests and C# compilation pass without warnings.
4. Open a Pull Request describing the motivation, changes, and test results.
