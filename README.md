# PowerPoint to Text

Extract slide text from a `.pptx` file into a `.txt` file.

## Run with Python

```powershell
python .\pptx_to_text.py "C:\path\to\slides.pptx"
python .\pptx_to_text.py "C:\path\to\slides.pptx" -o "C:\path\to\output.txt"
```

## Build Windows EXE (Standalone)

```powershell
.\build_exe.ps1
```

Output executable:

```text
.\dist\pptx_to_text.exe
```

## Run the EXE

```powershell
.\dist\pptx_to_text.exe "C:\path\to\slides.pptx"
.\dist\pptx_to_text.exe "C:\path\to\slides.pptx" -o "C:\path\to\output.txt"
```
