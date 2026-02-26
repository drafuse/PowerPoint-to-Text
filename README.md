# PowerPoint to Text

Extract slide text from a `.pptx` file into a `.txt` file.
The EXE opens a desktop window where you can choose input and output paths.

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

## Run the EXE (GUI)

```powershell
.\dist\pptx_to_text.exe
```

Then:
1. Click `Browse...` and select your `.pptx`
2. Click `Save As...` and choose output `.txt` name/location
3. Click `Convert`

## Optional CLI mode

```powershell
python .\pptx_to_text.py "C:\path\to\slides.pptx"
python .\pptx_to_text.py "C:\path\to\slides.pptx" -o "C:\path\to\output.txt"
```
