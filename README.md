# recruitment-ai-assistant

## Prerequisites

1. Install the latest Python 3 version from the Microsoft Store:
   https://apps.microsoft.com/detail/9nq7512cxl7t?hl=ro-RO&gl=RO
2. The MarkItDown source is available at:
   https://github.com/microsoft/markitdown

## Notes

- This repository uses Python-based MarkItDown.

## How to use `prompts/candidate_processor.prompt.md`

1. Create one candidate folder inside `input/` with the candidate name.
   - Example: `input/Alupei Gheorghe/`
   - Add the candidate CV PDF and the interview notes file:
     - `input/Alupei Gheorghe/Alupei Gheorghe.pdf`
     - `input/Alupei Gheorghe/interview_notes.txt`

2. Install Python 3 and MarkItDown as described below.

3. In chat, tell the assistant exactly what to do using the prompt file.
   - Example message: `Use prompts/candidate_processor.prompt.md for the folder input/Alupei Gheorghe.`

   - The assistant should then:
     - read `input/Alupei Gheorghe/Alupei Gheorghe.pdf`
     - read `input/Alupei Gheorghe/interview_notes.txt`
     - generate `input/Alupei Gheorghe/Alupei Gheorghe.md`
     - generate `output/Alupei Gheorghe/candidate_presentation_email.txt`
     - generate `output/Alupei Gheorghe/Alupei Gheorghe.md`
     - try to generate `output/Alupei Gheorghe/Alupei Gheorghe.pptx` if possible

## Install MarkItDown locally

After installing Python 3, open PowerShell in the repository root and run:

```powershell
py -3 -m pip install --upgrade pip
py -3 -m pip install --user "markitdown[all]"
py -3 -m pip install --user "markitdown[pptx]"
```

If the `markitdown` script is not available directly, use it through Python:

```powershell
py -3 -m markitdown --help
```

If needed, add the Python user scripts folder to your PATH. Then you can run MarkItDown as shown below.

## Setup

Run the installer script after Python is installed:

```powershell
cd .\scripts
.\install_prerequisites.ps1
```

## Usage

After prerequisites are installed, convert a PDF to Markdown with:

```powershell
py -3 -m markitdown "input\Vlad Gheorghe Branzei\GHEORGHE VLAD BRANZEI.pdf" -o "input\Vlad Gheorghe Branzei\GHEORGHE VLAD BRANZEI.md"
```
After generating the candidate Markdown, use the prompt instructions to create the email and internal CV.

For PowerPoint generation, run the template-based script:

```powershell
cd scripts
python generate_pptx_from_template.py
```
If you no longer want the npm package tree, remove `node_modules` and ignore `package-lock.json`.
