# LaTeX-to-Image Converter

A simple tool to convert LaTeX equations into images that can be easily pasted into applications like Word, OneNote, or PowerPoint. It can also convert the LaTeX into plain text.

![Tool Screenshot](https://i.imgur.com/your_screenshot.png) <!-- Placeholder for a future screenshot -->

## Quick Start (No Programming Knowledge Needed)

Follow these steps to get the tool running on your Windows computer.

### Step 1: Install Python

If you don't have Python, you need to install it first.

1.  Go to the official Python website: [python.org/downloads/windows](https://www.python.org/downloads/windows/)
2.  Download the latest stable version (e.g., Python 3.12).
3.  Run the installer. **Important:** On the first screen of the installer, make sure to check the box that says **"Add Python to PATH"**. This will make the next steps much easier.

![Add Python to PATH](https://i.imgur.com/add_python_to_path.png) <!-- Placeholder -->

### Step 2: Download This Tool

1.  Click the green **"Code"** button at the top of this page.
2.  Select **"Download ZIP"**.
3.  Extract the ZIP file to a folder you can easily find, for example, `C:\Users\YourUser\Documents\LatexTool`.

### Step 3: Install Required Libraries

1.  Open the **Command Prompt**. You can find it by searching for "cmd" in the Start Menu.
2.  Navigate to the folder where you extracted the tool. For example:
    ```cmd
    cd C:\Users\YourUser\Documents\LatexTool
    ```
3.  Install the necessary Python libraries by running this command:
    ```cmd
    pip install pillow matplotlib ttkthemes pywin32
    ```
    *On non-Windows systems the `pywin32` dependency is optional and can be
    omitted.*

### Step 4: Run the Application

Now, you can run the tool. Simply double-click the `latexclip.py` file, or run it from the Command Prompt:

```cmd
python latexclip.py
```

The application window will open, and you can start converting your LaTeX!

---

## Optional Upgrades

### For Full, High-Quality LaTeX Rendering

The default mode uses a built-in renderer that is fast but may not support all complex LaTeX packages. For publication-quality rendering, you can install a full LaTeX distribution.

**1. Install a LaTeX Distribution**

You only need one of the following:

*   **MiKTeX (Recommended for Windows):** It's free and automatically downloads packages as you need them, saving space.
    *   [**Download MiKTeX**](https://miktex.org/download)

*   **TeX Live (Alternative):** A larger, more comprehensive distribution.
    *   [**Download TeX Live**](https://www.tug.org/texlive/acquire-netinstall.html)

**2. Install Ghostscript**

This is a required companion for the LaTeX distribution.

*   [**Download Ghostscript**](https://ghostscript.com/releases/gsdnld.html)
    *   Make sure the installer adds Ghostscript to your system's PATH (it usually does this by default).

Once installed, simply check the **"Use full LaTeX"** box in the app to enable high-quality rendering.

### Create a Standalone `.exe` Application

If you want to run this tool without needing to open Command Prompt, you can bundle it into a single `.exe` file.

**1. Install PyInstaller**

Open Command Prompt and run:

```cmd
pip install pyinstaller
```

**2. Create the Executable**

In the Command Prompt, navigate to the tool's directory and run the following command:

Use the bundled spec file so PyInstaller collects the themed assets and
matplotlib backends automatically:

```cmd
pyinstaller latexclip.spec
```

After a few moments, you will find a `dist` folder. Inside, `latexclip.exe` is your standalone application. You can move this file anywhere on your computer or create a shortcut to it on your desktop.
