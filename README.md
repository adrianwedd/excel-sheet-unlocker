# 🔒 Excel Sheet Unlocker 🔒

Welcome to Excel Sheet Unlocker, a Python program that can unlock cells and dropdowns in an Excel sheet and apply password protection to it. 📖🔐

## 🚀 Getting Started 🚀

These instructions will help you set up Excel Sheet Unlocker on your local machine for development and testing purposes. 

### 📋 Prerequisites 📋

- Python 3.7 or (ideally) later
- pip (Python package installer)
- A clone of this repository on your local machine

### 🛠️ Setup Instructions 🛠️

1. Clone the repository to your local machine:

```bash
git clone https://github.com/adrianwedd/excel-sheet-unlocker.git
```

2. Navigate to the repository's directory:

```bash
cd excel-sheet-unlocker
```

3. Create a Python virtual environment:

```bash
python -m venv env
```

4. Activate the virtual environment:

```bash
# On Windows
env\Scripts\activate

# On Unix or MacOS
source env/bin/activate
```

5. Install the required Python packages:

```bash
pip install -r requirements.txt
```

Now, you're all set to run Excel Sheet Unlocker!

## 🏃 Usage 🏃

There are two ways to use Excel Sheet Unlocker: interactive mode and command-line arguments mode.

### 🗣️ Interactive Mode 🗣️

In interactive mode, you run the script without any command-line arguments, and it will prompt you to provide the necessary inputs interactively.

```bash
python main.py
```

### 🎛️ Command-Line Arguments Mode 🎛️

Alternatively, you can provide the inputs as command-line arguments when you run the script.

```bash
python main.py --input_file /path/to/input.xlsx --output_file /path/to/output.xlsx --sheet_name Sheet1 --password mypassword --progress_bar
```

Note: Be aware that using the `--password` argument will expose your password in your shell history. Please consider your security requirements.

## 📚 Documentation 📚

Please refer to the comments in `main.py` for a detailed explanation of the code and how it works.

## 🤝 Contributing 🤝

Please read `CONTRIBUTING.md` for details on our code of conduct, and the process for submitting pull requests to us.

## 📃 License 📃

This project is licensed under the MIT License - see the `LICENSE.md` file for details.

## ✨ Acknowledgments ✨

- OpenAI's ChatGPT for assistance in developing the code.
- The openpyxl library for enabling the manipulation of Excel files in Python.
- The tqdm library for enabling the progress bar.
