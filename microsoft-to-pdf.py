from sys import argv
from os import walk
from os.path import join
from subprocess import CompletedProcess, run, CalledProcessError


def is_microsoft(file: str) -> bool:
    return (
        file.endswith(".doc")
        or file.endswith(".docx")
        or file.endswith(".ppt")
        or file.endswith(".pptx")
    )


def convert_to_pdf(file: str) -> CompletedProcess[bytes]:
    return run(["pandoc", "-s", file, "-o", f"{file}.pdf"])


def process_directory(directory: str):
    for root, _, files in walk(directory):
        for file in files:
            if is_microsoft(file):
                full_path = join(root, file)
                convert_to_pdf(full_path)


def main():
    try:
        run(["pandoc", "--version"], check=True)
    except CalledProcessError:
        raise ValueError("Pandoc is not installed")
    
    if len(argv) != 2:
        raise ValueError("Usage: python convert_microsoft_to_pdf.py dir")

    process_directory(argv[1])


if __name__ == "__main__":
    main()
