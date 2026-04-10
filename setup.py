from setuptools import setup, find_packages

with open("README_EN.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pyzwcadmech",
    version="0.1.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="A Python wrapper for ZWCAD Mechanical COM API",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/pyzwcadmech",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires=">=3.6",
    install_requires=[
        "comtypes>=1.1.0",
    ],
)
