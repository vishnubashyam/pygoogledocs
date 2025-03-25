from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pygoogledocs",
    version="0.1.0",
    author="Vishnu Bashyam",
    author_email="example@example.com",
    description="A Python package for automating Google Docs operations",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/pygoogledocs",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    install_requires=[
        "google-api-python-client>=2.0.0",
        "google-auth>=2.0.0",
        "google-auth-oauthlib>=0.4.0",
        "google-auth-httplib2>=0.1.0",
    ],
    extras_require={
        "dev": [
            "pytest>=6.0.0",
            "black>=21.5b2",
            "isort>=5.9.1",
            "mypy>=0.812",
            "flake8>=3.9.2",
        ],
        "streamlit": [
            "streamlit>=1.0.0",
        ],
    },
)