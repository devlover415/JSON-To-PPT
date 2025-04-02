from setuptools import setup, find_packages

setup(
    name="ppt_updater",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "python-pptx>=0.6.21",
    ],
)