from setuptools import setup, find_packages

setup(
    name="image_downloader_pro",
    version="1.0.0",
    description="Professional image downloader with GUI",
    author="Your Name",
    packages=find_packages(),
    install_requires=[
        'pandas>=2.0.0',
        'requests>=2.31.0',
        'Pillow>=10.0.0',
        'openpyxl>=3.1.0',
        'xlrd>=2.0.1'
    ],
    entry_points={
        'console_scripts': [
            'image-downloader=image_downloader:main',
        ],
    },
)