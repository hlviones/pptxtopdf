from setuptools import setup, find_packages

setup(
    name='pptxtopdf',
    version='1.0',
    author='Victor Ionescu',
    author_email='hlviones@liverpool.ac.uk',
    description='Convert PowerPoint files to PDF',
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'pptxtopdf = pptxtopdf.__main__:main'
        ]
    },
    install_requires=[
        'comtypes'  
    ],
)
