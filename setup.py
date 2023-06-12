from setuptools import setup, find_packages

setup(
    name='pptxtopdf',
    version='0.0.2',
    author='Victor Ionescu',
    author_email='hlviones@liverpool.ac.uk',
    description='Convert PowerPoint files to PDF',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'pptxtopdf = pptxtopdf.__main__:main'
        ]
    },
    install_requires=[
        'comtypes'  
    ],
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
    ],#
    url='https://github.com/hlviones/pptxtopdf',
)
