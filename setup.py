from setuptools import setup

setup(
    version="0.2",
    name='py-xlsx',
    description="""Tiny python code for parsing data from Microsoft's Office
    Open XML Spreadsheet format""",
    long_description="",
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Programming Language :: Python',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
    ],
    author='Staale Undheim',
    author_email='staale@staale.org',
    url='http://github.com/staale/python-xlsx',
    packages=[
        "xlsx"
    ],
    test_suite = 'xlsx.tests'
)
