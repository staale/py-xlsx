from setuptools import setup

setup(
    version="0.3",
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
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.2',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
    ],
    author='Staale Undheim',
    author_email='staale@staale.org',
    url='http://github.com/staale/python-xlsx',
    tests_require = ['six'],
    packages=[
        "xlsx"
    ],
    test_suite = 'xlsx.tests'
)
