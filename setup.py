from setuptools import setup

setup(
    name='xlsxwriter-celldsl',
    version='0.0.1',
    packages=['xlsxwriter_celldsl'],
    url='https://github.com/DeltaEpsilon7787/xlsxwriter-celldsl',
    license='MIT',
    author='DeltaEpsilon7787',
    author_email='deltaepsilon7787@gmail.com',
    description='A library to simplify writing procedular generation of Excel files using XlsxWriter and take '
                'advantage of constant_memory mode easily.',
    install_requires=[
        "attrs",
        "pytest-mock",
        "xlsxwriter",
    ]
)
