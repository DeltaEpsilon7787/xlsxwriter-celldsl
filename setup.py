from setuptools import setup

with open("./README.md", encoding='utf-8') as in_:
    setup(
        name='xlsxwriter-celldsl',
        version='0.1.0',
        packages=['xlsxwriter_celldsl'],
        url='https://github.com/DeltaEpsilon7787/xlsxwriter-celldsl',
        license='MIT',
        author='DeltaEpsilon7787',
        author_email='deltaepsilon7787@gmail.com',
        description='A library to simplify writing procedular generation of Excel files using XlsxWriter and take '
                    'advantage of constant_memory mode easily.',
        description_content_type="text/plain",
        long_description=in_.read(),
        long_description_content_type="text/markdown",
        install_requires=[
            "attrs",
            "pytest-mock",
            "xlsxwriter",
        ]
    )
