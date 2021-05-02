from setuptools import setup

with open("./README.md", encoding='utf-8') as in_:
    setup(
        name='xlsxwriter-celldsl',
        version='0.2.0',
        packages=['xlsxwriter_celldsl'],
        url='https://github.com/DeltaEpsilon7787/xlsxwriter-celldsl',
        license='MIT',
        author='DeltaEpsilon7787',
        author_email='deltaepsilon7787@gmail.com',
        description='A library to write scripts for generating Excel files using XlsxWriter in a more structured '
                    'manner via a DSL and taking advantage of constant_memory mode easily.',
        long_description=in_.read(),
        long_description_content_type="text/markdown",
        install_requires=[
            "attrs",
            "xlsxwriter",
        ],
    )
