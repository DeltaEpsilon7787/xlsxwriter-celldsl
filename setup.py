from setuptools import find_packages, setup

with open("./README.md", encoding='utf-8') as in_:
    setup(
        name='xlsxwriter-celldsl',
        version='0.6.0',
        packages=find_packages(where='src'),
        package_dir={
            "": "src"
        },
        url='https://github.com/DeltaEpsilon7787/xlsxwriter-celldsl',
        license='MIT',
        author='DeltaEpsilon7787',
        author_email='deltaepsilon7787@gmail.com',
        description='A library to write scripts for generating Excel files using XlsxWriter in a more structured '
                    'manner by avoiding dealing with absolute coordinates.',
        long_description=in_.read(),
        long_description_content_type="text/markdown",
        python_requires='>=3.6',
        install_requires=[
            "attrs",
            "xlsxwriter",
        ],
        extras_require={
            'testing': ['pytest', 'pytest-mock']
        },
    )
