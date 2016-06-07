from distutils.core import setup
from setuptools import find_packages


REQUIRES = [
    'numpy==1.11.0',
    'pandas==0.18.1',
]

TRANSITIVE_REQUIRES = [
    'python-dateutil==2.5.3',
    'pytz==2016.4',
    'six==1.10.0',
]

setup(
    name='xlsmodel',
    version='0.0.1',
    description='XLS Model',
    author='Evgeny Pogrebnyak',
    packages=find_packages(),
    zip_safe=False,
    platforms='any',
    install_requires=REQUIRES + TRANSITIVE_REQUIRES,
    extras_require={
        'xlsxwriter': 'XlsxWriter==0.9.0',
        'xlrd': 'xlrd==1.0.0',
        'xlwings': 'xlwings',
    },
)
