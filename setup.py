from setuptools import setup, find_packages
import sys

# 基礎相依套件
requiredPkgs = [
    'packaging',
    'requests',
    'busm >= 0.9.5'
]

# production 才使用的相依套件
# test.pypi.org 沒有這兩個套件
if '--production' in sys.argv:
    at = sys.argv.index('--production')
    sys.argv = sys.argv[0:at] + sys.argv[at+1:]
    requiredPkgs += [
        'PyYAML >= 5.1',
        'cryptography >= 2.9.0',
        'comtypes >= 1.1.7; platform_system=="Windows"',
        'pywin32 >= 1.0; platform_system=="Windows"',
    ]

# Load Markdown description.
readme = open('README.md', 'r', encoding='utf-8')
longdesc = readme.read()
readme.close()

# See
# https://packaging.python.org/tutorials/packaging-projects/
# https://python-packaging.readthedocs.io/en/latest/non-code-files.html
setup(
    name='skcom',
    version='0.9.7',
    description='Get stock informations by Capital API.',
    long_description=longdesc,
    long_description_content_type='text/markdown',
    packages=find_packages(),
    url='https://github.com/tacosync/skcom',
    license='MIT',
    author='Raymond Wu',
    package_data={
        'skcom': ['conf/*', 'samples/*', 'tools/*']
    },
    install_requires=requiredPkgs,
    python_requires='>=3.7'
)
