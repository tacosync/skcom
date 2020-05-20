from setuptools import setup, find_packages
import sys
import getopt

# 檢查 --production 參數
# TODO: 這裡也要放行 bdist_wheel 的參數
releaseMode = False
filtered_argv = [sys.argv[0]]
(optlist, args) = getopt.getopt(sys.argv[1:], '', ['production'])

for (name, value) in optlist:
    if name != '--production':
        filtered_argv.append(name)
        filtered_argv.append(value)
    else:
        releaseMode = True

filtered_argv += args
sys.argv = filtered_argv

requiredPkgs = [
    'packaging',
    'requests',
    'PyYAML >= 5.1',
    'busm >= 0.9.4'
]

# production 模式才相依 comtypes, pywin32
# test.pypi.org 沒有這兩個套件
if releaseMode:
    requiredPkgs += [
        'comtypes >= 1.1.7; platform_system=="Windows"',
        'pywin32 >= 1.0; platform_system=="Windows"',
    ]

print(requiredPkgs)
sys.exit(1)

# Load Markdown description.
readme = open('README.md', 'r', encoding='utf-8')
longdesc = readme.read()
readme.close()

# See
# https://packaging.python.org/tutorials/packaging-projects/
# https://python-packaging.readthedocs.io/en/latest/non-code-files.html
setup(
    name='skcom',
    version='0.9.4',
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
    python_requires='>=3.5'
)
