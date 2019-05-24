from setuptools import setup, find_packages

# Load reStructedText description.
# Online Editor   - http://rst.ninjs.org/
# Quick Reference - http://docutils.sourceforge.net/docs/user/rst/quickref.html
readme = open('README.rst', 'r', encoding='utf-8')
longdesc = readme.read()
readme.close()

# See
# https://packaging.python.org/tutorials/packaging-projects/
# https://python-packaging.readthedocs.io/en/latest/non-code-files.html
setup(
    name='skcom',
    version='0.1.3',
    description='群益證券聽牌套件',
    long_description=longdesc,
    packages=find_packages(),
    url='https://github.com/virus-warnning/skcom',
    license='MIT',
    author='Raymond Wu',
    package_data={
        'twnews': ['conf/*', 'samples/*']
    },
    install_requires=[
        'comtypes >= 1.1.7; platform_system=="Windows"',
        'pywin32 >= 1.0; platform_system=="Windows"',
        'packaging',
        'requests'
    ],
    classifiers=[
        'Operating System :: Microsoft :: Windows :: Windows 10'
    ],
    python_requires='>=3.5'
)
