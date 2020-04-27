from setuptools import setup, find_packages

# Load Markdown description.
readme = open('README.md', 'r', encoding='utf-8')
longdesc = readme.read()
readme.close()

# See
# https://packaging.python.org/tutorials/packaging-projects/
# https://python-packaging.readthedocs.io/en/latest/non-code-files.html
setup(
    name='skcom',
    version='0.9.4.1',
    description='Get stock informations by Capital API.',
    long_description=longdesc,
    long_description_content_type='text/markdown',
    packages=find_packages(),
    url='https://github.com/tacosync/skcom',
    license='MIT',
    author='Raymond Wu',
    package_data={
        'skcom': ['conf/*', 'samples/*']
    },
    install_requires=[
        'comtypes >= 1.1.7; platform_system=="Windows"',
        'pywin32 >= 1.0; platform_system=="Windows"',
        'packaging',
        'requests',
        'PyYAML >= 5.1',
        'busm >= 0.9.4'
    ],
    platforms=[
        'win_amd64'
    ],
    classifiers=[
        'Operating System :: Microsoft :: Windows :: Windows 10'
    ],
    python_requires='>=3.5'
)
