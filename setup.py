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
    install_requires=[
        # 注意! 這兩個套件不存在於 test.pypi.org
        # 安裝測試版套件前需要預先裝好這兩項
        # 安裝正式版套件則不會發生問題
        'comtypes >= 1.1.7; platform_system=="Windows"',
        'pywin32 >= 1.0; platform_system=="Windows"',
        # 其他常規相依套件
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
