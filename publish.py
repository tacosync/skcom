# rstcheck README.rst
# pylint -f colorized -E twnews/*.py
# python setup.py -q bdist_wheel
# 佈署到 pypi.org
# twine upload dist/$WHEEL
# 佈署到 test.pypi.org
# twine upload --repository testpypi --verbose dist/skcom
