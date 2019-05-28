import subprocess
import re

# TODO:
'''
print('Lint *.py files.')
cmd = 'pylint -f colorized ...'
complete = subprocess.run(cmd, stdout=subprocess.PIPE)
if complete.returncode != 0:
    exit(1)
'''

print('Check README.rst.')
cmd = 'rstcheck README.rst'
complete = subprocess.run(cmd, stdout=subprocess.PIPE)
if complete.returncode != 0:
    exit(1)

print('Build wheel.')
cmd = 'python setup.py bdist_wheel --plat-name win_amd64'
complete = subprocess.run(cmd, stdout=subprocess.PIPE)
if complete.returncode != 0:
    exit(1)

wheel = ''
out_lines = complete.stdout.decode('cp950').split('\r\n')
for line in out_lines:
    m = re.search(r'dist\\skcom-.+-win_amd64.whl', line)
    if m is not None:
        (off_beg, off_end) = m.span(0)
        wheel = line[off_beg:off_end]
        break

if wheel == '':
    exit(1)

print('Upload wheel file: %s.' % wheel)
cmd = 'twine upload --repository testpypi --verbose ' + wheel
complete = subprocess.run(cmd)
if complete.returncode != 0:
    exit(1)
