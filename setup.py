from distutils.core import setup
import py2exe

# setup(console=['passwordGUI.py'])

data_files = [('', [r'flag.ico'])]
setup(
windows =['passwordGUI.py'],
data_files = data_files,
options={
         }
)