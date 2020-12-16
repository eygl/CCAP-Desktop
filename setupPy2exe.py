from distutils.core import setup
import py2exe
 
setup(name='Ccap',
    version='0.60',
    url='about:none',
    author='J. SooHoo',
    package_dir={'Ccap':'.'},
    packages=['Ccap'],
    windows=['Ccap.py']
    )

setup(windows = 
        [
            {
                "script": 'Ccap.py',
                "icon_resources": [(1,"Ccap.ico")]
            }
        ],
)
