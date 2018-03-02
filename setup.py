from distutils.core import setup
import py2exe
import matplotlib

setup(
    console=['upto_tabs_b.py'],
    data_files=matplotlib.get_py2exe_datafiles(),
    options={'py2exe': {
            "includes" : ["matplotlib.backends.backend_wxagg"],
            'excludes': ['_gtkagg','_tkagg'],
            }}
)
