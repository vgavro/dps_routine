from IPython import get_ipython


ip = get_ipython()
ip.run_line_magic('load_ext', 'autoreload')
ip.run_line_magic('autoreload', '2')  # smart autoreload modules on changes

import sfs_cabinet
import os
