
import os
from pathlib import Path
import fnmatch
import glob
import collections

c_pf_application = '*PowerFactory.exe'

def get_pf_locations():
	"""
		Finds all of the locations that PowerFactory is installed in to obtain the various version details
	:return collections.OrderedDict dict_paths:  Ordered Dictonary of {year and version: path}
	"""
	dict_paths = collections.OrderedDict()

	# TODO: Add function to confirm where exactly to look
	src_path = 'C:\\Program Files'

	for root, dirnames, filenames in os.walk(src_path):
		for _ in fnmatch.filter(filenames, c_pf_application):
			version_path = root
			version = os.path.basename(version_path).replace('PowerFactory ', '')

			dict_paths[version] = version_path


	return dict_paths

if __name__ == '__main__':
	pf_install = get_pf_locations()
	print(pf_install)
