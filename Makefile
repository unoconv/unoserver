root_dir := $(shell dirname $(realpath $(lastword $(MAKEFILE_LIST))))
bin_dir := $(root_dir)/ve/bin

all: check coverage

# The fullrelease script is a part of zest.releaser, which is the last
# package installed, so if it exists, the devenv is installed.
devenv:	ve/bin/fullrelease

ve/bin/fullrelease:
	virtualenv $(root_dir)/ve --python python3 --system-site-packages
	$(bin_dir)/pip install -e .[devenv]

check: devenv
	$(bin_dir)/black src/unoserver tests
	$(bin_dir)/flake8 src/unoserver tests
	$(bin_dir)/pyroma -d .
	$(bin_dir)/check-manifest


coverage: devenv
	$(bin_dir)/coverage run -m unittest
	$(bin_dir)/coverage html
	$(bin_dir)/coverage report

test: devenv
	PATH=$(bin_dir):$$PATH $(bin_dir)/pytest

release: devenv
	$(bin_dir)/fullrelease


clean:
	rm -rf ve build htmlcov
