#!/usr/bin/env python

from setuptools import setup, find_packages
setup(
    name = "oodocx",
    version = "0.1.0",
    packages = find_packages(),
    include_package_data = True,
    install_requires = ['lxml'],
    # metadata for upload to PyPI
    author = "Evan Fredericksen",
    author_email = "me@example.com",
    description = "This is an Example Package",
    license = "PSF",
    keywords = "Docx Microsoft Word",
    url = 'http://github.com/evfredericksen/oodocx'
)