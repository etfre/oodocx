#!/usr/bin/env python
""" Setup """

from setuptools import setup


setup(
    name="oodocx",
    version="0.1.0",
    packages=['oodocx'],
    include_package_data=True,
    install_requires=['lxml'],
    # metadata for upload to PyPI
    author="Evan Fredericksen",
    author_email="evfredericksen@gmail.com",
    description="Load Docx in memory.",
    license="PSF",
    keywords="Docx Microsoft Word",
    url='http://github.com/evfredericksen/oodocx'
)
