from setuptools import find_packages, setup
import os

setup(
  name='RPA 3.0',
  version='3.0.0',
  packages=find_packages(),
  include_package_data=True,
  zip_safe=False,
  install_requires=[
    'flask',
  ],
)
