from setuptools import setup, find_packages

setup(
  name='xlpt',
  version='0.1.0',
  packages=find_packages(),
  include_package_tree=True,
  install_requires=[
    'Click',
    'openpyxl',
  ],
  entry_points={
    'console_scripts': [
      'build = xlpt:xlpt.build',
    ],
  },
)
