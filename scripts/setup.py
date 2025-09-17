from setuptools import setup
import platform

common_requirements = []
requirements_64bit = []
requirements_32bit = []

is_32bit = platform.architecture()[0] == '32bit'

if is_32bit:
    requirements = requirements_32bit + common_requirements
else:
    requirements = requirements_64bit + common_requirements

setup(
        name='heywoodbess',
        version='0.0.1',
        author='Ben Kearney',
        description='heywood library',
        packages=['heywoodbess'],
        package_dir={'': 'src'},
        install_requires=requirements,
        )
