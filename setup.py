from setuptools import setup, find_packages

setup(
    name='OutlookHandler',
    version='1.0',
    description='A wrapper win32com for handling Outlook',
    author='Oliver Mohr',
    author_email='oliver.mohr.b@gmail.com',
    packages=find_packages(where='src')
)