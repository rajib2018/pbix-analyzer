from setuptools import setup, find_packages

setup(
    name='my_package',                  # Package name
    version='0.1',                     # Package version
    description='A sample Python package',  # Short description
    author='John Doe',                 # Author name
    author_email='jdoe@example.com',  # Author email
    packages=find_packages(),          # Automatically find package directories
    install_requires=[                 # Dependencies the package needs
        'numpy',
        'pandas',
        'pbixray'
    ],
)
