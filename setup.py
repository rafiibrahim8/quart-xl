from setuptools import setup, find_packages
from gddl import __version__

def read_file(filename):
    try:
        with open(filename, encoding='utf-8') as f:
            return f.read()
    except:
        return []

requirements = read_file('requirements.txt')
long_description = read_file('README.md')

setup(
    name='quart-xl',
    version=__version__,
    
    author='Ibrahim Rafi',
    author_email='me@ibrahimrafi.me',

    license='MIT',

    url='https://github.com/rafiibrahim8/quart-xl',
    download_url = 'https://github.com/rafiibrahim8/quart-xl/archive/v{}.tar.gz'.format(__version__),

    install_requires=requirements.strip().split('\n'),
    description='Convert Genarated Quartus Prime report into Excell.',
    long_description=long_description,
    long_description_content_type='text/markdown',
    keywords=['quart-xl', 'Quartus Prime', 'Report', 'Excell'],

    packages= find_packages(),
    entry_points=dict(
        console_scripts=[
            'quart-xl=quart_xl.quart_xl:main'
        ]
    ),

    platforms=['any'],
    classifiers=[
    'Development Status :: 3 - Alpha',
    'Intended Audience :: End Users/Desktop',
    'Topic :: Utilities',
    'License :: OSI Approved :: MIT License',
    'Programming Language :: Python :: 3',
    'Programming Language :: Python :: 3.4',
    'Programming Language :: Python :: 3.5',
    'Programming Language :: Python :: 3.6',
    'Programming Language :: Python :: 3.7',
    'Programming Language :: Python :: 3.8',
    'Programming Language :: Python :: 3.9',
  ],
)
