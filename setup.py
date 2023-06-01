from setuptools import setup
setup(
    name='Stats Orginizer',
    version='1.0',
    description='Orginize daily stats automatically instead of manually',
    long_description='Saving about 5 minutes every day in a year would be 21.67 hours saved per person',
    author='Lazaro Gonzalez',
    author_email='your-email@example.com',
    url='http://your-url.com',
    license='OSI', # Or your preferred license
    entry_points={
    'console_scripts': [
        'my-command=my_package.my_module:my_function',
    ],
    },
    python_requires='>=3.7, <4',
    data_files=[('', ['COPYRIGHT.txt'])],
)
