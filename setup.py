from setuptools import setup

requires = [
    'flask',
    'gevent',
    'gunicorn',
    'lxml',
    'pillow',
    'python-dotenv',
    'python-pptx',
    'xlsxwriter',
]

setup(
    name='poc',
    install_requires=requires,
    entry_points={
        'paste.app_factory': [
            'main = poc:main'
        ],
    },
)
