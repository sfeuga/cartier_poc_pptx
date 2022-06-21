from setuptools import setup

requires = [
    'gevent',
]

setup(
    name='cartier_poc',
    install_requires=requires,
    entry_points={
        'paste.app_factory': [
            'main = cartier_poc:main'
        ],
    },
)
