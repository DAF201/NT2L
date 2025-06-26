import setuptools

setuptools.setup(
    name="nt2l",
    version="2.0.1",
    description="NEIS_TOOLS_2 Lite",
    author="NEI",
    author_email="daf201@blink-in.com",
    install_requires=[
        "pywin32",
        "beautifulsoup4",
        "requests",
        "psutil",
    ],
    packages=[
        "nt2",
        "core",
        "config",
        "export",
        "service",
        "functions",
    ],
    include_package_data=True,
    package_data={
        "config": ["config.json"],
    },
    entry_points={
        "console_scripts": [
            "nt2=nt2.main:main",
        ],
    },
    python_requires=">=3.11",
)
