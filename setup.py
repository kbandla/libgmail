"""
libgmail

"""
from distutils.core import setup
from libgmail import __version__

setup(
    name = "libgmail",
    author = "Kiran Bandla",
    author_email = "kbandla@in2void.com",
    license = "BSD",
    version = __version__,
    description = "Gmail IMAP Interface in Python",
    url = "http://www.github.com/kbandla/libgmail",
    py_modules=[ 'libgmail' ],
)
