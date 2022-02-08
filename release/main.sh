cd release
pyinstaller --noconfirm -F --windowed build-osx.spec
open -n dist/Trendy.app
mkdir dist/Contents
