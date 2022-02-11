cd release
pyinstaller --noconfirm -F --windowed build-osx.spec
cd dist

mkdir Contents
mkdir TrendyDMG

rm -r TrendyDMG/Trendy.app
mv Trendy.app TrendyDMG/
open -n TrendyDMG/Trendy.app

rm '../Instalador Trendy.dmg'
create-dmg \
  --volname "Instalador Trendy" \
  --window-pos 200 120 \
  --window-size 800 400 \
  --icon-size 100 \
  --icon "Trendy.app" 200 190 \
  --hide-extension "Trendy.app" \
  --app-drop-link 600 185 \
  '../Instalador Trendy.dmg' \
  TrendyDMG
