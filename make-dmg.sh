rmdir -r dist/dmg
mkdir -p dist/dmg
cp -r "dist/xer-splitter.app" dist/dmg
create-dmg \
--volname "XER SPLITTER" \
--window-pos 200 120 \
--window-size 600 300 \
--icon-size 100 \
--app-drop-link 425 120 \
"dist/xer-splitter.dmg" "dist/dmg"
