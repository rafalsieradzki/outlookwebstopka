# Stopka Familijna - TaskPane

Podmień na GitHubie:
- manifest.xml
- taskpane.html
- taskpane.js

Zostaw:
- 32x32.png
- 128x128.png

Sprawdź:
https://rafalsieradzki.github.io/outlookwebstopka/taskpane.html

Walidacja:
office-addin-manifest validate manifest.xml

Aktualizacja dodatku:
Remove-App -Organization -Identity "Stopka Familijna"
New-App -Organization -Url "https://rafalsieradzki.github.io/outlookwebstopka/manifest.xml"
