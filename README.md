# Outlook Web Stopka - start

## Pliki
- manifest.xml
- commands.html
- commands.js
- 32x32.png
- 128x128.png

## Ważne
Pliki HTML/JS powinny być hostowane przez GitHub Pages:
https://rafalsieradzki.github.io/outlookwebstopka/

Nie używaj raw.githubusercontent.com jako URL do commands.html w manifeście.

## Kroki
1. Wrzuć pliki do repozytorium: rafalsieradzki/outlookwebstopka
2. W GitHub włącz Pages:
   Settings > Pages > Deploy from branch > main > /root
3. Sprawdź w przeglądarce:
   https://rafalsieradzki.github.io/outlookwebstopka/commands.html
4. Zweryfikuj:
   office-addin-manifest validate manifest.xml
5. Wgraj do tenanta:
   New-App -Organization -Url "https://rafalsieradzki.github.io/outlookwebstopka/manifest.xml"

Na tym etapie dodatek działa ręcznie: w nowej wiadomości kliknij przycisk "Wstaw stopkę".
