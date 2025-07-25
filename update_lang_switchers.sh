#!/bin/bash

for lang in ru pl zh es ar pt ja de fr; do
  if [ -d "$lang" ]; then
    echo "Updating $lang..."
    
    # Обновляем главную страницу языка
    if [ -f "$lang/index.html" ]; then
      sed -i '' \
        -e 's|value="../index.html">English|value="/">English|g' \
        -e "s|value=\"index.html\" selected>|value=\"/$lang/\" selected>|g" \
        -e 's|value="../ru/index.html">Русский|value="/ru/">Русский|g' \
        -e 's|value="../pl/index.html">Polski|value="/pl/">Polski|g' \
        -e 's|value="../zh/index.html">中文|value="/zh/">中文|g' \
        -e 's|value="../es/index.html">Español|value="/es/">Español|g' \
        -e 's|value="../ar/index.html">العربية|value="/ar/">العربية|g' \
        -e 's|value="../pt/index.html">Português|value="/pt/">Português|g' \
        -e 's|value="../ja/index.html">日本語|value="/ja/">日本語|g' \
        -e 's|value="../de/index.html">Deutsch|value="/de/">Deutsch|g' \
        "$lang/index.html"
    fi
    
    # Обновляем страницу compare языка
    if [ -f "$lang/compare/index.html" ]; then
      sed -i '' \
        -e 's|value="../../compare/">English|value="/compare/">English|g' \
        -e "s|value=\"compare/\" selected>|value=\"/$lang/compare/\" selected>|g" \
        -e 's|value="../ru/compare/">Русский|value="/ru/compare/">Русский|g' \
        -e 's|value="../pl/compare/">Polski|value="/pl/compare/">Polski|g' \
        -e 's|value="../zh/compare/">中文|value="/zh/compare/">中文|g' \
        -e 's|value="../es/compare/">Español|value="/es/compare/">Español|g' \
        -e 's|value="../ar/compare/">العربية|value="/ar/compare/">العربية|g' \
        -e 's|value="../pt/compare/">Português|value="/pt/compare/">Português|g' \
        -e 's|value="../ja/compare/">日本語|value="/ja/compare/">日本語|g' \
        -e 's|value="../de/compare/">Deutsch|value="/de/compare/">Deutsch|g' \
        "$lang/compare/index.html"
    fi
  fi
done

echo "All language switchers updated!"
