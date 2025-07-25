#!/bin/bash

# Исправляем ссылки в русской версии
sed -i '' '
s|value="../index.html">English|value="../index.html">English|g
s|value="../index.html">Polski|value="../pl/index.html">Polski|g
s|value="../index.html">中文|value="../zh/index.html">中文|g
s|value="../index.html">Español|value="../es/index.html">Español|g
s|value="../index.html">العربية|value="../ar/index.html">العربية|g
s|value="../index.html">Português|value="../pt/index.html">Português|g
s|value="../index.html">日本語|value="../ja/index.html">日本語|g
s|value="../index.html">Deutsch|value="../de/index.html">Deutsch|g
' ru/index.html ru/compare.html

# Исправляем ссылки в польской версии  
sed -i '' '
s|value="../index.html">English|value="../index.html">English|g
s|value="../index.html">Русский|value="../ru/index.html">Русский|g
s|value="../index.html">中文|value="../zh/index.html">中文|g
s|value="../index.html">Español|value="../es/index.html">Español|g
s|value="../index.html">العربية|value="../ar/index.html">العربية|g
s|value="../index.html">Português|value="../pt/index.html">Português|g
s|value="../index.html">日本語|value="../ja/index.html">日本語|g
s|value="../index.html">Deutsch|value="../de/index.html">Deutsch|g
' pl/index.html pl/compare.html

# И так далее для всех языков
for lang in zh es ar pt ja de; do
    sed -i '' "
    s|value=\"../index.html\">English|value=\"../index.html\">English|g
    s|value=\"../index.html\">Русский|value=\"../ru/index.html\">Русский|g
    s|value=\"../index.html\">Polski|value=\"../pl/index.html\">Polski|g
    s|value=\"../index.html\">中文|value=\"../zh/index.html\">中文|g
    s|value=\"../index.html\">Español|value=\"../es/index.html\">Español|g
    s|value=\"../index.html\">العربية|value=\"../ar/index.html\">العربية|g
    s|value=\"../index.html\">Português|value=\"../pt/index.html\">Português|g
    s|value=\"../index.html\">日本語|value=\"../ja/index.html\">日本語|g
    s|value=\"../index.html\">Deutsch|value=\"../de/index.html\">Deutsch|g
    " $lang/index.html $lang/compare.html
done
