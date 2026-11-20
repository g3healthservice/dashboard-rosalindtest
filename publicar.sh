#!/bin/bash
# Regenera a plataforma RosalindTest e publica no GitHub Pages (g3healthservice/dashboard-rosalindtest)
set -e
cd "$(dirname "$0")"
python3 gerar_plataforma_rosalind.py
python3 gerar_modelos_docx.py >/dev/null
cp index.html ~/Downloads/Plataforma_RosalindTest.html
# só o que é público: painel, modelos neutros e cartão de compartilhamento (acervo/anexos ficam locais)
git add -A modelos && git add index.html og.png .nojekyll
git commit -q -m "Atualiza plataforma RosalindTest ($(date +%F' '%H:%M))

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" || echo "nada a commitar"
git push -q origin main
echo "publicado: https://g3healthservice.github.io/dashboard-rosalindtest/"
