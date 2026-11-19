# AutoTransfere — distribuição pública

Repositório: https://github.com/SirSant/auto_transfere

## Como funciona
- O aplicativo consulta a API pública de GitHub Releases ao abrir.
- Quando encontra uma versão mais nova, oferece baixar `AutoTransfere_Setup.exe`.
- O instalador do Inno Setup atualiza a instalação anterior usando o mesmo AppId.
- O build de cada versão é executado no GitHub Actions para Windows.

## Publicar a primeira versão
Após enviar os arquivos para `main`, crie e envie uma tag:

```powershell
git add .
git commit -m "Prepare AutoTransfere installer and updater"
git push origin main
git tag v1.0.0
git push origin v1.0.0
```

O workflow `.github/workflows/release.yml` compilará o aplicativo e publicará `AutoTransfere_Setup.exe` na release.

## Publicar atualizações futuras
1. Altere o código e envie o commit para `main`.
2. Crie uma nova tag, por exemplo `v1.0.1`, e envie-a.
3. O workflow atualizará a versão no executável e no instalador, compilará e publicará a nova release.
4. Na próxima abertura, os usuários da versão anterior receberão a oferta para atualizar.

## OCR
O workflow instala o Tesseract no ambiente de build e inclui o modelo de português. Isso evita exigir Python, pip ou instalação separada do OCR nos computadores dos usuários. Preserve os avisos e licenças dos componentes redistribuídos.

## Antes de distribuir
Teste o instalador em um Windows limpo, sem Python e sem Tesseract. Valide abertura, OCR, automação, desinstalação e atualização sobre uma versão anterior.
