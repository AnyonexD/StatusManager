# Publicacao de nova versao

1. Atualize `APP_VERSION` em `status_manager.py`.
2. Compile usando o `StatusManager.spec`.
3. No GitHub, crie uma Release com tag no formato `v2.1.1`.
4. Anexe o executavel com o nome exato `StatusManager.exe`.
5. Publique a Release.

O aplicativo consulta a ultima Release do GitHub e baixa o asset `StatusManager.exe` quando a tag publicada for maior que a versao local.
