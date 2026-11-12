# Relato de Ocorrência

## Publicação na Streamlit Community Cloud

1. Crie um repositório no GitHub e envie `app.py`, `requirements.txt`, `.gitignore`, `.streamlit/config.toml` e `.streamlit/secrets.example.toml`.
2. **Não envie** `.streamlit/secrets.toml`, `dados_local.csv`, `base_prestadores.csv`, `qr_acesso_celular.png` nem as pastas de ambiente virtual.
3. Use a planilha Google informada pelo usuário. O app está configurado para a aba `Registro`; como ela está vazia, o app criará os cabeçalhos automaticamente no primeiro acesso com as credenciais corretas. Crie uma conta de serviço no Google Cloud, habilite as APIs Google Sheets e Drive, e compartilhe essa planilha com o e-mail da conta de serviço como Editor.
4. No Streamlit Community Cloud, crie o app apontando para `app.py`. Em **Advanced settings / Secrets**, configure os campos do exemplo:
   - use uma senha forte, diferente da senha local;
   - coloque o link CSV filtrado da base em `data.sheet_url`;
   - use `storage.backend = "google_sheets"`, o ID já indicado no exemplo e `worksheet = "Registro"`;
   - copie os campos da chave JSON da conta de serviço em `gcp_service_account`;
   - depois de receber a URL pública, preencha `app.public_url` e reinicie o app.

Nunca coloque a chave privada ou `secrets.toml` no GitHub. A senha local atual fica apenas no arquivo ignorado `.streamlit/secrets.toml`.

O histórico existente em `dados_local.csv` permanece local e não é importado automaticamente para a nova planilha.

Streamlit Community Cloud não oferece garantia de disponibilidade contínua 24 horas. Se isso for requisito obrigatório, use hospedagem com plano de disponibilidade apropriado.
