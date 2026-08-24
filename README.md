# HERMES

Painel web e desktop para conferências contábeis, financeiras, fiscais e operacionais a partir de Excel, CSV e XML.

## Documentação

- [Manual do Usuário](docs/GUIA_USUARIO_HERMES.md): procedimento de uso, arquivos exigidos, interpretação dos resultados, exportações e solução de erros.
- [Documentação Técnica](docs/DOCUMENTACAO_TECNICA.md): arquitetura, módulos, regras de negócio, execução, Docker, segurança, testes e manutenção.

## Automações

1. Conciliação de Receita.
2. Conciliação da Receita de Diárias.
3. Lançamento da Folha de Pagamento.
4. Cupons Emitidos x Conta do Hóspede.
5. RPS de Serviços Prestados.
6. Relatório de Notas de Débito.
7. Notas Fiscais de Entrada em Atraso.
8. Conferência dos Cupons.
9. Notas de Serviços Tomados.
10. Conferência do Contas a Receber.
11. Conferência do Contas a Pagar.
12. Custos da Mercadoria Vendida — CMV.

## Requisitos

- Python 3.12 ou superior;
- `uv` para gerenciamento do ambiente;
- Docker Desktop para execução em contêiner.

## Execução local

```powershell
$env:UV_CACHE_DIR = "$PWD\.uv-cache"
uv sync
uv run python main.py
```

## Execução com Docker

```powershell
docker compose build
docker compose up -d
docker compose ps
```

O endereço padrão é `http://localhost`. O Nginx publica a porta configurada por
`NGINX_PORT` e encaminha as conexões para o HERMES pela rede interna do Docker;
a porta `8000` da aplicação não é publicada no computador.

Na instalação atual, os computadores autorizados acessam
`http://10.197.0.127`. A regra da porta 80 deve ser liberada no Firewall do
Windows somente para a rede administrativa `10.197.0.0/22`, conforme a
[Documentação Técnica](docs/DOCUMENTACAO_TECNICA.md#92-docker).

O HERMES exibe uma tela de login e valida a credencial pelo arquivo local
`nginx/.htpasswd`. O arquivo armazena somente o hash bcrypt, não é versionado
pelo Git e é montado no contêiner apenas para leitura. A autenticação é apagada
quando a página perde a conexão, exigindo novo login no próximo acesso.

Para encerrar:

```powershell
docker compose down
```

## Testes

```powershell
$env:UV_CACHE_DIR = "$PWD\.uv-cache"
uv run python -m unittest discover -s tests -v
```

## Estrutura

```text
hermes-main/
├── automations/                         # Leitores, regras e exportadores
├── assets/                              # Identidade visual e modelo da folha
├── docs/
│   ├── DOCUMENTACAO_TECNICA.md
│   └── GUIA_USUARIO_HERMES.md
├── hermes_ui/                           # Interface, catálogo e runtime
├── nginx/                               # Configuração do proxy reverso
├── tests/                               # Testes automatizados
├── compose.yaml
├── Dockerfile
├── main.py
├── pyproject.toml
└── uv.lock
```

## Dados e arquivos

O HERMES não utiliza banco de dados nem mantém histórico das planilhas analisadas. Na execução web, uploads e exportações usam diretórios temporários. Nas atividades que inserem as planilhas-base no Excel final, o último conjunto permanece apenas na memória da sessão até ser limpo, substituído ou a sessão terminar.

Consulte a [Documentação Técnica](docs/DOCUMENTACAO_TECNICA.md#5-ciclo-de-vida-dos-arquivos) para o fluxo completo.
