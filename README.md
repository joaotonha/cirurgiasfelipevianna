# Fila Cirúrgica ORL

Aplicativo web em Google Apps Script para acompanhamento de pacientes com indicação cirúrgica em otorrinolaringologia. O sistema usa Google Sheets como base de dados privada e oferece uma interface própria em cards para a equipe acompanhar cobranças, exames, solicitações ao plano, autorizações e agendamentos.

## Objetivo

Substituir a edição direta de uma planilha por uma interface simples, visual e responsiva, mantendo os dados reais dos pacientes apenas em uma planilha Google privada.

## Arquivos do projeto

Para simplificar a manutenção, o projeto foi organizado com poucos arquivos:

- `Code.gs`: todo o backend do Apps Script, incluindo planilha, pacientes, usuários, histórico, WhatsApp, hospitais, planos e configurações.
- `Index.html`: toda a interface do aplicativo, incluindo HTML, CSS e JavaScript.
- `appsscript.json`: manifesto do Apps Script.
- `README.md`: documentação.

Na prática, os arquivos que você provavelmente vai editar são apenas `Code.gs` e `Index.html`.

## Como criar a planilha

1. Crie uma planilha Google privada.
2. Abra `Extensões > Apps Script`.
3. Copie os arquivos deste repositório para o projeto Apps Script.
4. Execute a função `inicializarPlanilha()` uma vez.

A função cria as abas:

- `Pacientes`
- `Hospitais`
- `Usuários`
- `Histórico`
- `ModelosWhatsApp`
- `PlanosDeSaúde`
- `Configurações`

Ela também cadastra os hospitais, usuários, modelos de WhatsApp e configurações iniciais.

## Onde o aplicativo fica hospedado

Existem duas formas de publicar:

1. **Mais simples:** publicar diretamente pelo Google Apps Script como `Aplicativo da Web`.
2. **Com link do GitHub:** publicar uma página no GitHub Pages que abre o Web App do Apps Script dentro dela.

O app completo ainda precisa do Google Apps Script para conversar com a planilha privada. O GitHub Pages sozinho é estático e não consegue gravar com segurança em uma planilha Google privada sem um backend.

Fluxo recomendado com GitHub Pages:

1. O código fica no GitHub.
2. A planilha real fica privada no Google Sheets.
3. O Apps Script é publicado como Web App.
4. O arquivo `docs/index.html` é publicado pelo GitHub Pages.
5. As secretárias acessam o link do GitHub Pages, que abre o app do Apps Script em tela cheia.

Para usar o GitHub Pages, edite `docs/index.html` e substitua:

```text
COLE_AQUI_A_URL_DO_WEB_APP_DO_APPS_SCRIPT
```

pela URL publicada pelo Apps Script.

Abrir `Index.html` direto no navegador serve apenas para ver parte da aparência. O app completo funciona pelo Apps Script, porque depende das funções internas do Google para acessar a planilha.

## Como configurar o Apps Script

Este projeto pode ser usado preferencialmente como script vinculado à planilha. Se for usado como script independente, execute `inicializarPlanilha()` para criar uma planilha de dados automaticamente e salvar o ID nas propriedades do script.

O código de acesso inicial é:

```text
septo123
```

## Como publicar como Web App

1. No Apps Script, clique em `Implantar > Nova implantação`.
2. Escolha o tipo `Aplicativo da Web`.
3. Em `Executar como`, use o proprietário/implantador do projeto.
4. Em `Quem tem acesso`, escolha apenas o grupo ou as pessoas autorizadas.
5. Publique e compartilhe o link apenas com a equipe autorizada.

## Usuários iniciais

- Felipe | Administrador | Sim
- João | Administrador | Sim
- Raquel | Secretária | Sim
- Lizandra | Secretária | Sim
- Livia | Secretária | Sim

Usuários administradores podem acessar a área administrativa para cadastrar hospitais, planos, modelos de WhatsApp, usuários e alterar o código de acesso.

## Como alterar o código de acesso

Entre como Administrador, abra `Administração` e use a seção `Código de acesso`.

Também é possível alterar diretamente na aba `Configurações`, chave `CodigoAcesso`.

## Como usar

1. Acesse o Web App.
2. Informe o usuário autorizado e o código de acesso.
3. Use os cards de resumo e filtros para localizar pacientes.
4. Cadastre novos pacientes no botão `Novo paciente`.
5. Em cada card, use `Editar`, `Mudar status`, `Ver histórico`, `WhatsApp` ou `Arquivar`.

## Cuidados de segurança

- Não coloque dados reais de pacientes no GitHub.
- Mantenha a planilha privada.
- Compartilhe o Web App apenas com pessoas autorizadas.
- Use arquivamento em vez de exclusão definitiva.
- Não armazene exames, PDFs, imagens ou tomografias nesta versão.
- A autenticação desta versão é simplificada para equipe pequena e uso interno. Ela não substitui autenticação profissional com conta individual e senha criptografada.
