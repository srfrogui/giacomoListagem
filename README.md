# Gerador de Relatório em PDF

Este projeto é um script Python que gera um relatório em PDF a partir de um arquivo Excel. O relatório inclui detalhes sobre peças e pode ser formatado para incluir informações específicas, como dimensões e cliente.

## Requisitos

Certifique-se de que você tem Python 3.x instalado. O projeto utiliza as seguintes bibliotecas:

- `pandas`
- `reportlab`
- `tkinter`

Você pode instalar as dependências necessárias executando:

```bash
pip install -r req.txt
```

## Uso

1. Execute o script Python:

    ```bash
    python amostrado.py
    ```

2. Uma janela de seleção de arquivo será aberta. Selecione o arquivo Excel que contém os dados das peças. O arquivo deve estar no formato `.xls` ou `.xlsx`.

3. O script solicitará que você escolha o relatório que deseja gerar. Atualmente, você pode gerar um "Relatório de Peças".

4. O relatório será salvo como `Relatorio_Pecas.pdf` no diretório atual.

## Funcionalidades

- **Formatação de Valores**: Remove o `.0` de valores flutuantes e formata números inteiros.
- **Geração de Relatório**: Cria um PDF com tabelas formatadas contendo informações sobre peças.
- **Personalização**: Inclui opções de formatação como alinhamento de texto, cores e tamanhos de fonte.

## Personalização

Você pode ajustar o estilo do relatório editando o script, incluindo:

- **Largura das Colunas**: Modifique as larguras das colunas no código.
- **Estilos da Tabela**: Altere a aparência das tabelas, como cor de fundo e estilo de fonte.

## Problemas Conhecidos

- **Altura das Linhas**: Pode haver problemas com a altura das linhas em alguns casos, ajuste manualmente se necessário.

## Contribuições

Contribuições são bem-vindas! Se você encontrar bugs ou tiver sugestões de melhorias, sinta-se à vontade para abrir uma issue ou enviar um pull request.

## Licença

Este projeto é licenciado sob a [Licença MIT](LICENSE).
