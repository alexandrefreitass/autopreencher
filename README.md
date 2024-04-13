# Autopreenchimento de Formulários com Python

Este projeto foi desenvolvido para automatizar o preenchimento de formulários em um site administrado no meu trabalho. Utiliza Python juntamente com as bibliotecas `openpyxl`, `pyperclip` e `pyautogui` para preencher automaticamente um formulário a partir de dados extraídos de uma grande planilha de Excel, uma tarefa que demoraria dias para ser realizada manualmente.

## Funcionalidades

- **Leitura de Dados**: Extrai informações de uma planilha Excel usando `openpyxl`.
- **Automação do Teclado e Mouse**: Usa `pyautogui` para simular interações com o teclado e o mouse, preenchendo o formulário no site.
- **Gerenciamento de Clipboard**: Utiliza `pyperclip` para copiar e colar dados no formulário.

## Requisitos

Para executar este script, você precisará instalar algumas dependências. As instruções a seguir assumem que você já possui Python instalado em sua máquina.

### Instalação de Dependências

Abra o terminal e execute o seguinte comando para instalar todas as dependências necessárias:

```bash
pip install openpyxl pyperclip pyautogui
```

### Uso
Prepare sua Planilha: Certifique-se de que a planilha Excel contém os dados no formato esperado pelo script.

### Para obter as coordenadas
Coordenadas do Formulário: Utilize a função mouseinfo() para obter as coordenadas exatas dos campos de entrada no formulário que você deseja automatizar.

Abra PowerShell:
```bash
pip install mouseinfo
python
from mouseinfo import mouseInfo
mouseInfo()
```

### Contribuições
Contribuições para melhorar este projeto são sempre bem-vindas. Sinta-se à vontade para clonar este repositório, fazer suas modificações e submeter um pull request.

### Licença
Este projeto está sob a licença MIT. Consulte o arquivo <a href="https://github.com/alexandrefreitass/autopreencher/blob/master/LICENSE.txt">LICENSE</a> para obter mais detalhes.