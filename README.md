# Extração de dados de documentos web via OCR

Lê links de uma mensagem copiada, acessa cada página, extrai o conteúdo dos documentos por reconhecimento de imagem e consolida tudo em um único arquivo Word.

---

## Contexto

Na hora de optar por um material era preciso acessar várias páginas, com várias informações técnicas, compará-las e avaliar as melhores propriedades de cada um, um processo lento e bem suscetível a erros.

Este programa encurta e agiliza esse caminho: um bot web acessa as páginas, lê o conteúdo, filtra e resume as informações por meio da API do Gemini com o prompt desejado e formata o resultado em um `.docx`, agilizando o processo e permitindo uma comparação mais rápida entre materiais, contudo a comparação e escolha final ainda são feitas por mãos humanas.

---

## Como funciona

```
Lê a mensagem da área de transferência
        ↓
Filtra os links presentes na mensagem
        ↓
Bot acessa os links, sequencialmente
        ↓
Bot caminha até a pagina em que está as informações
        ↓
Captura a tela da página
        ↓
Pré-processa a imagem para melhorar o reconhecimento
        ↓
Tesseract (OCR) converte a imagem em texto
        ↓
Gemini estrutura o resultado com um prompt predefinido
        ↓
Gera documento Word (.docx) consolidado
```

---

## Decisões técnicas

### Prints e uso de OCR

O documento com as especificações era renderizado como uma iamgem na página, não havia arquivo para baixar e processar com uma biblioteca de leitura específica do formato e o texto não era um elemento no DOM da pagina. Assim tive que capturar a tela e usar reconhecimento de imagem, já a própria pagina tratava as informacões como imagem, não como um objeto textual no DOM ou como arquivo.

### Leitura dos links pela área de transferência

Os links chegavam por mensagem, misturados a texto solto. Em vez de obrigar o usuário a copiar a mensagem na linha de comando, algo que para leigos pode gerar confusões, já que, dependendo das configurações, ctrl + v não cola. Assim, optei por usar a biblioteca Pyperclip para ler direto da área de transferência, não sendo necessário colar o texto na linha de comando.

### Desativar a suspensão do Windows

Por o programa ficar rodando por um tempo considerável, é comum o computador entrar em suspensão e interromper a execução. Por isso mantem - se uma chamada à API do sistema que impede a suspensão enquanto o processamento está em andamento, liberando-a novamente ao final.

### Verificação de carregamento da página

Dependendo da velocidade da máquina e da rede, o carregamento das páginas pode demorar mais ou menos. Por isso o programa faz paradas e verificações durante a execução para garantir que a página foi carregada antes de prosseguir, em vez de confiar em um tempo de espera fixo.

### A captura cortava o documento

Quando o bot apenas capturava a tela sem checar a altura, o texto ficava cortado e parte das especificações se perdia. Foi necessário extrair o maior parâmetro de altura da página, adicionar uma margem e forçar a janela a esse novo número de pixels antes da captura, garantindo que a imagem cobrisse o documento inteiro.

### O OCR não funcionava com a imagem crua

O Tesseract é treinado com documentos escaneados, ou seja, em preto e branco e em alta resolução, desse modo errava muito com uma captura de tela em Full HD.

Primeiro removi a cor, convertendo o RGB para uma escala de cinza de 8 bits (0 a 255). Contudo, a imagem mantinha sombreamentos e gradientes que atrapalhavam o reconhecimento das letras.

Como o Tesseract é treinado com alta resolução, o upscaling era uma solução natural, mas há várias técnicas para preencher os pixels criados ao aumentar a escala. Como o motor trabalha com preto e branco, ele é privilegiado por bordas bem definidas, que é exatamente onde a técnica de **Lanczos** se destaca, ja que ela realça o contraste nas transições abruptas, como nas bordas das letras, em vez de suavizá-las, deixando o contorno das letras mais nítido.

Por fim, o Lanczos melhora a definição das bordas, mas não elimina os sombreamentos e gradientes de fundo. Para isso binarizei a imagem, definindo que até 160 é preto e acima é branco.

---

## Stack

- **Python 3**
- **Selenium webdriver-manager** — bot de navegação web e captura de tela
- **Pillow** — pré-processamento da imagem
- **pytesseract** — interface com o Tesseract OCR
- **google-generativeai** — estruturação do texto extraído via LLM
- **python-docx** — geração do documento final em formato `.docx`
- **pyperclip** — leitura da área de transferência

---

## Como executar

### Pré-requisitos

- **Python 3.8+**
- **Google Chrome** instalado (o webdriver é baixado automaticamente)
- **Tesseract OCR** com o pacote de idioma português (`por`)
  - Windows: [instalador oficial](https://github.com/UB-Mannheim/tesseract/wiki) — marque "Portuguese" na seleção de idiomas
- **Chave de API** do Google AI Studio (usa funçoes específicas da biblioteca do Google Ai Studio)
- Sistema **Windows** (o controle de suspensão e a abertura automatica do arquivo usam APIs específicas do Windows)

### Instalação

```bash
git clone https://github.com/SEU-USUARIO/extracao-ocr-documentos.git
cd extracao-ocr-documentos
pip install -r requirements.txt
```

### Configuração

Abra `coleta_links.py` e preencha:

```python
GENAI_API_KEY = "sua-chave-aqui"

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

Ajuste também, conforme o site alvo:

- `padrao` em `extrair_links_do_texto()` — formato dos links a serem procurados na mensagem
- o texto do botão em `processar_url_individual()` — texto no HTML do botão o qual o pode tera que clicar
- `prompt_mestre` em `formatar_com_ia()` — prompt com o qual o Gemini processara as informaçoes extraidas

### Execução

```bash
python coleta_links.py
```

Copie a mensagem contendo os links (Ctrl+C) e pressione Enter no programa. O arquivo consolidado é salvo em `DOCUMENTOS_PROCESSADOS/` e aberto automaticamente ao final.

---

## Limitações conhecidas

- **Depende do layout do site alvo:** os seletores foram escritos para uma estrutura específica, mudanças na página exigem ajuste.
- **Windows apenas.** o controle de suspensão e a abertura do arquivo final usam APIs nativas do Windows.
- **Limiar de binarização fixo:** funciona para imagens de contraste consistente, como, nesse caso, prints sempre de uma mesma pagina com o mesmo formato, para imagens de padrões variadas esse limiar pode não garantir uma boa precisão, sendo necessário testar e achar um ideal manualmente ou aplicar diretamente Otsu, um metodo que analisa a frequencia de cada cor (já em preto e branco) e acha o limiar que maximiza o contraste através de estatística, contudo a bliblioteca Pillow não possui o metodo de Otsu nativamente.
- **Não gera relatorios de erros:** erros em alguma das etapas do processamento não são detectados (exceto em uma parte em que o codigo verifica se há mais de 50 palavras no texto gerado pelo Tesseract, que é mais para verificar se ele leu alguma coisa, do que realmente para verificar se ele leu certo), logo saem direto no documento final. Poderia - se implementar algumas verificações em 2 etapas, usando mais um software de OCR em paralelo e mais uma IA para comparar as imagens geradas, porém como é uma aplicação simples e não critíca, julgo que seja um desperdício de recursos.
- **Sem processamento paralelo:** os links são processados sequencialmente, então a execução pode demorar, de modo geral, como não havia muita limitação de tempo, a velocidade de execução não foi muito otimizada, tanto é que há varios momentos em que o pode para e espera, que poderiam ser testados e encurtados, caso necessário, por exemplo.

