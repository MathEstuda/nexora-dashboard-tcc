# Nexora

**Inteligência para dados, processos e relatórios**

A Nexora é um dashboard desenvolvido como parte de um Trabalho de Conclusão de Curso em Engenharia da Computação.

A aplicação processa bases administrativas estruturadas, apresenta indicadores e auxilia na elaboração, revisão, avaliação, validação e exportação de relatórios.

## Aplicação online

Acesse a versão pública da Nexora:

https://nexora-dashboard-tcc-matheus-av4-final-tcc.streamlit.app

## Funcionalidades

- leitura de planilhas Excel;
- base acadêmica simulada;
- importação de planilhas externas;
- normalização inteligente de abas e colunas;
- reconhecimento de variações e sinônimos nos cabeçalhos;
- validação estrutural e de qualidade dos dados;
- alternância entre base acadêmica e planilha importada;
- indicadores, gráficos, ocorrências e alertas;
- verificação de coerência matemática;
- geração padronizada de prompts;
- cópia e download do prompt;
- preservação do rascunho original;
- revisão humana;
- comparação entre texto original e revisado;
- avaliação humana condicional;
- classificação objetiva do relatório;
- validação de fidelidade;
- relatório executivo Word com gráficos e texto;
- exportação em DOCX e TXT.

## Fluxo de uso da inteligência artificial

A interação com a inteligência artificial generativa é realizada manualmente:

1. a Nexora prepara o prompt com os dados do cenário;
2. o usuário copia o prompt e o envia ao Microsoft Copilot;
3. a primeira resposta da IA é inserida na Nexora sem correções;
4. o texto original é preservado;
5. o usuário realiza a revisão humana;
6. a Nexora analisa as alterações, registra a avaliação e valida o relatório;
7. o documento final pode ser exportado em formato executivo ou textual.

A versão atual não utiliza integração direta por API.

## Dados

Todos os dados incluídos neste repositório são fictícios ou simulados para fins acadêmicos.

A base acadêmica apresenta seis cenários administrativos e 210 registros operacionais simulados.

A aplicação também permite importar planilhas externas compatíveis, analisar seus cabeçalhos e confirmar o mapeamento antes da ativação dos dados.

## Execução local

1. Instale as dependências:

    pip install -r requirements.txt

2. Inicie a aplicação:

    streamlit run streamlit_app.py

3. Abra no navegador o endereço informado pelo Streamlit.

## Estrutura do repositório

- `streamlit_app.py`: aplicação principal;
- `requirements.txt`: dependências do projeto;
- `README.md`: documentação;
- `.gitignore`: regras para arquivos não versionados;
- `.streamlit/`: configuração visual da aplicação;
- `data/`: base acadêmica simulada;
- `examples/`: exemplos de relatórios;
- `prompts/`: prompts utilizados nos experimentos;
- `results/`: avaliações e resultados;
- `screenshots/`: capturas e evidências visuais.

## Código-fonte

Repositório público:

https://github.com/MathEstuda/nexora-dashboard-tcc

## Autor

**Matheus Freire de Oliveira Dutra**

Curso de Engenharia da Computação  
Universidade Veiga de Almeida  
Rio de Janeiro, RJ  
2026

## Uso de inteligência artificial

A inteligência artificial foi utilizada como objeto de estudo da aplicação e como apoio à programação, organização, revisão e documentação do projeto.

Todas as sugestões e saídas produzidas com assistência de IA foram submetidas à análise e à revisão do autor.

## Aviso

Este projeto é acadêmico e utiliza dados simulados.

A aplicação não deve ser utilizada para decisões administrativas reais sem procedimentos adicionais de segurança, privacidade, governança, validação e adequação ao ambiente organizacional.

## Licença

Este projeto está disponibilizado sob a licença MIT. Consulte o arquivo `LICENSE`.
