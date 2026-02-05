📸 Automatizador de Relatórios Fotográficos - IML/AL

Este projeto foi desenvolvido como uma solução prática para a Polícia Científica de Alagoas, atendendo a uma solicitação do Diretor Geral do IML. O objetivo principal é otimizar o fluxo de trabalho dos médicos legistas, transformando pastas de evidências fotográficas em relatórios PDF organizados e padronizados.
📝 O Problema

Anteriormente, o compartilhamento de imagens de casos com outros órgãos era feito através do envio de pastas completas. Isso gerava:

    Ineficiência Logística: Dificuldade no manuseio e visualização ordenada das fotos.

    Desperdício de Armazenamento: Como o IML não possuía um storage específico para esse fim na época, o acúmulo de arquivos brutos localmente era insustentável.

    Trabalho Manual: Médicos gastavam tempo excessivo organizando arquivos para consultas externas.

🚀 A Solução

O software permite que, com poucos cliques, o usuário selecione uma pasta de fotos e um modelo de documento, gerando um PDF final com 6 fotos por página (configurável), legendado com o nome do caso e formatado para impressão ou envio digital.
Principais Funcionalidades:

    Interface Gráfica (GUI): Desenvolvida para usuários que não possuem familiaridade com linha de comando.

    Processamento Inteligente: Redimensionamento e corte automático das imagens para ajuste perfeito no template PDF.

    Portabilidade: Compilado como um executável de arquivo único (.exe), eliminando a necessidade de instalação de Python ou dependências na máquina do usuário.

    Multithreading: A interface permanece responsiva durante o processamento pesado de imagens.