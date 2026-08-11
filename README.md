# GPFO v2

## Gestão e Processamento de Formações Online

O **GPFO** é uma aplicação desenvolvida para apoiar a gestão, processamento e análise de formações realizadas através de plataformas de videoconferência, com especial atenção ao tratamento dos dados de participação e presença dos formandos.

O sistema permite importar dados de sessões de formação a partir de ficheiros Excel, processar as informações dos participantes, armazenar os dados numa base de dados SQLite e gerar relatórios consolidados em Excel.

O projeto encontra-se em desenvolvimento contínuo, com evolução progressiva da sua arquitetura e funcionalidades.

![Python](https://img.shields.io/badge/Python-3.11%2B-blue.svg)
![FastAPI](https://img.shields.io/badge/FastAPI-Framework-009688.svg)
![SQLite](https://img.shields.io/badge/Database-SQLite-003B57.svg)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel-217346.svg)
![Uvicorn](https://img.shields.io/badge/Uvicorn-ASGI-499848.svg)
![OpenAPI](https://img.shields.io/badge/OpenAPI-3.x-6BA539.svg)
![Swagger UI](https://img.shields.io/badge/Swagger%20UI-Documentation-85EA2D.svg)


---

## Objetivo

O principal objetivo do GPFO é automatizar e centralizar o tratamento dos dados resultantes de sessões de formação online.

A aplicação foi concebida para reduzir tarefas manuais relacionadas com:

- Importação de dados de sessões;
- Tratamento dos participantes;
- Registo das presenças;
- Organização das formações;
- Gestão das sessões realizadas;
- Cálculo da duração das participações;
- Registo de reconexões;
- Consulta dos dados armazenados;
- Geração de relatórios Excel.

Desta forma, o sistema procura transformar dados brutos provenientes das sessões de formação em informação consolidada e estruturada, pronta para análise e utilização administrativa.

## Tecnologias utilizadas

- Python
- FastAPI
- Uvicorn
- SQLite
- OpenPyXL
- OpenAPI
- Swagger UI

## Evolução da arquitetura

O projeto iniciou-se como uma aplicação monolítica como versão inicial para treinamento e apresentação do conceito, após a validação positiva me concentro na interface, regras de negócio, acesso à base de dados e processamento dos ficheiros num único programa.

Durante o desenvolvimento, essa estrutura foi migrada para uma arquitetura baseada em **FastAPI**, separando as principais responsabilidades da aplicação. A migração teve como objetivo manter a lógica funcional existente, mas proporcionar uma estrutura mais organizada, modular e preparada para futuras evoluções.

A estrutura atual é organizada aproximadamente da seguinte forma:

GPFO/
│
├── app/
│   │
│   ├── api/
│   │   └── routes.py
│   │
│   ├── services/
│   │   ├── teams.py
│   │   ├── formacoes.py
│   │   └── exportacao.py
│   │
│   ├── config.py
│   ├── database.py
│   ├── logger.py
│   └── main.py
│
├── database/
│
├── uploads/
│
├── exports/
│
├── run.py
├── requirements.txt
├── .gitignore
└── README.md

## 📄 Autoria e Licença

© 2026 Eduardo Oliveira. Este código é disponibilizado apenas para fins de estudo e demonstração de portfólio. É proibida a utilização comercial, redistribuição ou incorporação em outros projetos sem autorização expressa do autor.
