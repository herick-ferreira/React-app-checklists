# 📊 Dashboard React Native - Checklist de Qualidade Exemplo

<img width="1906" height="912" alt="image" src="https://github.com/user-attachments/assets/a48444d8-218d-445b-a6b3-e377455b13f3" />


[Link para o projeto](https://react-native-app-checklists.vercel.app/)

---

## Descrição

Dashboard interativo e responsivo construído com React, Expo e TypeScript. Exibe métricas de desempenho com gráficos (gauge, linha), rankings e tabelas filtráveis. Dados podem ser carregados de um arquivo Excel ou gerados programaticamente.

### Funcionalidades

- **Gauge Chart**: Visualização de métrica geral com meta
- **Line Chart**: Evolução de média por ano e mês
- **Rankings**: Dados ordenados por Loja, Tópico e Tag
- **Tabela Geral**: Listagem completa com paginação
- **Filtros**: Multi-select para Ano, Mês, Loja, Tópico e Tag
- **Responsivo**: Adapta-se a dispositivos móveis e desktop
- **Carregamento de Excel**: Suporta arquivos XLSX com fallback para dados de exemplo

---

## Início Rápido

### 1. Instalar dependências

```bash
npm install --legacy-peer-deps
```

### 2. Iniciar o app

```bash
# Web
npx expo start --web

# Android emulator
npx expo start --android

# iOS simulator
npx expo start --ios
```

### 3. Desenvolvimento

Edite os arquivos em `src/app/` — o projeto usa [file-based routing](https://docs.expo.dev/router/introduction).

---

## Estrutura do Projeto

```
src/
├── app/
│   ├── _layout.tsx          # Layout raiz
│   └── index.tsx            # Dashboard (tela principal)
├── components/              # Componentes reutilizáveis
├── constants/               # Temas e constantes
├── hooks/                   # Custom hooks
└── global.css               # Estilos globais
```

---

## Recursos

- [Expo Docs](https://docs.expo.dev/)
- [React Documentation](https://react.dev/)
- [TypeScript Guide](https://docs.expo.dev/guides/typescript/)
- [Expo Router](https://docs.expo.dev/router/introduction/)
